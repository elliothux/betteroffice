#![deny(clippy::all)]

use std::collections::BTreeMap;
use std::sync::mpsc::{self, Sender};

use betteroffice_pptx::{
    Background, CommentFlavor, EditCtx, EditOrigin, MAX_COLLABORATION_CLIENT_ID,
    Presentation as CorePresentation, RenderOptions,
};
use napi::bindgen_prelude::{AsyncTask, Buffer, Task};
use napi::{Env, Error, Result};
use napi_derive::napi;
use serde::Serialize;

fn error(reason: impl ToString) -> Error {
    Error::from_reason(reason.to_string())
}

type JobResult = std::result::Result<Response, String>;
type Job = Box<dyn FnOnce(&mut CorePresentation) -> JobResult + Send>;

enum Response {
    Json(serde_json::Value),
    Bytes(Vec<u8>),
    Media(Vec<MediaData>),
    Render(RenderData),
}

struct Request {
    job: Job,
    reply: Sender<JobResult>,
}

#[derive(Clone)]
pub struct Worker {
    sender: Sender<Request>,
}

impl Worker {
    fn call(&self, job: Job) -> Result<Response> {
        let (reply, receive) = mpsc::channel();
        self.sender
            .send(Request { job, reply })
            .map_err(|_| error("presentation worker stopped"))?;
        receive
            .recv()
            .map_err(|_| error("presentation worker stopped"))?
            .map_err(error)
    }

    fn json(&self, job: Job) -> Result<serde_json::Value> {
        match self.call(job)? {
            Response::Json(value) => Ok(value),
            _ => Err(error("presentation worker returned an invalid response")),
        }
    }
}

fn json<T: Serialize>(value: T) -> JobResult {
    serde_json::to_value(value)
        .map(Response::Json)
        .map_err(|value| value.to_string())
}

fn parse<T: serde::de::DeserializeOwned>(value: serde_json::Value) -> Result<T> {
    serde_json::from_value(value).map_err(error)
}

fn client_id(value: f64) -> Result<u64> {
    if !value.is_finite()
        || value.fract() != 0.0
        || value < 1.0
        || value > MAX_COLLABORATION_CLIENT_ID as f64
    {
        return Err(error("clientId must be a positive safe integer"));
    }
    Ok(value as u64)
}

fn origin(value: &str) -> Result<EditOrigin> {
    match value {
        "local" => Ok(EditOrigin::Local),
        "agent" => Ok(EditOrigin::Agent),
        "remote" => Ok(EditOrigin::Remote),
        "system" => Ok(EditOrigin::System),
        _ => Err(error("origin must be local, agent, remote, or system")),
    }
}

fn origin_name(value: EditOrigin) -> &'static str {
    match value {
        EditOrigin::Local => "local",
        EditOrigin::Agent => "agent",
        EditOrigin::Remote => "remote",
        EditOrigin::System => "system",
    }
}

fn comment_flavor(value: &str) -> Result<CommentFlavor> {
    match value {
        "legacy" => Ok(CommentFlavor::Legacy),
        "modern" => Ok(CommentFlavor::Modern),
        _ => Err(error("commentFlavor must be legacy or modern")),
    }
}

fn comment_flavor_name(value: CommentFlavor) -> &'static str {
    match value {
        CommentFlavor::Legacy => "legacy",
        CommentFlavor::Modern => "modern",
    }
}

#[napi(object)]
pub struct OpenPresentationOptions {
    pub client_id: Option<f64>,
    pub author: Option<String>,
    pub origin: Option<String>,
}

#[napi(object)]
pub struct FontFace {
    pub family: String,
    pub data: Buffer,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
}

#[napi(object)]
pub struct RenderSlideOptions {
    pub scale: Option<f64>,
    pub transparent: Option<bool>,
    pub background: Option<String>,
    pub max_shadow_pixels: Option<f64>,
}

#[napi(object)]
pub struct RenderedSlide {
    pub data: Buffer,
    pub width: u32,
    pub height: u32,
    pub skipped_images: u32,
}

#[napi(object)]
pub struct MediaResource {
    pub path: String,
    pub content_type: String,
    pub data: Buffer,
}

struct MediaData {
    path: String,
    content_type: String,
    bytes: Vec<u8>,
}

#[napi(object)]
pub struct CommentInput {
    pub slide_id: String,
    pub text: String,
    pub author: String,
    pub initials: Option<String>,
    pub created: String,
    pub x: Option<i64>,
    pub y: Option<i64>,
}

#[napi(object)]
pub struct CommentReplyInput {
    pub comment_id: String,
    pub text: String,
    pub author: String,
    pub initials: Option<String>,
    pub created: String,
}

pub struct RenderData {
    bytes: Vec<u8>,
    width: u32,
    height: u32,
    skipped_images: usize,
}

pub struct OpenTask {
    bytes: Vec<u8>,
    client_id: Option<u64>,
    author: String,
    origin: EditOrigin,
}

impl Task for OpenTask {
    type Output = Worker;
    type JsValue = PptxPresentation;

    fn compute(&mut self) -> Result<Self::Output> {
        let bytes = std::mem::take(&mut self.bytes);
        let client_id = self.client_id;
        let (sender, receive) = mpsc::channel::<Request>();
        let (ready, opened) = mpsc::channel();
        std::thread::Builder::new()
            .name("betteroffice-pptx".to_owned())
            .spawn(move || {
                let presentation = match client_id {
                    Some(client_id) => CorePresentation::open_collaborative(&bytes, client_id),
                    None => CorePresentation::open(&bytes),
                };
                let mut presentation = match presentation {
                    Ok(value) => {
                        let _ = ready.send(Ok(()));
                        value
                    }
                    Err(value) => {
                        let _ = ready.send(Err(value.to_string()));
                        return;
                    }
                };
                while let Ok(request) = receive.recv() {
                    let result = (request.job)(&mut presentation);
                    let _ = request.reply.send(result);
                }
            })
            .map_err(error)?;
        opened
            .recv()
            .map_err(|_| error("presentation worker failed to start"))?
            .map_err(error)?;
        Ok(Worker { sender })
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(PptxPresentation {
            worker: output,
            author: self.author.clone(),
            origin: self.origin,
            collaborative: self.client_id.is_some(),
        })
    }
}

pub struct RenderTask {
    worker: Worker,
    slide: usize,
    options: RenderOptions,
}

impl Task for RenderTask {
    type Output = RenderData;
    type JsValue = RenderedSlide;

    fn compute(&mut self) -> Result<Self::Output> {
        let slide = self.slide;
        let options = self.options.clone();
        match self.worker.call(Box::new(move |presentation| {
            presentation
                .render_png(slide, &options)
                .map(|value| {
                    Response::Render(RenderData {
                        bytes: value.bytes,
                        width: value.width,
                        height: value.height,
                        skipped_images: value.skipped_images,
                    })
                })
                .map_err(|value| value.to_string())
        }))? {
            Response::Render(value) => Ok(value),
            _ => Err(error("presentation worker returned an invalid response")),
        }
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(RenderedSlide {
            data: output.bytes.into(),
            width: output.width,
            height: output.height,
            skipped_images: output.skipped_images as u32,
        })
    }
}

pub struct SaveTask {
    worker: Worker,
}

impl Task for SaveTask {
    type Output = Vec<u8>;
    type JsValue = Buffer;

    fn compute(&mut self) -> Result<Self::Output> {
        match self.worker.call(Box::new(|presentation| {
            presentation
                .save()
                .map(Response::Bytes)
                .map_err(|value| value.to_string())
        }))? {
            Response::Bytes(value) => Ok(value),
            _ => Err(error("presentation worker returned an invalid response")),
        }
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output.into())
    }
}

#[napi(js_name = "Presentation")]
pub struct PptxPresentation {
    worker: Worker,
    author: String,
    origin: EditOrigin,
    collaborative: bool,
}

impl PptxPresentation {
    fn context(&self) -> EditCtx {
        EditCtx {
            origin: self.origin,
            author: self.author.clone(),
        }
    }

    fn require_collaborative(&self) -> Result<()> {
        if self.collaborative {
            Ok(())
        } else {
            Err(error("operation requires a collaborative presentation"))
        }
    }
}

#[napi]
impl PptxPresentation {
    #[napi(getter)]
    pub fn author(&self) -> String {
        self.author.clone()
    }

    #[napi(setter)]
    pub fn set_author(&mut self, author: String) {
        self.author = author;
    }

    #[napi(getter)]
    pub fn origin(&self) -> &'static str {
        origin_name(self.origin)
    }

    #[napi(setter)]
    pub fn set_origin(&mut self, value: String) -> Result<()> {
        self.origin = origin(&value)?;
        Ok(())
    }

    #[napi(getter)]
    pub fn collaborative(&self) -> bool {
        self.collaborative
    }

    #[napi(getter)]
    pub fn client_id(&self) -> Result<f64> {
        self.worker
            .json(Box::new(|presentation| json(presentation.client_id())))
            .and_then(|value| value.as_f64().ok_or_else(|| error("invalid client ID")))
    }

    #[napi]
    pub fn snapshot(&self) -> Result<serde_json::Value> {
        self.worker.json(Box::new(|presentation| {
            presentation
                .snapshot()
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi(getter)]
    pub fn slide_count(&self) -> Result<u32> {
        self.worker
            .json(Box::new(|presentation| {
                presentation
                    .snapshot()
                    .map(|snapshot| snapshot.slides.len() as u32)
                    .map_err(|value| value.to_string())
                    .and_then(json)
            }))?
            .as_u64()
            .map(|value| value as u32)
            .ok_or_else(|| error("invalid slide count"))
    }

    #[napi(getter)]
    pub fn slide_ids(&self) -> Result<Vec<String>> {
        let value = self.worker.json(Box::new(|presentation| {
            presentation
                .snapshot()
                .map(|snapshot| {
                    snapshot
                        .slides
                        .into_iter()
                        .map(|slide| slide.id)
                        .collect::<Vec<_>>()
                })
                .map_err(|value| value.to_string())
                .and_then(json)
        }))?;
        parse(value)
    }

    #[napi(getter)]
    pub fn width_emu(&self) -> Result<i64> {
        self.worker
            .json(Box::new(|presentation| {
                presentation
                    .snapshot()
                    .map(|snapshot| snapshot.width_emu)
                    .map_err(|value| value.to_string())
                    .and_then(json)
            }))?
            .as_i64()
            .ok_or_else(|| error("invalid presentation width"))
    }

    #[napi(getter)]
    pub fn height_emu(&self) -> Result<i64> {
        self.worker
            .json(Box::new(|presentation| {
                presentation
                    .snapshot()
                    .map(|snapshot| snapshot.height_emu)
                    .map_err(|value| value.to_string())
                    .and_then(json)
            }))?
            .as_i64()
            .ok_or_else(|| error("invalid presentation height"))
    }

    #[napi]
    pub fn slide(&self, slide: u32) -> Result<serde_json::Value> {
        self.worker.json(Box::new(move |presentation| {
            let snapshot = presentation.snapshot().map_err(|value| value.to_string())?;
            snapshot
                .slides
                .get(slide as usize)
                .ok_or_else(|| format!("slide index {slide} is out of range"))
                .and_then(json)
        }))
    }

    #[napi]
    pub fn story(&self, story_id: String) -> Result<serde_json::Value> {
        self.worker.json(Box::new(move |presentation| {
            presentation
                .story(&story_id)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn layouts(&self) -> Result<serde_json::Value> {
        self.worker.json(Box::new(|presentation| {
            json(
                presentation
                    .layouts()
                    .iter()
                    .map(|layout| layout.part_path.clone())
                    .collect::<Vec<_>>(),
            )
        }))
    }

    #[napi]
    pub fn media(&self) -> Result<Vec<MediaResource>> {
        match self.worker.call(Box::new(|presentation| {
            let value = presentation
                .media()
                .iter()
                .map(|part| MediaData {
                    path: part.part_path.clone(),
                    content_type: part.content_type.clone(),
                    bytes: part.bytes.clone(),
                })
                .collect::<Vec<_>>();
            Ok(Response::Media(value))
        }))? {
            Response::Media(value) => Ok(value
                .into_iter()
                .map(|part| MediaResource {
                    path: part.path,
                    content_type: part.content_type,
                    data: part.bytes.into(),
                })
                .collect()),
            _ => Err(error("presentation worker returned an invalid response")),
        }
    }

    #[napi]
    pub fn insert_slide(&self, index: u32, layout: Option<String>) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .insert_slide(&context, index, layout.as_deref())
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn delete_slide(&self, slide_id: String) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .delete_slide(&context, &slide_id)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn move_slide(&self, slide_id: String, index: u32) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .move_slide(&context, &slide_id, index)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn set_slide_notes(&self, slide_id: String, text: String) -> Result<()> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .set_slide_notes(&context, &slide_id, &text)
                .map_err(|value| value.to_string())?;
            json(true)
        }))?;
        Ok(())
    }

    #[napi]
    pub fn add_text_box(
        &self,
        slide_id: String,
        draft: serde_json::Value,
    ) -> Result<serde_json::Value> {
        let context = self.context();
        let draft = parse(draft)?;
        self.worker.json(Box::new(move |presentation| {
            presentation
                .add_text_box(&context, &slide_id, &draft)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn add_shape(
        &self,
        slide_id: String,
        draft: serde_json::Value,
    ) -> Result<serde_json::Value> {
        let context = self.context();
        let draft = parse(draft)?;
        self.worker.json(Box::new(move |presentation| {
            presentation
                .add_shape(&context, &slide_id, &draft)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn remove_shape(&self, slide_id: String, shape_id: String) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .remove_shape(&context, &slide_id, &shape_id)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn set_shape_fill(
        &self,
        slide_id: String,
        shape_id: String,
        color: Option<String>,
    ) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .set_shape_fill(&context, &slide_id, &shape_id, color.as_deref())
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn set_shape_stroke(
        &self,
        slide_id: String,
        shape_id: String,
        stroke: serde_json::Value,
    ) -> Result<serde_json::Value> {
        let context = self.context();
        let stroke = parse(stroke)?;
        self.worker.json(Box::new(move |presentation| {
            presentation
                .set_shape_stroke(&context, &slide_id, &shape_id, &stroke)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn set_shape_adjust(
        &self,
        slide_id: String,
        shape_id: String,
        adjustments: serde_json::Value,
    ) -> Result<serde_json::Value> {
        let context = self.context();
        let adjustments: BTreeMap<String, f64> = parse(adjustments)?;
        self.worker.json(Box::new(move |presentation| {
            presentation
                .set_shape_adjust(&context, &slide_id, &shape_id, &adjustments)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn move_shape(
        &self,
        slide_id: String,
        shape_id: String,
        x: i64,
        y: i64,
    ) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .move_shape(&context, &slide_id, &shape_id, x, y)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn resize_shape(
        &self,
        slide_id: String,
        shape_id: String,
        width: i64,
        height: i64,
    ) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .resize_shape(&context, &slide_id, &shape_id, width, height)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn set_shape_rect(
        &self,
        slide_id: String,
        shape_id: String,
        rect: serde_json::Value,
    ) -> Result<serde_json::Value> {
        let context = self.context();
        let rect = parse(rect)?;
        self.worker.json(Box::new(move |presentation| {
            presentation
                .set_shape_rect(&context, &slide_id, &shape_id, rect)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn insert_text(
        &self,
        story_id: String,
        index: u32,
        text: String,
        style: Option<serde_json::Value>,
    ) -> Result<serde_json::Value> {
        let context = self.context();
        let style = style.map(parse).transpose()?.unwrap_or_default();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .insert_text(&context, &story_id, index, &text, &style)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn delete_text(&self, story_id: String, start: u32, end: u32) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .delete_text(&context, &story_id, start, end)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn format_text(
        &self,
        story_id: String,
        start: u32,
        end: u32,
        patch: serde_json::Value,
    ) -> Result<serde_json::Value> {
        let context = self.context();
        let patch = parse(patch)?;
        self.worker.json(Box::new(move |presentation| {
            presentation
                .format_text(&context, &story_id, start, end, &patch)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn set_paragraph_alignment(
        &self,
        story_id: String,
        start: u32,
        end: u32,
        alignment: Option<String>,
    ) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .set_paragraph_alignment(&context, &story_id, start, end, alignment.as_deref())
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn insert_paragraph_break(
        &self,
        story_id: String,
        index: u32,
    ) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .insert_paragraph_break(&context, &story_id, index)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn delete_paragraph_break(
        &self,
        story_id: String,
        index: u32,
    ) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .delete_paragraph_break(&context, &story_id, index)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn add_comment(&self, comment: CommentInput) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .add_comment(
                    &context,
                    &comment.slide_id,
                    &comment.author,
                    comment.initials.as_deref().unwrap_or_default(),
                    &comment.text,
                    &comment.created,
                    comment.x.unwrap_or(0),
                    comment.y.unwrap_or(0),
                )
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn reply_to_comment(&self, reply: CommentReplyInput) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .reply_to_comment(
                    &context,
                    &reply.comment_id,
                    &reply.author,
                    reply.initials.as_deref().unwrap_or_default(),
                    &reply.text,
                    &reply.created,
                )
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn set_comment_status(
        &self,
        comment_id: String,
        resolved: Option<bool>,
    ) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .set_comment_status(&context, &comment_id, resolved.unwrap_or(true))
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn remove_comment(&self, comment_id: String) -> Result<serde_json::Value> {
        let context = self.context();
        self.worker.json(Box::new(move |presentation| {
            presentation
                .remove_comment(&context, &comment_id)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi(getter)]
    pub fn comments(&self) -> Result<serde_json::Value> {
        self.worker.json(Box::new(|presentation| {
            presentation
                .comments()
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi(getter)]
    pub fn comment_flavor(&self) -> Result<String> {
        let value = self.worker.json(Box::new(|presentation| {
            presentation
                .comment_flavor()
                .map(comment_flavor_name)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))?;
        value
            .as_str()
            .map(ToOwned::to_owned)
            .ok_or_else(|| error("invalid comment flavor"))
    }

    #[napi]
    pub fn set_comment_flavor(&self, flavor: String) -> Result<String> {
        let context = self.context();
        let flavor = comment_flavor(&flavor)?;
        let value = self.worker.json(Box::new(move |presentation| {
            presentation
                .set_comment_flavor(&context, flavor)
                .map(comment_flavor_name)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))?;
        value
            .as_str()
            .map(ToOwned::to_owned)
            .ok_or_else(|| error("invalid comment flavor"))
    }

    #[napi]
    pub fn register_font(&self, face: FontFace) -> Result<u32> {
        self.worker
            .json(Box::new(move |presentation| {
                presentation
                    .register_font(
                        &face.family,
                        face.bold.unwrap_or(false),
                        face.italic.unwrap_or(false),
                        &face.data,
                    )
                    .map_err(|value| value.to_string())
                    .and_then(json)
            }))?
            .as_u64()
            .map(|value| value as u32)
            .ok_or_else(|| error("invalid font ID"))
    }

    #[napi(ts_return_type = "Promise<RenderedSlide>")]
    pub fn render_slide(
        &self,
        slide: u32,
        options: Option<RenderSlideOptions>,
    ) -> Result<AsyncTask<RenderTask>> {
        let options = options.unwrap_or(RenderSlideOptions {
            scale: None,
            transparent: None,
            background: None,
            max_shadow_pixels: None,
        });
        let background = if options.transparent.unwrap_or(false) {
            Background::Transparent
        } else if let Some(color) = options.background {
            Background::Color(color)
        } else {
            Background::Slide
        };
        let defaults = RenderOptions::default();
        Ok(AsyncTask::new(RenderTask {
            worker: self.worker.clone(),
            slide: slide as usize,
            options: RenderOptions {
                scale: options.scale.unwrap_or(1.0) as f32,
                background,
                max_shadow_pixels: options
                    .max_shadow_pixels
                    .map(|value| value as u64)
                    .unwrap_or(defaults.max_shadow_pixels),
            },
        }))
    }

    #[napi]
    pub fn encode_state_vector(&self) -> Result<Buffer> {
        self.require_collaborative()?;
        match self.worker.call(Box::new(|presentation| {
            Ok(Response::Bytes(presentation.encode_state_vector_v1()))
        }))? {
            Response::Bytes(value) => Ok(value.into()),
            _ => Err(error("presentation worker returned an invalid response")),
        }
    }

    #[napi]
    pub fn encode_state_as_update(&self) -> Result<Buffer> {
        self.require_collaborative()?;
        match self.worker.call(Box::new(|presentation| {
            Ok(Response::Bytes(presentation.encode_state_as_update_v1()))
        }))? {
            Response::Bytes(value) => Ok(value.into()),
            _ => Err(error("presentation worker returned an invalid response")),
        }
    }

    #[napi]
    pub fn encode_diff(&self, state_vector: Buffer) -> Result<Buffer> {
        self.require_collaborative()?;
        match self.worker.call(Box::new(move |presentation| {
            presentation
                .encode_diff_v1(&state_vector)
                .map(Response::Bytes)
                .map_err(|value| value.to_string())
        }))? {
            Response::Bytes(value) => Ok(value.into()),
            _ => Err(error("presentation worker returned an invalid response")),
        }
    }

    #[napi]
    pub fn apply_update(&self, update: Buffer) -> Result<serde_json::Value> {
        self.require_collaborative()?;
        self.worker.json(Box::new(move |presentation| {
            presentation
                .apply_update_v1(&update)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn propose(&self, request: serde_json::Value) -> Result<serde_json::Value> {
        let request = parse(request)?;
        self.worker.json(Box::new(move |presentation| {
            presentation
                .propose(request)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi(getter)]
    pub fn proposals(&self) -> Result<serde_json::Value> {
        self.worker.json(Box::new(|presentation| {
            presentation
                .proposals()
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn preview_proposal(&self, proposal_id: String) -> Result<serde_json::Value> {
        self.worker.json(Box::new(move |presentation| {
            presentation
                .preview_proposal(&proposal_id)
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn accept_proposal(
        &self,
        proposal_id: String,
        force: Option<bool>,
    ) -> Result<serde_json::Value> {
        self.worker.json(Box::new(move |presentation| {
            presentation
                .accept_proposal(&proposal_id, force.unwrap_or(false))
                .map_err(|value| value.to_string())
                .and_then(json)
        }))
    }

    #[napi]
    pub fn reject_proposal(&self, proposal_id: String) -> Result<bool> {
        self.worker
            .json(Box::new(move |presentation| {
                json(presentation.reject_proposal(&proposal_id))
            }))?
            .as_bool()
            .ok_or_else(|| error("invalid proposal response"))
    }

    #[napi(getter)]
    pub fn can_undo(&self) -> Result<bool> {
        self.worker
            .json(Box::new(|presentation| json(presentation.can_undo())))?
            .as_bool()
            .ok_or_else(|| error("invalid undo state"))
    }

    #[napi(getter)]
    pub fn can_redo(&self) -> Result<bool> {
        self.worker
            .json(Box::new(|presentation| json(presentation.can_redo())))?
            .as_bool()
            .ok_or_else(|| error("invalid redo state"))
    }

    #[napi]
    pub fn undo(&self) -> Result<bool> {
        self.worker
            .json(Box::new(|presentation| json(presentation.undo())))?
            .as_bool()
            .ok_or_else(|| error("invalid undo response"))
    }

    #[napi]
    pub fn redo(&self) -> Result<bool> {
        self.worker
            .json(Box::new(|presentation| json(presentation.redo())))?
            .as_bool()
            .ok_or_else(|| error("invalid redo response"))
    }

    #[napi]
    pub fn add_undo_barrier(&self) -> Result<()> {
        self.worker.json(Box::new(|presentation| {
            presentation.add_undo_barrier();
            json(true)
        }))?;
        Ok(())
    }

    #[napi(ts_return_type = "Promise<Buffer>")]
    pub fn save(&self) -> AsyncTask<SaveTask> {
        AsyncTask::new(SaveTask {
            worker: self.worker.clone(),
        })
    }
}

#[napi(ts_return_type = "Promise<Presentation>")]
pub fn open_presentation(
    data: Buffer,
    options: Option<OpenPresentationOptions>,
) -> Result<AsyncTask<OpenTask>> {
    let options = options.unwrap_or(OpenPresentationOptions {
        client_id: None,
        author: None,
        origin: None,
    });
    Ok(AsyncTask::new(OpenTask {
        bytes: data.to_vec(),
        client_id: options.client_id.map(client_id).transpose()?,
        author: options.author.unwrap_or_else(|| "node".to_owned()),
        origin: options
            .origin
            .as_deref()
            .map(origin)
            .transpose()?
            .unwrap_or_default(),
    }))
}
