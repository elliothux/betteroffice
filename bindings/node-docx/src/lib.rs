#![deny(clippy::all)]

use std::sync::{Arc, Mutex};

use betteroffice_docx::{
    Document, EditCtx, EditOrigin, ImageScope, LayoutInput, NoteKind, SaveOptions,
};
use napi::bindgen_prelude::{AsyncTask, Buffer, Task};
use napi::{Env, Error, Result};
use napi_derive::napi;

fn error(reason: impl ToString) -> Error {
    Error::from_reason(reason.to_string())
}

fn lock<T>(value: &Mutex<T>) -> Result<std::sync::MutexGuard<'_, T>> {
    value.lock().map_err(|_| error("document lock is poisoned"))
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

#[napi(object)]
pub struct OpenDocumentOptions {
    pub author: Option<String>,
    pub origin: Option<String>,
    pub timestamp: Option<String>,
}

#[napi(object)]
pub struct FontFace {
    pub family: String,
    pub data: Buffer,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
}

#[napi(object)]
pub struct ImageResource {
    pub relationship_id: String,
    pub data: Buffer,
    pub scope: Option<String>,
    pub part: Option<String>,
}

#[napi(object)]
pub struct SaveDocumentOptions {
    pub timestamp: Option<String>,
    pub update_modified_date: Option<bool>,
    pub modified_by: Option<String>,
}

#[napi(object)]
pub struct RenderedPage {
    pub data: Buffer,
    pub skipped_images: u32,
}

#[napi(object)]
pub struct DocumentStructure {
    pub body_paragraphs: u32,
    pub body_tables: u32,
    pub sections: u32,
    pub headers: u32,
    pub footers: u32,
    pub footnotes: u32,
    pub endnotes: u32,
}

#[napi(object)]
pub struct EditReceipt {
    pub paragraph_id: Option<String>,
    pub story: Option<String>,
    pub start: Option<u32>,
    pub end: Option<u32>,
    pub new_paragraph_ids: Vec<String>,
    pub revision_ids: Vec<String>,
}

#[napi(object)]
pub struct LayoutResult {
    pub layout: serde_json::Value,
    pub display_list: serde_json::Value,
}

pub struct OpenTask {
    bytes: Vec<u8>,
    author: String,
    origin: EditOrigin,
    timestamp: String,
}

impl Task for OpenTask {
    type Output = Document;
    type JsValue = DocxDocument;

    fn compute(&mut self) -> Result<Self::Output> {
        Document::open(&self.bytes).map_err(error)
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(DocxDocument {
            inner: Arc::new(Mutex::new(output)),
            author: self.author.clone(),
            origin: self.origin,
            timestamp: self.timestamp.clone(),
        })
    }
}

pub struct LayoutTask {
    document: Arc<Mutex<Document>>,
    input: LayoutInput,
}

impl Task for LayoutTask {
    type Output = (serde_json::Value, serde_json::Value);
    type JsValue = LayoutResult;

    fn compute(&mut self) -> Result<Self::Output> {
        let result = lock(&self.document)?
            .layout(self.input.clone())
            .map_err(error)?;
        Ok((
            serde_json::to_value(result.layout).map_err(error)?,
            serde_json::to_value(result.display_list).map_err(error)?,
        ))
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(LayoutResult {
            layout: output.0,
            display_list: output.1,
        })
    }
}

pub struct RenderTask {
    document: Arc<Mutex<Document>>,
    display_list: betteroffice_docx::DisplayList,
    page: usize,
}

impl Task for RenderTask {
    type Output = betteroffice_docx::RenderedPage;
    type JsValue = RenderedPage;

    fn compute(&mut self) -> Result<Self::Output> {
        lock(&self.document)?
            .render_png(&self.display_list, self.page)
            .map_err(error)
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(RenderedPage {
            data: output.bytes.into(),
            skipped_images: output.skipped_images as u32,
        })
    }
}

pub struct SaveTask {
    document: Arc<Mutex<Document>>,
    options: SaveOptions,
}

impl Task for SaveTask {
    type Output = Vec<u8>;
    type JsValue = Buffer;

    fn compute(&mut self) -> Result<Self::Output> {
        lock(&self.document)?
            .save_with_options(self.options.clone())
            .map_err(error)
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output.into())
    }
}

#[napi(js_name = "Document")]
pub struct DocxDocument {
    inner: Arc<Mutex<Document>>,
    author: String,
    origin: EditOrigin,
    timestamp: String,
}

#[napi]
impl DocxDocument {
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
    pub fn timestamp(&self) -> String {
        self.timestamp.clone()
    }

    #[napi(setter)]
    pub fn set_timestamp(&mut self, timestamp: String) {
        self.timestamp = timestamp;
    }

    #[napi(getter)]
    pub fn paragraph_ids(&self) -> Result<Vec<Option<String>>> {
        Ok(lock(&self.inner)?
            .paragraphs()
            .into_iter()
            .map(|paragraph| paragraph.para_id.clone())
            .collect())
    }

    #[napi(getter)]
    pub fn warnings(&self) -> Result<Vec<String>> {
        Ok(lock(&self.inner)?.model().warnings.clone())
    }

    #[napi(getter)]
    pub fn template_variables(&self) -> Result<Vec<String>> {
        Ok(lock(&self.inner)?.model().template_variables.clone())
    }

    #[napi(getter)]
    pub fn text(&self) -> Result<String> {
        Ok(lock(&self.inner)?
            .paragraphs()
            .into_iter()
            .map(betteroffice_docx::get_paragraph_text)
            .filter(|text| !text.is_empty())
            .collect::<Vec<_>>()
            .join("\n"))
    }

    #[napi(getter)]
    pub fn structure(&self) -> Result<DocumentStructure> {
        let value = lock(&self.inner)?.structure();
        Ok(DocumentStructure {
            body_paragraphs: value.body_paragraphs as u32,
            body_tables: value.body_tables as u32,
            sections: value.sections as u32,
            headers: value.headers as u32,
            footers: value.footers as u32,
            footnotes: value.footnotes as u32,
            endnotes: value.endnotes as u32,
        })
    }

    #[napi]
    pub fn body(&self) -> Result<serde_json::Value> {
        serde_json::to_value(lock(&self.inner)?.body()).map_err(error)
    }

    #[napi]
    pub fn headers(&self) -> Result<serde_json::Value> {
        serde_json::to_value(lock(&self.inner)?.headers()).map_err(error)
    }

    #[napi]
    pub fn footers(&self) -> Result<serde_json::Value> {
        serde_json::to_value(lock(&self.inner)?.footers()).map_err(error)
    }

    #[napi]
    pub fn sections(&self) -> Result<serde_json::Value> {
        serde_json::to_value(lock(&self.inner)?.sections()).map_err(error)
    }

    #[napi]
    pub fn paragraphs(&self) -> Result<serde_json::Value> {
        serde_json::to_value(lock(&self.inner)?.paragraphs()).map_err(error)
    }

    #[napi]
    pub fn tables(&self) -> Result<serde_json::Value> {
        serde_json::to_value(lock(&self.inner)?.tables()).map_err(error)
    }

    #[napi]
    pub fn paragraph(&self, paragraph_id: String) -> Result<Option<serde_json::Value>> {
        lock(&self.inner)?
            .paragraph(&paragraph_id)
            .map(serde_json::to_value)
            .transpose()
            .map_err(error)
    }

    #[napi]
    pub fn replace_paragraph_text(
        &self,
        paragraph_id: String,
        text: String,
    ) -> Result<EditReceipt> {
        let context = EditCtx {
            author: self.author.clone(),
            origin: self.origin,
            suggesting: None,
            now_iso: self.timestamp.clone(),
        };
        let receipt = lock(&self.inner)?
            .replace_paragraph_text_with(&paragraph_id, &text, 1, &context)
            .map_err(error)?;
        let range = receipt.range;
        Ok(EditReceipt {
            paragraph_id: range.as_ref().map(|value| value.start.para.clone()),
            story: range.as_ref().map(|value| value.start.story.clone()),
            start: range.as_ref().map(|value| value.start.offset),
            end: range.as_ref().map(|value| value.end.offset),
            new_paragraph_ids: receipt.new_para_ids,
            revision_ids: receipt.revision_ids,
        })
    }

    #[napi(ts_return_type = "Promise<LayoutResult>")]
    pub fn layout(&self, input: serde_json::Value) -> Result<AsyncTask<LayoutTask>> {
        let input = serde_json::from_value(input).map_err(error)?;
        Ok(AsyncTask::new(LayoutTask {
            document: self.inner.clone(),
            input,
        }))
    }

    #[napi]
    pub fn register_font(&self, face: FontFace) -> Result<u32> {
        lock(&self.inner)?
            .register_font(
                &face.family,
                face.bold.unwrap_or(false),
                face.italic.unwrap_or(false),
                &face.data,
            )
            .map_err(error)
    }

    #[napi]
    pub fn register_image(&self, image: ImageResource) -> Result<()> {
        let scope = image.scope.as_deref().unwrap_or("body");
        let scope = match scope {
            "body" => ImageScope::Body,
            "headerFooter" => ImageScope::HeaderFooter(
                image
                    .part
                    .as_deref()
                    .ok_or_else(|| error("headerFooter images require part"))?,
            ),
            "footnotes" => ImageScope::Notes(NoteKind::Footnote),
            "endnotes" => ImageScope::Notes(NoteKind::Endnote),
            _ => {
                return Err(error(
                    "scope must be body, headerFooter, footnotes, or endnotes",
                ));
            }
        };
        lock(&self.inner)?
            .register_image(scope, &image.relationship_id, &image.data)
            .map_err(error)
    }

    #[napi(ts_return_type = "Promise<RenderedPage>")]
    pub fn render_page(
        &self,
        display_list: serde_json::Value,
        page: Option<u32>,
    ) -> Result<AsyncTask<RenderTask>> {
        Ok(AsyncTask::new(RenderTask {
            document: self.inner.clone(),
            display_list: serde_json::from_value(display_list).map_err(error)?,
            page: page.unwrap_or(0) as usize,
        }))
    }

    #[napi(ts_return_type = "Promise<Buffer>")]
    pub fn save(&self, options: Option<SaveDocumentOptions>) -> Result<AsyncTask<SaveTask>> {
        let options = options.unwrap_or(SaveDocumentOptions {
            timestamp: None,
            update_modified_date: None,
            modified_by: None,
        });
        Ok(AsyncTask::new(SaveTask {
            document: self.inner.clone(),
            options: SaveOptions {
                now: options.timestamp.unwrap_or_else(|| self.timestamp.clone()),
                update_modified_date: options.update_modified_date.unwrap_or(false),
                modified_by: options.modified_by,
            },
        }))
    }
}

#[napi(ts_return_type = "Promise<Document>")]
pub fn open_document(
    data: Buffer,
    options: Option<OpenDocumentOptions>,
) -> Result<AsyncTask<OpenTask>> {
    let options = options.unwrap_or(OpenDocumentOptions {
        author: None,
        origin: None,
        timestamp: None,
    });
    Ok(AsyncTask::new(OpenTask {
        bytes: data.to_vec(),
        author: options.author.unwrap_or_else(|| "node".to_owned()),
        origin: options
            .origin
            .as_deref()
            .map(origin)
            .transpose()?
            .unwrap_or(EditOrigin::Local),
        timestamp: options
            .timestamp
            .unwrap_or_else(|| SaveOptions::default().now),
    }))
}
