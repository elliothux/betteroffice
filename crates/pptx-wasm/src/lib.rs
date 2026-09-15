//! PPTX display-list wasm boundary.

use std::io::Cursor;

use wasm_bindgen::prelude::*;

pub use pptx_edit::wasm::PptxDocument;

const MAX_IMAGE_PIXELS: u64 = 33_554_432;
const MAX_IMAGE_BYTES: u64 = 268_435_456;

#[wasm_bindgen]
pub struct PptxRenderer {
    renderer: pptx_render::SlideRenderer,
    rendered: Option<pptx_render::RenderedSlide>,
}

#[wasm_bindgen]
impl PptxRenderer {
    #[wasm_bindgen(constructor)]
    pub fn new() -> PptxRenderer {
        Self {
            renderer: pptx_render::SlideRenderer::new(),
            rendered: None,
        }
    }

    #[wasm_bindgen(js_name = registerFont)]
    pub fn register_font(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        bytes: &[u8],
    ) -> Result<u32, JsValue> {
        self.renderer
            .register_font(family, bold, italic, bytes)
            .map_err(js_error)
    }

    #[wasm_bindgen(js_name = layoutSlideJson)]
    pub fn layout_slide_json(
        &mut self,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        let session = document.session();
        let deck = session.snapshot().map_err(js_error)?;
        let rendered = self
            .renderer
            .layout_slide(session.package(), &deck, slide_index as usize)
            .map_err(js_error)?;
        let json = serde_json::to_string(&rendered.display_list).map_err(js_error)?;
        self.rendered = Some(rendered);
        Ok(json)
    }

    #[wasm_bindgen(js_name = hitTestJson)]
    pub fn hit_test_json(&self, x: f32, y: f32) -> Result<String, JsValue> {
        let result = self
            .rendered
            .as_ref()
            .and_then(|rendered| rendered.hit_test(x, y));
        serde_json::to_string(&result).map_err(js_error)
    }

    #[wasm_bindgen(js_name = layoutProposalSlideJson)]
    pub fn layout_proposal_slide_json(
        &self,
        document: &PptxDocument,
        id: &str,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        let preview = document
            .session()
            .proposal_preview_session(id)
            .map_err(js_error)?;
        let rendered = self
            .renderer
            .layout_slide(
                preview.package(),
                &preview.snapshot().map_err(js_error)?,
                slide_index as usize,
            )
            .map_err(js_error)?;
        serde_json::to_string(&rendered.display_list).map_err(js_error)
    }

    #[wasm_bindgen(js_name = layoutProposalDiffSlideJson)]
    pub fn layout_proposal_diff_slide_json(
        &self,
        document: &PptxDocument,
        id: &str,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        let session = document.session();
        let preview = session.preview_proposal_diff(id).map_err(js_error)?;
        let rendered = self
            .renderer
            .layout_slide(session.package(), &preview.snapshot, slide_index as usize)
            .map_err(js_error)?;
        serde_json::to_string(&serde_json::json!({
            "proposal": preview.proposal,
            "snapshot": preview.snapshot,
            "textChanges": preview.text_changes,
            "frame": rendered.display_list,
        }))
        .map_err(js_error)
    }
}

impl Default for PptxRenderer {
    fn default() -> Self {
        Self::new()
    }
}

#[wasm_bindgen(js_name = parsePptxJson)]
pub fn parse_pptx_json(data: &[u8]) -> Result<String, JsValue> {
    let package = pptx_parse::parse_pptx(data).map_err(js_error)?;
    serde_json::to_string(&package).map_err(js_error)
}

#[wasm_bindgen(js_name = compileSlideJson)]
pub fn compile_slide_json(slide_json: &str) -> Result<String, JsValue> {
    pptx_render::compile_json(slide_json).map_err(|error| JsValue::from_str(&error))
}

#[wasm_bindgen(js_name = rendererVersion)]
pub fn renderer_version() -> String {
    env!("CARGO_PKG_VERSION").to_owned()
}

#[wasm_bindgen(js_name = decodeTiffPng)]
pub fn decode_tiff_png(data: &[u8]) -> Result<Vec<u8>, JsValue> {
    decode_tiff(data).map_err(js_error)
}

fn decode_tiff(data: &[u8]) -> Result<Vec<u8>, String> {
    use image::{ImageDecoder as _, ImageEncoder as _};

    let mut decoder = image::ImageReader::with_format(Cursor::new(data), image::ImageFormat::Tiff)
        .into_decoder()
        .map_err(|error| error.to_string())?;
    let (width, height) = decoder.dimensions();
    let pixels = u64::from(width) * u64::from(height);
    let bytes = decoder
        .total_bytes()
        .saturating_add(pixels.saturating_mul(4));
    if !image_fits_budget(pixels, bytes) {
        return Err("TIFF image exceeds the browser decode budget".to_owned());
    }
    let orientation = decoder.orientation().map_err(|error| error.to_string())?;
    let mut decoded =
        image::DynamicImage::from_decoder(decoder).map_err(|error| error.to_string())?;
    decoded.apply_orientation(orientation);
    let rgba = decoded.into_rgba8();
    let mut png = Vec::new();
    image::codecs::png::PngEncoder::new(&mut png)
        .write_image(
            rgba.as_raw(),
            rgba.width(),
            rgba.height(),
            image::ExtendedColorType::Rgba8,
        )
        .map_err(|error| error.to_string())?;
    Ok(png)
}

fn image_fits_budget(pixels: u64, bytes: u64) -> bool {
    pixels <= MAX_IMAGE_PIXELS && bytes <= MAX_IMAGE_BYTES
}

fn js_error(error: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use image::ImageEncoder as _;

    #[test]
    fn decodes_tiff_to_png() {
        let mut tiff = Cursor::new(Vec::new());
        image::codecs::tiff::TiffEncoder::new(&mut tiff)
            .write_image(
                &[255, 0, 0, 255, 0, 128, 255, 64],
                2,
                1,
                image::ExtendedColorType::Rgba8,
            )
            .unwrap();
        let png = decode_tiff(&tiff.into_inner()).unwrap();
        let decoded = image::load_from_memory_with_format(&png, image::ImageFormat::Png)
            .unwrap()
            .into_rgba8();
        assert_eq!(decoded.dimensions(), (2, 1));
        assert_eq!(decoded.as_raw(), &[255, 0, 0, 255, 0, 128, 255, 64]);
    }

    #[test]
    fn rejects_invalid_tiff() {
        assert!(decode_tiff(b"II*\0").is_err());
    }

    #[test]
    fn rejects_images_past_the_decode_budget() {
        assert!(!image_fits_budget(MAX_IMAGE_PIXELS + 1, 0));
        assert!(!image_fits_budget(0, MAX_IMAGE_BYTES + 1));
    }
}
