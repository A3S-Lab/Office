mod html;
mod image;
mod output;
mod svg;

use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use crate::{DocumentKind, NativeOfficeDocument};

/// Maximum UTF-8 bytes produced by one native semantic render.
pub const MAX_NATIVE_OFFICE_RENDER_BYTES: usize = 16 * 1024 * 1024;

/// Standalone semantic render formats owned by the native Office engine.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeRenderFormat {
    Html,
    Svg,
}

impl NativeOfficeRenderFormat {
    pub fn media_type(self) -> &'static str {
        match self {
            Self::Html => "text/html; charset=utf-8",
            Self::Svg => "image/svg+xml",
        }
    }
}

/// Deterministic, standalone semantic view of one native Office document.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeRenderedView {
    pub kind: DocumentKind,
    pub format: NativeOfficeRenderFormat,
    pub media_type: String,
    pub content: String,
    pub unit_count: usize,
    pub byte_length: usize,
    pub sha256: String,
}

impl NativeOfficeDocument {
    /// Renders a bounded, deterministic semantic view without an Office runtime.
    pub fn render(&self, format: NativeOfficeRenderFormat) -> UseResult<NativeOfficeRenderedView> {
        render_with_limit(self, format, MAX_NATIVE_OFFICE_RENDER_BYTES)
    }

    /// Renders standalone semantic HTML for Word, Spreadsheet, or Presentation.
    pub fn html_view(&self) -> UseResult<NativeOfficeRenderedView> {
        self.render(NativeOfficeRenderFormat::Html)
    }

    /// Renders standalone semantic SVG for Word, Spreadsheet, or Presentation.
    pub fn svg_view(&self) -> UseResult<NativeOfficeRenderedView> {
        self.render(NativeOfficeRenderFormat::Svg)
    }
}

pub(crate) fn render_with_limit(
    document: &NativeOfficeDocument,
    format: NativeOfficeRenderFormat,
    limit: usize,
) -> UseResult<NativeOfficeRenderedView> {
    if limit == 0 || limit > MAX_NATIVE_OFFICE_RENDER_BYTES {
        return Err(render_error(
            "use.office.render_limit_invalid",
            format!(
                "Native Office render limit must be between 1 and {MAX_NATIVE_OFFICE_RENDER_BYTES} bytes."
            ),
        ));
    }
    let content = match format {
        NativeOfficeRenderFormat::Html => html::render(document, limit)?,
        NativeOfficeRenderFormat::Svg => svg::render(document, limit)?,
    };
    let byte_length = content.len();
    let sha256 = format!("{:x}", Sha256::digest(content.as_bytes()));
    Ok(NativeOfficeRenderedView {
        kind: document.kind(),
        format,
        media_type: format.media_type().to_string(),
        content,
        unit_count: unit_count(document),
        byte_length,
        sha256,
    })
}

fn unit_count(document: &NativeOfficeDocument) -> usize {
    match document.kind() {
        DocumentKind::Word => 1,
        DocumentKind::Spreadsheet | DocumentKind::Presentation => document.root().children.len(),
    }
}

pub(super) fn render_error(code: &'static str, message: impl Into<String>) -> UseError {
    UseError::new(code, message)
}
