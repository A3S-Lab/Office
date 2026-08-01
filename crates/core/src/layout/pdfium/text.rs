use std::ops::Range;

use a3s_use_core::{UseError, UseResult};
use pdfium_render::prelude::{PdfPage, PdfRect};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use super::engine::{signed_points_to_millipoints, PDFIUM_ENGINE_VERSION};
use super::{NativeOfficePdfPageBox, NativeOfficePdfPageGeometry, NativeOfficePdfPageInventory};
use crate::layout::layout_error;
use crate::{NativeOfficeUnit, PackageRevision};

/// Schema version for native PDF text-layer and outline receipts.
pub const NATIVE_OFFICE_PDF_TEXT_SCHEMA_VERSION: u32 = 1;
/// Default maximum number of PDFium characters accepted from one page.
pub const DEFAULT_NATIVE_OFFICE_PDF_TEXT_CHARACTERS: usize = 1_000_000;
/// Hard maximum number of PDFium characters accepted from one page.
pub const MAX_NATIVE_OFFICE_PDF_TEXT_CHARACTERS: usize = 5_000_000;
/// Default maximum UTF-8 text bytes accepted from one page.
pub const DEFAULT_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES: usize = 16 * 1024 * 1024;
/// Hard maximum UTF-8 text bytes accepted from one page.
pub const MAX_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES: usize = 64 * 1024 * 1024;

/// Explicit bounds for one native PDF page text-layer extraction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfTextLayerOptions {
    pub max_characters: usize,
    pub max_text_bytes: usize,
    pub timeout_ms: u64,
}

impl Default for NativeOfficePdfTextLayerOptions {
    fn default() -> Self {
        Self {
            max_characters: DEFAULT_NATIVE_OFFICE_PDF_TEXT_CHARACTERS,
            max_text_bytes: DEFAULT_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES,
            timeout_ms: 120_000,
        }
    }
}

impl NativeOfficePdfTextLayerOptions {
    pub(super) fn validate(&self) -> UseResult<()> {
        if (1..=MAX_NATIVE_OFFICE_PDF_TEXT_CHARACTERS).contains(&self.max_characters)
            && (1..=MAX_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES).contains(&self.max_text_bytes)
            && (1..=super::super::MAX_LAYOUT_TIMEOUT_MS).contains(&self.timeout_ms)
        {
            return Ok(());
        }
        Err(layout_error(
            "use.office.pdf_text_options_invalid",
            "PDF text character, byte, and timeout bounds are invalid.",
        ))
    }
}

/// One PDFium character in source order with exact string offsets and optional
/// native PDF-space glyph geometry.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfTextCharacter {
    pub index: u32,
    pub text: String,
    pub utf8_start: u64,
    pub utf8_end: u64,
    pub utf16_start: u64,
    pub utf16_end: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bounds: Option<NativeOfficePdfPageBox>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size_millipoints: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation_millidegrees: Option<i32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub generated: Option<bool>,
}

impl NativeOfficePdfTextCharacter {
    pub fn utf8_range(&self) -> Range<u64> {
        self.utf8_start..self.utf8_end
    }

    pub fn utf16_range(&self) -> Range<u64> {
        self.utf16_start..self.utf16_end
    }

    fn validate(
        &self,
        expected_index: usize,
        utf8_offset: u64,
        utf16_offset: u64,
    ) -> UseResult<(u64, u64)> {
        let index = u32::try_from(expected_index).map_err(|_| invalid_text_layer())?;
        let scalar_count = self.text.chars().count();
        let utf8_length = u64::try_from(self.text.len()).map_err(|_| invalid_text_layer())?;
        let utf16_length =
            u64::try_from(self.text.encode_utf16().count()).map_err(|_| invalid_text_layer())?;
        let expected_utf8_end = utf8_offset
            .checked_add(utf8_length)
            .ok_or_else(invalid_text_layer)?;
        let expected_utf16_end = utf16_offset
            .checked_add(utf16_length)
            .ok_or_else(invalid_text_layer)?;
        if self.index != index
            || scalar_count > 1
            || self.utf8_start != utf8_offset
            || self.utf8_end != expected_utf8_end
            || self.utf16_start != utf16_offset
            || self.utf16_end != expected_utf16_end
            || self.font_size_millipoints == Some(0)
            || self
                .rotation_millidegrees
                .is_some_and(|angle| !(-360_000..=360_000).contains(&angle))
        {
            return Err(invalid_text_layer());
        }
        if let Some(bounds) = &self.bounds {
            bounds.validate().map_err(|_| invalid_text_layer())?;
        }
        Ok((expected_utf8_end, expected_utf16_end))
    }
}

/// Complete native text evidence for one inventoried PDF page.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfPageTextLayer {
    pub schema_version: u32,
    pub source_revision: PackageRevision,
    pub unit: NativeOfficeUnit,
    pub page_geometry: NativeOfficePdfPageGeometry,
    pub engine_version: String,
    pub max_characters: usize,
    pub max_text_bytes: usize,
    pub text_sha256: String,
    pub text: String,
    pub characters: Vec<NativeOfficePdfTextCharacter>,
}

impl NativeOfficePdfPageTextLayer {
    /// Validates source/page identity, bounds, exact UTF-8/UTF-16 ranges, and
    /// deterministic content identity against the admitted inventory.
    pub fn validate(&self, inventory: &NativeOfficePdfPageInventory) -> UseResult<()> {
        inventory.validate()?;
        if self.source_revision != inventory.source_revision {
            return Err(layout_error(
                "use.office.pdf_text_layer_source_mismatch",
                "The PDF text layer belongs to a different immutable source revision.",
            ));
        }
        let page = inventory.validated_page(&self.unit)?;
        if self.schema_version != NATIVE_OFFICE_PDF_TEXT_SCHEMA_VERSION
            || self.page_geometry != *page
            || self.engine_version != PDFIUM_ENGINE_VERSION
            || !(1..=MAX_NATIVE_OFFICE_PDF_TEXT_CHARACTERS).contains(&self.max_characters)
            || !(1..=MAX_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES).contains(&self.max_text_bytes)
            || self.characters.len() > self.max_characters
            || self.text.len() > self.max_text_bytes
            || self.text_sha256 != format!("{:x}", Sha256::digest(self.text.as_bytes()))
        {
            return Err(invalid_text_layer());
        }

        let mut rebuilt = String::with_capacity(self.text.len());
        let mut utf8_offset = 0_u64;
        let mut utf16_offset = 0_u64;
        for (index, character) in self.characters.iter().enumerate() {
            (utf8_offset, utf16_offset) = character.validate(index, utf8_offset, utf16_offset)?;
            rebuilt.push_str(&character.text);
        }
        if rebuilt != self.text
            || utf8_offset != u64::try_from(self.text.len()).unwrap_or(u64::MAX)
            || utf16_offset != u64::try_from(self.text.encode_utf16().count()).unwrap_or(u64::MAX)
        {
            return Err(invalid_text_layer());
        }
        Ok(())
    }
}

pub(super) fn extract_page_text(
    page: &PdfPage<'_>,
    source_revision: PackageRevision,
    page_geometry: NativeOfficePdfPageGeometry,
    options: NativeOfficePdfTextLayerOptions,
) -> UseResult<NativeOfficePdfPageTextLayer> {
    options.validate()?;
    let page_text = page.text().map_err(|_| text_unsupported())?;
    let count = usize::try_from(page_text.len()).map_err(|_| text_unsupported())?;
    if count > options.max_characters {
        return Err(layout_error(
            "use.office.pdf_text_character_limit",
            format!(
                "PDF page contains {count} text characters; the configured limit is {}.",
                options.max_characters
            ),
        ));
    }

    let chars = page_text.chars();
    let mut text = String::new();
    let mut characters = Vec::with_capacity(count);
    let mut utf16_offset = 0_u64;
    for (offset, character) in chars.iter().enumerate() {
        if character.index() != offset {
            return Err(text_unsupported());
        }
        let character_text = character
            .unicode_char()
            .map(|value| value.to_string())
            .unwrap_or_default();
        let utf8_start = u64::try_from(text.len()).map_err(|_| text_byte_limit(options))?;
        let utf16_start = utf16_offset;
        text.push_str(&character_text);
        if text.len() > options.max_text_bytes {
            return Err(text_byte_limit(options));
        }
        utf16_offset = utf16_offset
            .checked_add(
                u64::try_from(character_text.encode_utf16().count())
                    .map_err(|_| text_byte_limit(options))?,
            )
            .ok_or_else(|| text_byte_limit(options))?;
        let index = u32::try_from(offset).map_err(|_| text_unsupported())?;
        characters.push(NativeOfficePdfTextCharacter {
            index,
            text: character_text,
            utf8_start,
            utf8_end: u64::try_from(text.len()).map_err(|_| text_byte_limit(options))?,
            utf16_start,
            utf16_end: utf16_offset,
            bounds: character.loose_bounds().ok().and_then(pdf_rect_to_page_box),
            font_size_millipoints: font_size_millipoints(character.scaled_font_size().value),
            rotation_millidegrees: character
                .angle_degrees()
                .ok()
                .and_then(rotation_millidegrees),
            generated: character.is_generated().ok(),
        });
    }

    let layer = NativeOfficePdfPageTextLayer {
        schema_version: NATIVE_OFFICE_PDF_TEXT_SCHEMA_VERSION,
        source_revision,
        unit: page_geometry.unit.clone(),
        page_geometry,
        engine_version: PDFIUM_ENGINE_VERSION.to_string(),
        max_characters: options.max_characters,
        max_text_bytes: options.max_text_bytes,
        text_sha256: format!("{:x}", Sha256::digest(text.as_bytes())),
        text,
        characters,
    };
    Ok(layer)
}

fn pdf_rect_to_page_box(rect: PdfRect) -> Option<NativeOfficePdfPageBox> {
    let bounds = NativeOfficePdfPageBox {
        left_millipoints: signed_points_to_millipoints(rect.left().value).ok()?,
        bottom_millipoints: signed_points_to_millipoints(rect.bottom().value).ok()?,
        right_millipoints: signed_points_to_millipoints(rect.right().value).ok()?,
        top_millipoints: signed_points_to_millipoints(rect.top().value).ok()?,
    };
    bounds.validate().ok().map(|()| bounds)
}

fn font_size_millipoints(points: f32) -> Option<u32> {
    let millipoints = f64::from(points) * 1_000.0;
    if millipoints.is_finite() && millipoints > 0.0 && millipoints <= f64::from(u32::MAX) {
        Some(millipoints.round() as u32)
    } else {
        None
    }
}

fn rotation_millidegrees(degrees: f32) -> Option<i32> {
    let millidegrees = f64::from(degrees) * 1_000.0;
    if millidegrees.is_finite() && (-360_000.0..=360_000.0).contains(&millidegrees) {
        Some(millidegrees.round() as i32)
    } else {
        None
    }
}

fn invalid_text_layer() -> UseError {
    layout_error(
        "use.office.pdf_text_layer_invalid",
        "PDF text-layer content, ranges, geometry, or identity are inconsistent.",
    )
}

fn text_unsupported() -> UseError {
    layout_error(
        "use.office.pdf_text_unsupported",
        "PDFium could not extract a valid native text layer from the requested page.",
    )
}

fn text_byte_limit(options: NativeOfficePdfTextLayerOptions) -> UseError {
    layout_error(
        "use.office.pdf_text_byte_limit",
        format!(
            "PDF page text exceeds the configured {}-byte limit.",
            options.max_text_bytes
        ),
    )
}
