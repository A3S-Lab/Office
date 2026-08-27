use std::ops::Range;

use a3s_use_core::{UseError, UseResult};
use pdfium_render::prelude::{PdfPage, PdfRect};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use super::engine::{signed_points_to_millipoints, PDFIUM_ENGINE_VERSION};
use super::{NativeOfficePdfPageBox, NativeOfficePdfPageGeometry, NativeOfficePdfPageInventory};
use crate::layout::layout_error;
use crate::{NativeOfficeUnit, PackageRevision};

mod objects;
mod style;

use objects::page_object_inventory;
pub use objects::{
    NativeOfficePdfPageObjectSummary, NativeOfficePdfVisualObject,
    NativeOfficePdfVisualObjectInventory, NativeOfficePdfVisualObjectKind,
    MAX_NATIVE_OFFICE_PDF_VISUAL_OBJECTS,
};
use style::font_weight_value;
pub use style::{NativeOfficePdfTextColor, NativeOfficePdfTextRenderMode};

/// Schema version for native PDF text-layer and outline receipts.
pub const NATIVE_OFFICE_PDF_TEXT_SCHEMA_VERSION: u32 = 5;
/// Default maximum number of PDFium characters accepted from one page.
pub const DEFAULT_NATIVE_OFFICE_PDF_TEXT_CHARACTERS: usize = 1_000_000;
/// Hard maximum number of PDFium characters accepted from one page.
pub const MAX_NATIVE_OFFICE_PDF_TEXT_CHARACTERS: usize = 5_000_000;
/// Default maximum number of PDFium text runs accepted from one page.
pub const DEFAULT_NATIVE_OFFICE_PDF_TEXT_RUNS: usize = 250_000;
/// Hard maximum number of PDFium text runs accepted from one page.
pub const MAX_NATIVE_OFFICE_PDF_TEXT_RUNS: usize = 1_000_000;
/// Default maximum UTF-8 text bytes accepted from one page.
pub const DEFAULT_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES: usize = 16 * 1024 * 1024;
/// Hard maximum UTF-8 text bytes accepted from one page.
pub const MAX_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES: usize = 64 * 1024 * 1024;

/// Explicit bounds for one native PDF page text-layer extraction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfTextLayerOptions {
    pub max_characters: usize,
    pub max_runs: usize,
    pub max_text_bytes: usize,
    pub timeout_ms: u64,
}

impl Default for NativeOfficePdfTextLayerOptions {
    fn default() -> Self {
        Self {
            max_characters: DEFAULT_NATIVE_OFFICE_PDF_TEXT_CHARACTERS,
            max_runs: DEFAULT_NATIVE_OFFICE_PDF_TEXT_RUNS,
            max_text_bytes: DEFAULT_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES,
            timeout_ms: 120_000,
        }
    }
}

impl NativeOfficePdfTextLayerOptions {
    pub(super) fn validate(&self) -> UseResult<()> {
        if (1..=MAX_NATIVE_OFFICE_PDF_TEXT_CHARACTERS).contains(&self.max_characters)
            && (1..=MAX_NATIVE_OFFICE_PDF_TEXT_RUNS).contains(&self.max_runs)
            && (1..=MAX_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES).contains(&self.max_text_bytes)
            && (1..=super::super::MAX_LAYOUT_TIMEOUT_MS).contains(&self.timeout_ms)
        {
            return Ok(());
        }
        Err(layout_error(
            "use.office.pdf_text_options_invalid",
            "PDF text character, run, byte, and timeout bounds are invalid.",
        ))
    }
}

/// One conservatively representable PDFium-native same-line text segment in
/// source order. `runs` may contain gaps when a PDFium
/// rectangle cannot be mapped to one unambiguous contiguous character range.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfTextRun {
    pub index: u32,
    pub character_start: u32,
    pub character_end: u32,
    pub utf8_start: u64,
    pub utf8_end: u64,
    pub utf16_start: u64,
    pub utf16_end: u64,
    pub text: String,
    pub bounds: NativeOfficePdfPageBox,
}

impl NativeOfficePdfTextRun {
    fn validate(
        &self,
        expected_index: usize,
        characters: &[NativeOfficePdfTextCharacter],
        text: &str,
    ) -> UseResult<()> {
        let index = u32::try_from(expected_index).map_err(|_| invalid_text_layer())?;
        let start = usize::try_from(self.character_start).map_err(|_| invalid_text_layer())?;
        let end = usize::try_from(self.character_end).map_err(|_| invalid_text_layer())?;
        let first = characters.get(start).ok_or_else(invalid_text_layer)?;
        let last = end
            .checked_sub(1)
            .and_then(|offset| characters.get(offset))
            .ok_or_else(invalid_text_layer)?;
        let utf8_start = usize::try_from(first.utf8_start).map_err(|_| invalid_text_layer())?;
        let utf8_end = usize::try_from(last.utf8_end).map_err(|_| invalid_text_layer())?;
        if self.index != index
            || start >= end
            || end > characters.len()
            || self.utf8_start != first.utf8_start
            || self.utf8_end != last.utf8_end
            || self.utf16_start != first.utf16_start
            || self.utf16_end != last.utf16_end
            || text.get(utf8_start..utf8_end) != Some(self.text.as_str())
        {
            return Err(invalid_text_layer());
        }
        self.bounds.validate().map_err(|_| invalid_text_layer())
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
    pub font_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_weight: Option<u16>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<NativeOfficePdfTextColor>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_color: Option<NativeOfficePdfTextColor>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub render_mode: Option<NativeOfficePdfTextRenderMode>,
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
            || self.font_name.as_ref().is_some_and(|name| {
                name.is_empty() || name.len() > 1_024 || name.chars().any(char::is_control)
            })
            || self
                .font_weight
                .is_some_and(|weight| !(1..=1_000).contains(&weight))
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
    pub max_runs: usize,
    pub max_text_bytes: usize,
    pub page_objects: NativeOfficePdfPageObjectSummary,
    pub visual_objects: NativeOfficePdfVisualObjectInventory,
    pub text_sha256: String,
    pub text: String,
    pub characters: Vec<NativeOfficePdfTextCharacter>,
    pub runs: Vec<NativeOfficePdfTextRun>,
}

impl NativeOfficePdfPageTextLayer {
    /// Validates source/page identity, bounds, exact UTF-8/UTF-16 ranges, and
    /// deterministic content identity against the admitted inventory.
    pub fn validate(&self, inventory: &NativeOfficePdfPageInventory) -> UseResult<()> {
        inventory.validate()?;
        self.validate_for_prevalidated_inventory(inventory)
    }

    /// Validate one layer after the complete inventory envelope, order, and
    /// page geometries were validated by the enclosing batch request.
    pub(super) fn validate_for_prevalidated_inventory(
        &self,
        inventory: &NativeOfficePdfPageInventory,
    ) -> UseResult<()> {
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
            || !(1..=MAX_NATIVE_OFFICE_PDF_TEXT_RUNS).contains(&self.max_runs)
            || !(1..=MAX_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES).contains(&self.max_text_bytes)
            || self.characters.len() > self.max_characters
            || self.runs.len() > self.max_runs
            || self.text.len() > self.max_text_bytes
            || self.text_sha256 != format!("{:x}", Sha256::digest(self.text.as_bytes()))
        {
            return Err(invalid_text_layer());
        }
        self.page_objects.validate()?;
        self.visual_objects.validate(self.page_objects)?;

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
        let mut previous_end = 0_usize;
        for (index, run) in self.runs.iter().enumerate() {
            let start = usize::try_from(run.character_start).map_err(|_| invalid_text_layer())?;
            if start < previous_end {
                return Err(invalid_text_layer());
            }
            run.validate(index, &self.characters, &self.text)?;
            previous_end = usize::try_from(run.character_end).map_err(|_| invalid_text_layer())?;
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
    let (page_objects, visual_objects) = page_object_inventory(page)?;
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
            font_name: non_empty_font_name(character.font_name()),
            font_weight: character.font_weight().and_then(font_weight_value),
            italic: Some(character.font_is_italic()),
            fill_color: character.fill_color().ok().map(Into::into),
            stroke_color: character.stroke_color().ok().map(Into::into),
            render_mode: character.render_mode().ok().map(Into::into),
            generated: character.is_generated().ok(),
        });
    }

    let mut runs = Vec::new();
    let mut previous_end = 0_usize;
    for segment in page_text.segments().iter() {
        let Ok(segment_chars) = segment.chars() else {
            continue;
        };
        let Some(range) = pdfium_segment_range(
            segment_chars.iter().map(|character| character.index()),
            characters.len(),
            previous_end,
        )?
        else {
            continue;
        };
        let start = range.start;
        let end = range.end;
        let Some(bounds) = pdf_rect_to_page_box(segment.bounds()) else {
            continue;
        };
        if runs.len() >= options.max_runs {
            return Err(layout_error(
                "use.office.pdf_text_run_limit",
                format!(
                    "PDF page contains more than {} text runs.",
                    options.max_runs
                ),
            ));
        }
        let first = &characters[start];
        let last = &characters[end - 1];
        let utf8_start = usize::try_from(first.utf8_start).map_err(|_| text_unsupported())?;
        let utf8_end = usize::try_from(last.utf8_end).map_err(|_| text_unsupported())?;
        runs.push(NativeOfficePdfTextRun {
            index: u32::try_from(runs.len()).map_err(|_| text_unsupported())?,
            character_start: u32::try_from(start).map_err(|_| text_unsupported())?,
            character_end: u32::try_from(end).map_err(|_| text_unsupported())?,
            utf8_start: first.utf8_start,
            utf8_end: last.utf8_end,
            utf16_start: first.utf16_start,
            utf16_end: last.utf16_end,
            text: text
                .get(utf8_start..utf8_end)
                .ok_or_else(text_unsupported)?
                .to_string(),
            bounds,
        });
        previous_end = end;
    }

    let layer = NativeOfficePdfPageTextLayer {
        schema_version: NATIVE_OFFICE_PDF_TEXT_SCHEMA_VERSION,
        source_revision,
        unit: page_geometry.unit.clone(),
        page_geometry,
        engine_version: PDFIUM_ENGINE_VERSION.to_string(),
        max_characters: options.max_characters,
        max_runs: options.max_runs,
        max_text_bytes: options.max_text_bytes,
        page_objects,
        visual_objects,
        text_sha256: format!("{:x}", Sha256::digest(text.as_bytes())),
        text,
        characters,
        runs,
    };
    Ok(layer)
}

fn pdfium_segment_range(
    indices: impl IntoIterator<Item = usize>,
    character_count: usize,
    previous_end: usize,
) -> UseResult<Option<Range<usize>>> {
    let mut indices = indices.into_iter();
    let Some(start) = indices.next() else {
        return Ok(None);
    };
    let mut end = start;
    for index in indices {
        let expected = end.checked_add(1).ok_or_else(text_unsupported)?;
        if index != expected {
            return Ok(None);
        }
        end = index;
    }

    // pdfium-render 0.9.3 exposes one look-ahead character after each
    // PdfPageTextSegment. Its index is the segment's exclusive end, and the
    // final page segment therefore reports character_count itself.
    if start >= end || end > character_count {
        return Ok(None);
    }
    // Adjacent PDFium rectangles can resolve both endpoints to a character
    // already consumed by the previous rectangle. It carries no new source
    // range and must not duplicate text or geometry in the receipt.
    if end <= previous_end {
        return Ok(None);
    }
    if start < previous_end {
        return Ok(None);
    }
    Ok(Some(start..end))
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

fn non_empty_font_name(value: String) -> Option<String> {
    (!value.is_empty() && value.len() <= 1_024 && !value.chars().any(char::is_control))
        .then_some(value)
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

#[cfg(test)]
mod tests {
    use super::{
        pdfium_segment_range, NativeOfficePdfTextLayerOptions, DEFAULT_NATIVE_OFFICE_PDF_TEXT_RUNS,
    };

    #[test]
    fn text_options_have_a_strict_bounded_run_wire_contract() {
        let options = NativeOfficePdfTextLayerOptions::default();
        options.validate().unwrap();
        let value = serde_json::to_value(options).unwrap();
        assert_eq!(value["maxRuns"], DEFAULT_NATIVE_OFFICE_PDF_TEXT_RUNS);
        assert_eq!(
            serde_json::from_value::<NativeOfficePdfTextLayerOptions>(value.clone()).unwrap(),
            options
        );

        let mut unknown = value;
        unknown["unbounded"] = serde_json::json!(true);
        assert!(serde_json::from_value::<NativeOfficePdfTextLayerOptions>(unknown).is_err());

        let mut invalid = options;
        invalid.max_runs = 0;
        assert_eq!(
            invalid.validate().unwrap_err().code,
            "use.office.pdf_text_options_invalid"
        );
    }

    #[test]
    fn pdfium_segment_ranges_consume_the_exclusive_look_ahead_index() {
        assert_eq!(pdfium_segment_range(0..=5, 9, 0).unwrap(), Some(0..5));
        assert_eq!(pdfium_segment_range(5..=9, 9, 5).unwrap(), Some(5..9));
        assert_eq!(pdfium_segment_range(4..=5, 9, 5).unwrap(), None);
        assert_eq!(
            pdfium_segment_range(std::iter::empty(), 0, 0).unwrap(),
            None
        );
    }

    #[test]
    fn pdfium_segment_ranges_skip_ambiguous_or_unrepresentable_rectangles() {
        for indices in [vec![0], vec![0, 2], vec![4, 5, 6], vec![8, 9, 10]] {
            assert_eq!(pdfium_segment_range(indices, 9, 5).unwrap(), None);
        }
        assert_eq!(
            pdfium_segment_range([usize::MAX, 0], usize::MAX, 0)
                .unwrap_err()
                .code,
            "use.office.pdf_text_unsupported"
        );
    }
}
