mod engine;
mod outline;
mod renderer;
mod text;

use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};

use super::{layout_error, validate_revision, validate_unit, NativeOfficeLayoutSourceKind};
use crate::{NativeOfficeUnit, NativeOfficeUnitLocator, PackageRevision};

pub use outline::{
    NativeOfficePdfOutline, NativeOfficePdfOutlineEntry, NativeOfficePdfOutlineOptions,
    DEFAULT_NATIVE_OFFICE_PDF_OUTLINE_DEPTH, DEFAULT_NATIVE_OFFICE_PDF_OUTLINE_ENTRIES,
    DEFAULT_NATIVE_OFFICE_PDF_OUTLINE_TITLE_BYTES, MAX_NATIVE_OFFICE_PDF_OUTLINE_DEPTH,
    MAX_NATIVE_OFFICE_PDF_OUTLINE_ENTRIES, MAX_NATIVE_OFFICE_PDF_OUTLINE_TITLE_BYTES,
};
pub use renderer::NativeOfficePdfiumLayoutRenderer;
pub use text::{
    NativeOfficePdfPageTextLayer, NativeOfficePdfTextCharacter, NativeOfficePdfTextLayerOptions,
    DEFAULT_NATIVE_OFFICE_PDF_TEXT_CHARACTERS, DEFAULT_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES,
    MAX_NATIVE_OFFICE_PDF_TEXT_CHARACTERS, MAX_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES,
    NATIVE_OFFICE_PDF_TEXT_SCHEMA_VERSION,
};

/// Default maximum number of pages accepted by one PDF inventory.
pub const DEFAULT_NATIVE_OFFICE_PDF_PAGE_LIMIT: usize = 10_000;
/// Hard maximum number of pages accepted by one PDF inventory.
pub const MAX_NATIVE_OFFICE_PDF_PAGE_LIMIT: usize = 100_000;
/// Hard maximum PDF source size admitted by the PDFium layout provider.
pub const MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES: u64 = 512 * 1024 * 1024;

const MIN_PDF_DPI_MILLI: u32 = 1_000;
const MAX_PDF_DPI_MILLI: u32 = 2_400_000;
const MAX_PDF_BOX_ABS_MILLIPOINTS: i64 = 10_000_000_000;

/// One PDF page boundary in thousandths of a PDF point.
///
/// Integer coordinates keep inventory identities deterministic while retaining
/// sub-micrometer precision. One PDF point is 1/72 inch.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfPageBox {
    pub left_millipoints: i64,
    pub bottom_millipoints: i64,
    pub right_millipoints: i64,
    pub top_millipoints: i64,
}

impl NativeOfficePdfPageBox {
    pub(super) fn validate(&self) -> UseResult<()> {
        let coordinates = [
            self.left_millipoints,
            self.bottom_millipoints,
            self.right_millipoints,
            self.top_millipoints,
        ];
        if coordinates
            .into_iter()
            .all(|value| value.abs() <= MAX_PDF_BOX_ABS_MILLIPOINTS)
            && self.right_millipoints > self.left_millipoints
            && self.top_millipoints > self.bottom_millipoints
        {
            return Ok(());
        }
        Err(invalid_page_geometry())
    }

    fn width_millipoints(&self) -> u64 {
        u64::try_from(self.right_millipoints - self.left_millipoints).unwrap_or(0)
    }

    fn height_millipoints(&self) -> u64 {
        u64::try_from(self.top_millipoints - self.bottom_millipoints).unwrap_or(0)
    }

    fn is_within(&self, outer: &Self) -> bool {
        self.left_millipoints >= outer.left_millipoints
            && self.bottom_millipoints >= outer.bottom_millipoints
            && self.right_millipoints <= outer.right_millipoints
            && self.top_millipoints <= outer.top_millipoints
    }
}

/// Stable identity and exact visible geometry of one PDF page.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfPageGeometry {
    pub unit: NativeOfficeUnit,
    pub media_box: NativeOfficePdfPageBox,
    pub crop_box: NativeOfficePdfPageBox,
    pub rotation_degrees: u16,
    pub surface_width_micrometers: u32,
    pub surface_height_micrometers: u32,
    pub output_width_px: u32,
    pub output_height_px: u32,
}

impl NativeOfficePdfPageGeometry {
    fn validate(&self, dpi_milli: u32) -> UseResult<()> {
        validate_unit(&self.unit).map_err(|_| invalid_page_geometry())?;
        self.media_box.validate()?;
        self.crop_box.validate()?;
        if !self.crop_box.is_within(&self.media_box)
            || !matches!(self.rotation_degrees, 0 | 90 | 180 | 270)
        {
            return Err(invalid_page_geometry());
        }

        let (width_millipoints, height_millipoints) = if matches!(self.rotation_degrees, 90 | 270) {
            (
                self.crop_box.height_millipoints(),
                self.crop_box.width_millipoints(),
            )
        } else {
            (
                self.crop_box.width_millipoints(),
                self.crop_box.height_millipoints(),
            )
        };
        let expected_width_micrometers = millipoints_to_micrometers(width_millipoints)?;
        let expected_height_micrometers = millipoints_to_micrometers(height_millipoints)?;
        let expected_width_px = millipoints_to_pixels(width_millipoints, dpi_milli)?;
        let expected_height_px = millipoints_to_pixels(height_millipoints, dpi_milli)?;
        if self.surface_width_micrometers != expected_width_micrometers
            || self.surface_height_micrometers != expected_height_micrometers
            || self.output_width_px != expected_width_px
            || self.output_height_px != expected_height_px
        {
            return Err(invalid_page_geometry());
        }
        Ok(())
    }
}

/// Bounds for one complete, non-truncated PDF page inventory.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfPageInventoryOptions {
    pub max_pages: usize,
    pub max_source_bytes: u64,
    pub dpi_milli: u32,
    pub timeout_ms: u64,
}

impl Default for NativeOfficePdfPageInventoryOptions {
    fn default() -> Self {
        Self {
            max_pages: DEFAULT_NATIVE_OFFICE_PDF_PAGE_LIMIT,
            max_source_bytes: MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES,
            dpi_milli: 144_000,
            timeout_ms: 120_000,
        }
    }
}

impl NativeOfficePdfPageInventoryOptions {
    pub(super) fn validate(&self) -> UseResult<()> {
        if (1..=MAX_NATIVE_OFFICE_PDF_PAGE_LIMIT).contains(&self.max_pages)
            && (1..=MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES).contains(&self.max_source_bytes)
            && (MIN_PDF_DPI_MILLI..=MAX_PDF_DPI_MILLI).contains(&self.dpi_milli)
            && (1..=super::MAX_LAYOUT_TIMEOUT_MS).contains(&self.timeout_ms)
        {
            return Ok(());
        }
        Err(layout_error(
            "use.office.pdf_inventory_options_invalid",
            "PDF inventory page, source-byte, DPI, and timeout bounds are invalid.",
        ))
    }
}

/// Complete, source-bound page inventory emitted by the PDFium provider.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfPageInventory {
    pub kind: NativeOfficeLayoutSourceKind,
    pub source_revision: PackageRevision,
    pub max_pages: usize,
    pub total_pages: usize,
    pub dpi_milli: u32,
    pub pages: Vec<NativeOfficePdfPageGeometry>,
}

impl NativeOfficePdfPageInventory {
    /// Validates completeness, source identity, ordering, and exact page geometry.
    pub fn validate(&self) -> UseResult<()> {
        self.validate_envelope()?;
        for (offset, page) in self.pages.iter().enumerate() {
            let expected_number = u32::try_from(offset + 1).map_err(|_| invalid_inventory())?;
            if page.unit
                != (NativeOfficeUnit {
                    ordinal: expected_number,
                    locator: NativeOfficeUnitLocator::Page {
                        number: expected_number,
                    },
                    path: format!("/page[{expected_number}]"),
                })
            {
                return Err(invalid_inventory());
            }
            page.validate(self.dpi_milli)?;
        }
        Ok(())
    }

    fn validate_envelope(&self) -> UseResult<()> {
        validate_revision(&self.source_revision).map_err(|_| invalid_inventory())?;
        if self.kind != NativeOfficeLayoutSourceKind::Pdf
            || !(1..=MAX_NATIVE_OFFICE_PDF_PAGE_LIMIT).contains(&self.max_pages)
            || !(MIN_PDF_DPI_MILLI..=MAX_PDF_DPI_MILLI).contains(&self.dpi_milli)
            || self.total_pages == 0
            || self.pages.len() != self.total_pages
        {
            return Err(invalid_inventory());
        }
        if self.total_pages > self.max_pages {
            return Err(layout_error(
                "use.office.pdf_page_limit",
                format!(
                    "PDF contains {} pages; the configured limit is {}.",
                    self.total_pages, self.max_pages
                ),
            ));
        }
        Ok(())
    }

    pub(super) fn validated_page(
        &self,
        unit: &NativeOfficeUnit,
    ) -> UseResult<&NativeOfficePdfPageGeometry> {
        self.validate_envelope()?;
        let offset = usize::try_from(unit.ordinal.saturating_sub(1))
            .map_err(|_| pdf_page_identity_mismatch())?;
        let page = self
            .pages
            .get(offset)
            .filter(|page| page.unit == *unit)
            .ok_or_else(pdf_page_identity_mismatch)?;
        page.validate(self.dpi_milli)?;
        Ok(page)
    }

    /// Confirms that an inventory belongs to the caller's immutable source.
    pub fn validate_source(&self, source_revision: &PackageRevision) -> UseResult<()> {
        self.validate()?;
        if &self.source_revision == source_revision {
            return Ok(());
        }
        Err(layout_error(
            "use.office.pdf_inventory_source_mismatch",
            "The PDF page inventory belongs to a different immutable source revision.",
        ))
    }
}

pub(super) fn millipoints_to_micrometers(value: u64) -> UseResult<u32> {
    let rounded = (u128::from(value) * 25_400 + 36_000) / 72_000;
    u32::try_from(rounded).map_err(|_| invalid_page_geometry())
}

pub(super) fn millipoints_to_pixels(value: u64, dpi_milli: u32) -> UseResult<u32> {
    let rounded = (u128::from(value) * u128::from(dpi_milli) + 36_000_000) / 72_000_000;
    u32::try_from(rounded)
        .ok()
        .filter(|pixels| *pixels > 0)
        .ok_or_else(invalid_page_geometry)
}

fn invalid_inventory() -> UseError {
    layout_error(
        "use.office.pdf_inventory_invalid",
        "A PDF page inventory must be complete, ordered, bounded, and source-bound.",
    )
}

pub(super) fn invalid_page_geometry() -> UseError {
    layout_error(
        "use.office.pdf_page_geometry_invalid",
        "PDF media box, crop box, rotation, physical surface, and pixel geometry disagree.",
    )
}

pub(super) fn pdf_page_identity_mismatch() -> UseError {
    layout_error(
        "use.office.pdf_page_identity_mismatch",
        "The requested PDF page locator does not match the observed one-based page identity.",
    )
}
