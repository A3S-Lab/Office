use a3s_use_core::{UseError, UseResult};
use pdfium_render::prelude::{PdfBookmark, PdfDocument};
use serde::{Deserialize, Serialize};

use super::engine::PDFIUM_ENGINE_VERSION;
use super::text::NATIVE_OFFICE_PDF_TEXT_SCHEMA_VERSION;
use super::NativeOfficePdfPageInventory;
use crate::layout::layout_error;
use crate::{NativeOfficeUnit, PackageRevision};

/// Default maximum number of native PDF outline entries.
pub const DEFAULT_NATIVE_OFFICE_PDF_OUTLINE_ENTRIES: usize = 10_000;
/// Hard maximum number of native PDF outline entries.
pub const MAX_NATIVE_OFFICE_PDF_OUTLINE_ENTRIES: usize = 100_000;
/// Default maximum zero-based depth accepted from a PDF outline.
pub const DEFAULT_NATIVE_OFFICE_PDF_OUTLINE_DEPTH: usize = 64;
/// Hard maximum zero-based depth accepted from a PDF outline.
pub const MAX_NATIVE_OFFICE_PDF_OUTLINE_DEPTH: usize = 256;
/// Default maximum UTF-8 bytes accepted for one PDF outline title.
pub const DEFAULT_NATIVE_OFFICE_PDF_OUTLINE_TITLE_BYTES: usize = 16 * 1024;
/// Hard maximum UTF-8 bytes accepted for one PDF outline title.
pub const MAX_NATIVE_OFFICE_PDF_OUTLINE_TITLE_BYTES: usize = 1024 * 1024;

/// Explicit bounds for one native PDF document-outline extraction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfOutlineOptions {
    pub max_entries: usize,
    pub max_depth: usize,
    pub max_title_bytes: usize,
    pub timeout_ms: u64,
}

impl Default for NativeOfficePdfOutlineOptions {
    fn default() -> Self {
        Self {
            max_entries: DEFAULT_NATIVE_OFFICE_PDF_OUTLINE_ENTRIES,
            max_depth: DEFAULT_NATIVE_OFFICE_PDF_OUTLINE_DEPTH,
            max_title_bytes: DEFAULT_NATIVE_OFFICE_PDF_OUTLINE_TITLE_BYTES,
            timeout_ms: 120_000,
        }
    }
}

impl NativeOfficePdfOutlineOptions {
    pub(super) fn validate(&self) -> UseResult<()> {
        if (1..=MAX_NATIVE_OFFICE_PDF_OUTLINE_ENTRIES).contains(&self.max_entries)
            && (1..=MAX_NATIVE_OFFICE_PDF_OUTLINE_DEPTH).contains(&self.max_depth)
            && (1..=MAX_NATIVE_OFFICE_PDF_OUTLINE_TITLE_BYTES).contains(&self.max_title_bytes)
            && (1..=super::super::MAX_LAYOUT_TIMEOUT_MS).contains(&self.timeout_ms)
        {
            return Ok(());
        }
        Err(layout_error(
            "use.office.pdf_outline_options_invalid",
            "PDF outline entry, depth, title-byte, and timeout bounds are invalid.",
        ))
    }
}

/// One native PDF outline item in deterministic prefix order.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfOutlineEntry {
    pub index: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parent_index: Option<u32>,
    pub depth: u16,
    pub title: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub target_unit: Option<NativeOfficeUnit>,
}

/// Complete, bounded native outline for one inventoried PDF source.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfOutline {
    pub schema_version: u32,
    pub source_revision: PackageRevision,
    pub engine_version: String,
    pub max_entries: usize,
    pub max_depth: usize,
    pub max_title_bytes: usize,
    pub entries: Vec<NativeOfficePdfOutlineEntry>,
}

impl NativeOfficePdfOutline {
    /// Validates hierarchy, limits, source identity, and exact page targets
    /// against the admitted complete page inventory.
    pub fn validate(&self, inventory: &NativeOfficePdfPageInventory) -> UseResult<()> {
        inventory.validate()?;
        if self.source_revision != inventory.source_revision {
            return Err(layout_error(
                "use.office.pdf_outline_source_mismatch",
                "The PDF outline belongs to a different immutable source revision.",
            ));
        }
        if self.schema_version != NATIVE_OFFICE_PDF_TEXT_SCHEMA_VERSION
            || self.engine_version != PDFIUM_ENGINE_VERSION
            || !(1..=MAX_NATIVE_OFFICE_PDF_OUTLINE_ENTRIES).contains(&self.max_entries)
            || !(1..=MAX_NATIVE_OFFICE_PDF_OUTLINE_DEPTH).contains(&self.max_depth)
            || !(1..=MAX_NATIVE_OFFICE_PDF_OUTLINE_TITLE_BYTES).contains(&self.max_title_bytes)
            || self.entries.len() > self.max_entries
        {
            return Err(invalid_outline());
        }

        for (offset, entry) in self.entries.iter().enumerate() {
            let index = u32::try_from(offset).map_err(|_| invalid_outline())?;
            if entry.index != index
                || usize::from(entry.depth) > self.max_depth
                || entry.title.len() > self.max_title_bytes
            {
                return Err(invalid_outline());
            }
            match entry.parent_index {
                None if entry.depth == 0 => {}
                Some(parent_index) if parent_index < entry.index => {
                    let parent = self
                        .entries
                        .get(usize::try_from(parent_index).map_err(|_| invalid_outline())?)
                        .ok_or_else(invalid_outline)?;
                    if parent.depth.checked_add(1) != Some(entry.depth) {
                        return Err(invalid_outline());
                    }
                }
                _ => return Err(invalid_outline()),
            }
            if let Some(unit) = &entry.target_unit {
                inventory
                    .validated_page(unit)
                    .map_err(|_| invalid_outline())?;
            }
        }
        Ok(())
    }
}

pub(super) fn extract_outline(
    document: &PdfDocument<'_>,
    inventory: &NativeOfficePdfPageInventory,
    options: NativeOfficePdfOutlineOptions,
) -> UseResult<NativeOfficePdfOutline> {
    options.validate()?;
    let mut entries = Vec::new();
    let mut pending = Vec::new();
    if let Some(root) = document.bookmarks().root() {
        pending.push((root, None, 0_usize));
    }

    while let Some((bookmark, parent_index, depth)) = pending.pop() {
        if entries.len() >= options.max_entries {
            return Err(layout_error(
                "use.office.pdf_outline_entry_limit",
                format!(
                    "PDF outline exceeds the configured {}-entry limit.",
                    options.max_entries
                ),
            ));
        }
        if depth > options.max_depth {
            return Err(layout_error(
                "use.office.pdf_outline_depth_limit",
                format!(
                    "PDF outline exceeds the configured depth limit of {}.",
                    options.max_depth
                ),
            ));
        }

        let title = bookmark.title().unwrap_or_default();
        if title.len() > options.max_title_bytes {
            return Err(layout_error(
                "use.office.pdf_outline_title_limit",
                format!(
                    "PDF outline title exceeds the configured {}-byte limit.",
                    options.max_title_bytes
                ),
            ));
        }
        let index = u32::try_from(entries.len()).map_err(|_| invalid_outline())?;
        let target_unit = bookmark_target(&bookmark, inventory)?;
        entries.push(NativeOfficePdfOutlineEntry {
            index,
            parent_index,
            depth: u16::try_from(depth).map_err(|_| invalid_outline())?,
            title,
            target_unit,
        });

        if let Some(sibling) = bookmark.next_sibling() {
            pending.push((sibling, parent_index, depth));
        }
        if let Some(child) = bookmark.first_child() {
            pending.push((
                child,
                Some(index),
                depth.checked_add(1).ok_or_else(invalid_outline)?,
            ));
        }
    }

    let outline = NativeOfficePdfOutline {
        schema_version: NATIVE_OFFICE_PDF_TEXT_SCHEMA_VERSION,
        source_revision: inventory.source_revision.clone(),
        engine_version: PDFIUM_ENGINE_VERSION.to_string(),
        max_entries: options.max_entries,
        max_depth: options.max_depth,
        max_title_bytes: options.max_title_bytes,
        entries,
    };
    outline.validate(inventory)?;
    Ok(outline)
}

fn bookmark_target(
    bookmark: &PdfBookmark<'_>,
    inventory: &NativeOfficePdfPageInventory,
) -> UseResult<Option<NativeOfficeUnit>> {
    let direct = bookmark
        .destination()
        .and_then(|target| target.page_index().ok());
    let page_index = direct.or_else(|| {
        bookmark.action().and_then(|action| {
            action.as_local_destination_action().and_then(|local| {
                local
                    .destination()
                    .ok()
                    .and_then(|target| target.page_index().ok())
            })
        })
    });
    let Some(page_index) = page_index else {
        return Ok(None);
    };
    let page_index = usize::try_from(page_index).map_err(|_| invalid_outline())?;
    inventory
        .pages
        .get(page_index)
        .map(|page| Some(page.unit.clone()))
        .ok_or_else(invalid_outline)
}

fn invalid_outline() -> UseError {
    layout_error(
        "use.office.pdf_outline_invalid",
        "PDF outline hierarchy, limits, targets, or identity are inconsistent.",
    )
}
