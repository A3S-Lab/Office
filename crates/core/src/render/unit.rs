use std::collections::BTreeSet;

use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use super::{html, render_error, svg, unit_count, validate_render_limit};
use crate::{DocumentKind, DocumentNode, NativeOfficeDocument, OfficeNodeType};

use super::{NativeOfficeRenderFormat, MAX_NATIVE_OFFICE_RENDER_BYTES};

/// Default maximum number of natural Office units returned by one inventory.
pub const DEFAULT_NATIVE_OFFICE_UNIT_INVENTORY_LIMIT: usize = 10_000;
/// Hard maximum number of natural Office units returned by one inventory.
pub const MAX_NATIVE_OFFICE_UNIT_INVENTORY_LIMIT: usize = 100_000;

/// Exact natural-unit selector for one immutable native Office snapshot.
///
/// Worksheet indexes and slide numbers are one-based. A worksheet locator
/// carries both its position and preserved name so callers cannot silently
/// retarget work after a reorder or rename.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "kebab-case", deny_unknown_fields)]
pub enum NativeOfficeUnitLocator {
    Document,
    Worksheet {
        index: u32,
        name: String,
    },
    Slide {
        number: u32,
    },
    #[cfg(feature = "pdfium")]
    Page {
        number: u32,
    },
}

/// One exact natural unit observed in a native Office semantic snapshot.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeUnit {
    /// One-based position among natural units of the same document.
    pub ordinal: u32,
    pub locator: NativeOfficeUnitLocator,
    /// Stable semantic path emitted into HTML/SVG node evidence.
    pub path: String,
}

/// Resource bound for one complete natural-unit inventory.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeUnitInventoryOptions {
    pub max_units: usize,
}

impl Default for NativeOfficeUnitInventoryOptions {
    fn default() -> Self {
        Self {
            max_units: DEFAULT_NATIVE_OFFICE_UNIT_INVENTORY_LIMIT,
        }
    }
}

/// Complete, bounded natural-unit inventory for one native Office snapshot.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeUnitInventory {
    pub kind: DocumentKind,
    pub max_units: usize,
    pub total_units: usize,
    pub units: Vec<NativeOfficeUnit>,
}

/// Format and byte bound for rendering exactly one natural Office unit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeUnitRenderOptions {
    pub format: NativeOfficeRenderFormat,
    pub max_output_bytes: usize,
}

impl Default for NativeOfficeUnitRenderOptions {
    fn default() -> Self {
        Self {
            format: NativeOfficeRenderFormat::Html,
            max_output_bytes: MAX_NATIVE_OFFICE_RENDER_BYTES,
        }
    }
}

/// Deterministic semantic render containing exactly one natural Office unit.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeRenderedUnit {
    pub kind: DocumentKind,
    pub unit: NativeOfficeUnit,
    pub document_unit_count: usize,
    pub format: NativeOfficeRenderFormat,
    pub media_type: String,
    pub content: String,
    pub byte_length: usize,
    pub sha256: String,
}

impl NativeOfficeDocument {
    /// Inventories all natural units under the default explicit bound.
    pub fn unit_inventory(&self) -> UseResult<NativeOfficeUnitInventory> {
        self.inventory_units(NativeOfficeUnitInventoryOptions::default())
    }

    /// Inventories all natural units, failing instead of truncating identity.
    pub fn inventory_units(
        &self,
        options: NativeOfficeUnitInventoryOptions,
    ) -> UseResult<NativeOfficeUnitInventory> {
        validate_inventory_limit(options.max_units)?;
        let total_units = unit_count(self);
        if total_units > options.max_units {
            return Err(render_error(
                "use.office.unit_inventory_too_large",
                format!(
                    "Native Office document contains {total_units} natural units; the inventory limit is {}.",
                    options.max_units
                ),
            )
            .with_suggestion("Raise maxUnits within the supported bound or reject the document.")
            .with_detail("totalUnits", total_units)
            .with_detail("maxUnits", options.max_units));
        }

        let units = collect_units(self)?;
        if units.len() != total_units {
            return Err(invalid_inventory(
                "Natural Office unit counting and identity collection disagree.",
            )
            .with_detail("countedUnits", total_units)
            .with_detail("collectedUnits", units.len()));
        }
        ensure_unique(&units)?;
        Ok(NativeOfficeUnitInventory {
            kind: self.kind(),
            max_units: options.max_units,
            total_units,
            units,
        })
    }

    /// Renders exactly one inventory locator under an explicit output bound.
    pub fn render_unit(
        &self,
        locator: &NativeOfficeUnitLocator,
        options: NativeOfficeUnitRenderOptions,
    ) -> UseResult<NativeOfficeRenderedUnit> {
        validate_render_limit(options.max_output_bytes)?;
        validate_locator(locator)?;
        let resolved = resolve_unit(self, locator)?;
        let content = match options.format {
            NativeOfficeRenderFormat::Html => html::render_unit(
                self,
                resolved.node,
                resolved.unit.ordinal,
                options.max_output_bytes,
            )?,
            NativeOfficeRenderFormat::Svg => svg::render_unit(
                self,
                resolved.node,
                resolved.unit.ordinal,
                options.max_output_bytes,
            )?,
        };
        let byte_length = content.len();
        let sha256 = format!("{:x}", Sha256::digest(content.as_bytes()));
        Ok(NativeOfficeRenderedUnit {
            kind: self.kind(),
            unit: resolved.unit,
            document_unit_count: unit_count(self),
            format: options.format,
            media_type: options.format.media_type().to_string(),
            content,
            byte_length,
            sha256,
        })
    }
}

struct ResolvedUnit<'a> {
    unit: NativeOfficeUnit,
    node: &'a DocumentNode,
}

fn collect_units(document: &NativeOfficeDocument) -> UseResult<Vec<NativeOfficeUnit>> {
    match document.kind() {
        DocumentKind::Word => {
            validate_word_root(document.root())?;
            Ok(vec![NativeOfficeUnit {
                ordinal: 1,
                locator: NativeOfficeUnitLocator::Document,
                path: document.root().path.clone(),
            }])
        }
        DocumentKind::Spreadsheet => document
            .root()
            .children
            .iter()
            .filter(|node| node.node_type == OfficeNodeType::Worksheet)
            .enumerate()
            .map(|(offset, node)| worksheet_unit(offset, node))
            .collect(),
        DocumentKind::Presentation => document
            .root()
            .children
            .iter()
            .filter(|node| node.node_type == OfficeNodeType::Slide)
            .enumerate()
            .map(|(offset, node)| slide_unit(offset, node))
            .collect(),
    }
}

fn resolve_unit<'a>(
    document: &'a NativeOfficeDocument,
    locator: &NativeOfficeUnitLocator,
) -> UseResult<ResolvedUnit<'a>> {
    match (document.kind(), locator) {
        (DocumentKind::Word, NativeOfficeUnitLocator::Document) => {
            validate_word_root(document.root())?;
            Ok(ResolvedUnit {
                unit: NativeOfficeUnit {
                    ordinal: 1,
                    locator: locator.clone(),
                    path: document.root().path.clone(),
                },
                node: document.root(),
            })
        }
        (DocumentKind::Spreadsheet, NativeOfficeUnitLocator::Worksheet { index, name }) => {
            let offset = usize::try_from(index.saturating_sub(1)).map_err(|_| unit_not_found())?;
            let node = document
                .root()
                .children
                .iter()
                .filter(|node| node.node_type == OfficeNodeType::Worksheet)
                .nth(offset)
                .ok_or_else(unit_not_found)?;
            let unit = worksheet_unit(offset, node)?;
            let (actual_index, actual_name) = match &unit.locator {
                NativeOfficeUnitLocator::Worksheet { index, name } => (index, name),
                _ => {
                    return Err(invalid_inventory(
                        "Worksheet resolution produced a non-worksheet identity.",
                    ))
                }
            };
            if actual_index != index || actual_name != name {
                return Err(render_error(
                    "use.office.unit_identity_mismatch",
                    "The requested worksheet index and name do not identify the same observed unit.",
                )
                .with_detail("requestedIndex", *index)
                .with_detail("requestedName", name.clone())
                .with_detail("actualIndex", *actual_index)
                .with_detail("actualName", actual_name.clone())
                .with_detail("actualPath", unit.path.clone()));
            }
            Ok(ResolvedUnit { unit, node })
        }
        (DocumentKind::Presentation, NativeOfficeUnitLocator::Slide { number }) => {
            let offset = usize::try_from(number.saturating_sub(1)).map_err(|_| unit_not_found())?;
            let node = document
                .root()
                .children
                .iter()
                .filter(|node| node.node_type == OfficeNodeType::Slide)
                .nth(offset)
                .ok_or_else(unit_not_found)?;
            let unit = slide_unit(offset, node)?;
            if unit.locator != *locator {
                return Err(render_error(
                    "use.office.unit_identity_mismatch",
                    "The requested slide number conflicts with the observed slide identity.",
                )
                .with_detail("requestedNumber", *number)
                .with_detail("actualPath", unit.path.clone()));
            }
            Ok(ResolvedUnit { unit, node })
        }
        _ => Err(render_error(
            "use.office.unit_kind_mismatch",
            "The requested unit locator is not valid for this Office document kind.",
        )
        .with_detail("documentKind", document_kind_label(document.kind()))
        .with_detail("locatorKind", locator_kind_label(locator))),
    }
}

fn worksheet_unit(offset: usize, node: &DocumentNode) -> UseResult<NativeOfficeUnit> {
    let ordinal = checked_ordinal(offset)?;
    let name = node
        .path
        .strip_prefix('/')
        .filter(|name| !name.is_empty() && !name.contains('/'))
        .ok_or_else(|| {
            invalid_inventory("A worksheet has an invalid semantic path.")
                .with_detail("path", node.path.clone())
        })?;
    Ok(NativeOfficeUnit {
        ordinal,
        locator: NativeOfficeUnitLocator::Worksheet {
            index: ordinal,
            name: name.to_string(),
        },
        path: node.path.clone(),
    })
}

fn slide_unit(offset: usize, node: &DocumentNode) -> UseResult<NativeOfficeUnit> {
    let ordinal = checked_ordinal(offset)?;
    let expected_path = format!("/slide[{ordinal}]");
    if node.path != expected_path {
        return Err(
            invalid_inventory("A slide semantic path conflicts with its observed order.")
                .with_detail("ordinal", ordinal)
                .with_detail("expectedPath", expected_path)
                .with_detail("actualPath", node.path.clone()),
        );
    }
    Ok(NativeOfficeUnit {
        ordinal,
        locator: NativeOfficeUnitLocator::Slide { number: ordinal },
        path: node.path.clone(),
    })
}

fn checked_ordinal(offset: usize) -> UseResult<u32> {
    u32::try_from(offset.saturating_add(1)).map_err(|_| {
        invalid_inventory("Native Office unit order exceeds the supported identity range.")
    })
}

fn validate_word_root(root: &DocumentNode) -> UseResult<()> {
    if root.node_type == OfficeNodeType::Document && root.path == "/" {
        return Ok(());
    }
    Err(
        invalid_inventory("The Word document root does not have the canonical document identity.")
            .with_detail("path", root.path.clone())
            .with_detail("nodeType", root.node_type.label()),
    )
}

fn ensure_unique(units: &[NativeOfficeUnit]) -> UseResult<()> {
    let mut locators = BTreeSet::new();
    let mut paths = BTreeSet::new();
    let mut worksheet_names = BTreeSet::new();
    for unit in units {
        let unique_locator = locators.insert(unit.locator.clone());
        let unique_path = paths.insert(unit.path.clone());
        let unique_name = match &unit.locator {
            NativeOfficeUnitLocator::Worksheet { name, .. } => {
                worksheet_names.insert(name.to_lowercase())
            }
            _ => true,
        };
        if !unique_locator || !unique_path || !unique_name {
            return Err(render_error(
                "use.office.unit_identity_duplicate",
                "Native Office natural-unit inventory contains a duplicate identity.",
            )
            .with_detail("ordinal", unit.ordinal)
            .with_detail("path", unit.path.clone()));
        }
    }
    Ok(())
}

fn validate_locator(locator: &NativeOfficeUnitLocator) -> UseResult<()> {
    let valid = match locator {
        NativeOfficeUnitLocator::Document => true,
        NativeOfficeUnitLocator::Worksheet { index, name } => {
            *index > 0 && !name.is_empty() && !name.contains('/')
        }
        NativeOfficeUnitLocator::Slide { number } => *number > 0,
        #[cfg(feature = "pdfium")]
        NativeOfficeUnitLocator::Page { number } => *number > 0,
    };
    if valid {
        return Ok(());
    }
    Err(render_error(
        "use.office.unit_locator_invalid",
        "Native Office unit positions must be one-based and worksheet names must be non-empty.",
    )
    .with_detail("locatorKind", locator_kind_label(locator)))
}

fn validate_inventory_limit(limit: usize) -> UseResult<()> {
    if (1..=MAX_NATIVE_OFFICE_UNIT_INVENTORY_LIMIT).contains(&limit) {
        return Ok(());
    }
    Err(render_error(
        "use.office.unit_inventory_limit_invalid",
        format!(
            "Native Office unit inventory limit must be between 1 and {MAX_NATIVE_OFFICE_UNIT_INVENTORY_LIMIT}."
        ),
    )
    .with_detail("requestedMaxUnits", limit)
    .with_detail("supportedMaxUnits", MAX_NATIVE_OFFICE_UNIT_INVENTORY_LIMIT))
}

fn invalid_inventory(message: impl Into<String>) -> UseError {
    render_error("use.office.unit_inventory_invalid", message)
}

fn unit_not_found() -> UseError {
    render_error(
        "use.office.unit_not_found",
        "The requested natural unit does not exist in this Office document snapshot.",
    )
}

fn document_kind_label(kind: DocumentKind) -> &'static str {
    match kind {
        DocumentKind::Word => "word",
        DocumentKind::Spreadsheet => "spreadsheet",
        DocumentKind::Presentation => "presentation",
    }
}

fn locator_kind_label(locator: &NativeOfficeUnitLocator) -> &'static str {
    match locator {
        NativeOfficeUnitLocator::Document => "document",
        NativeOfficeUnitLocator::Worksheet { .. } => "worksheet",
        NativeOfficeUnitLocator::Slide { .. } => "slide",
        #[cfg(feature = "pdfium")]
        NativeOfficeUnitLocator::Page { .. } => "page",
    }
}
