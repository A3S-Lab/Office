use std::collections::BTreeMap;

use a3s_use_core::{UseError, UseResult};

use super::{transform_text_elements, CompiledTextReplacement, ReplacementAccumulator};
use crate::editor::{editor_error, node_not_found, MAX_NATIVE_OFFICE_TEXT_SCOPE_CELLS};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::{CellRange, CellReference};
use crate::xml_edit::{
    apply_patches, decoded_element_text, element_fragment, index_xml, patch_start_tag_attributes,
    replace_element_text_patch, IndexedXmlElement, XmlPatch,
};
use crate::{LosslessXmlPart, NativeOfficePackage};

const SHARED_STRINGS_PART: &str = "xl/sharedStrings.xml";

pub(super) fn replace(
    package: &mut NativeOfficePackage,
    path: &str,
    compiled: &CompiledTextReplacement,
    accumulator: &mut ReplacementAccumulator,
) -> UseResult<()> {
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let worksheets = worksheet_parts(&snapshot)?;
    let scope = SpreadsheetScope::parse(path, &snapshot)?;
    let usage = collect_shared_usage(package, &worksheets, &scope)?;
    if usage.selected_cells > MAX_NATIVE_OFFICE_TEXT_SCOPE_CELLS {
        return Err(editor_error(
            "use.office.text_scope_too_large",
            format!(
                "Native Spreadsheet text replacement scopes cannot contain more than {MAX_NATIVE_OFFICE_TEXT_SCOPE_CELLS} existing cells."
            ),
        )
        .with_detail("cells", usage.selected_cells));
    }

    let shared = replace_shared_strings(package, &usage, compiled, accumulator)?;
    for part_name in worksheets {
        replace_worksheet(
            package,
            &part_name,
            &scope,
            &shared.clone_indices,
            compiled,
            accumulator,
        )?;
    }
    Ok(())
}

#[derive(Debug)]
struct SpreadsheetScope {
    worksheet_part: Option<String>,
    range: Option<CellRange>,
}

impl SpreadsheetScope {
    fn parse(path: &str, snapshot: &NativeOfficeDocument) -> UseResult<Self> {
        if path == "/" {
            return Ok(Self {
                worksheet_part: None,
                range: None,
            });
        }
        let requested = path.trim_start_matches('/');
        let (sheet_name, reference) = requested
            .split_once('/')
            .map_or((requested, None), |(sheet, reference)| {
                (sheet, Some(reference))
            });
        if sheet_name.is_empty() {
            return Err(node_not_found(path));
        }
        let sheet = snapshot
            .root()
            .children
            .iter()
            .find(|node| {
                node.node_type == OfficeNodeType::Worksheet
                    && node.path[1..].eq_ignore_ascii_case(sheet_name)
            })
            .ok_or_else(|| node_not_found(path))?;
        let part = sheet.format.get("part").ok_or_else(|| {
            editor_error(
                "use.office.text_scope_invalid",
                format!("Spreadsheet scope '{path}' has no source worksheet part."),
            )
        })?;
        let range = reference
            .map(|reference| {
                if reference.is_empty() {
                    return Err(node_not_found(path));
                }
                let range = CellRange::parse(reference)?;
                let cells = range.cell_count()?;
                if cells > MAX_NATIVE_OFFICE_TEXT_SCOPE_CELLS {
                    return Err(editor_error(
                        "use.office.text_scope_too_large",
                        format!(
                            "Native Spreadsheet text replacement ranges cannot exceed {MAX_NATIVE_OFFICE_TEXT_SCOPE_CELLS} cells."
                        ),
                    )
                    .with_detail("cells", cells));
                }
                Ok(range)
            })
            .transpose()?;
        Ok(Self {
            worksheet_part: Some(part.trim_start_matches('/').to_string()),
            range,
        })
    }

    fn selects(&self, part: &str, reference: CellReference) -> bool {
        self.worksheet_part
            .as_deref()
            .is_none_or(|selected| selected == part)
            && self.range.is_none_or(|range| range.contains(reference))
    }
}

#[derive(Debug, Default)]
struct SharedUsage {
    total: BTreeMap<usize, usize>,
    selected: BTreeMap<usize, usize>,
    selected_cells: usize,
}

fn collect_shared_usage(
    package: &NativeOfficePackage,
    worksheets: &[String],
    scope: &SpreadsheetScope,
) -> UseResult<SharedUsage> {
    let mut usage = SharedUsage::default();
    for part_name in worksheets {
        let part = package.xml_part(part_name)?;
        let root = index_xml(&part)?;
        for (reference, cell) in worksheet_cells(&root)? {
            let selected = scope.selects(part_name, reference);
            if selected {
                usage.selected_cells = usage.selected_cells.checked_add(1).ok_or_else(|| {
                    spreadsheet_error("Spreadsheet selected-cell count overflowed.")
                })?;
            }
            if cell.attributes.get("t").map(String::as_str) != Some("s") {
                continue;
            }
            let index = shared_string_index(&part, cell)?;
            increment(&mut usage.total, index)?;
            if selected {
                increment(&mut usage.selected, index)?;
            }
        }
    }
    Ok(usage)
}

fn increment(counts: &mut BTreeMap<usize, usize>, index: usize) -> UseResult<()> {
    let count = counts.entry(index).or_default();
    *count = count
        .checked_add(1)
        .ok_or_else(|| spreadsheet_error("Spreadsheet shared-string usage count overflowed."))?;
    Ok(())
}

#[derive(Debug, Default)]
struct SharedReplacement {
    clone_indices: BTreeMap<usize, usize>,
}

fn replace_shared_strings(
    package: &mut NativeOfficePackage,
    usage: &SharedUsage,
    compiled: &CompiledTextReplacement,
    accumulator: &mut ReplacementAccumulator,
) -> UseResult<SharedReplacement> {
    if usage.selected.is_empty() {
        return Ok(SharedReplacement::default());
    }
    if !package.contains_part(SHARED_STRINGS_PART) {
        return Err(spreadsheet_error(
            "Spreadsheet cells reference shared strings, but xl/sharedStrings.xml is missing.",
        ));
    }
    let part = package.xml_part(SHARED_STRINGS_PART)?;
    let root = index_xml(&part)?;
    let items = direct_children(&root, "si");
    for index in usage.total.keys() {
        if items.get(*index).is_none() {
            return Err(spreadsheet_error(format!(
                "Spreadsheet cell references missing shared-string index {index}."
            ))
            .with_detail("sharedStringIndex", *index));
        }
    }

    let mut original_patches = Vec::new();
    let mut clone_source_patches = Vec::new();
    let mut changed_partial = Vec::new();
    for (index, selected_count) in &usage.selected {
        let item = items.get(*index).copied().ok_or_else(|| {
            spreadsheet_error(format!(
                "Spreadsheet cell references missing shared-string index {index}."
            ))
        })?;
        let mut text_elements = Vec::new();
        collect_spreadsheet_text(item, &mut text_elements);
        let total_count = usage.total.get(index).copied().unwrap_or(*selected_count);
        let target_patches = if *selected_count == total_count {
            &mut original_patches
        } else {
            &mut clone_source_patches
        };
        let transform = transform_text_elements(
            &part,
            &text_elements,
            compiled,
            accumulator,
            *selected_count,
            target_patches,
        )?;
        if *selected_count < total_count && transform.changed {
            changed_partial.push(*index);
        }
    }

    let mut clone_indices = BTreeMap::new();
    let mut clone_fragments = Vec::new();
    if !changed_partial.is_empty() {
        let clone_source = apply_patches(&part, clone_source_patches)?;
        let clone_part = LosslessXmlPart::parse(SHARED_STRINGS_PART.to_string(), clone_source)?;
        let clone_root = index_xml(&clone_part)?;
        let clone_items = direct_children(&clone_root, "si");
        for index in changed_partial {
            let item = clone_items.get(index).copied().ok_or_else(|| {
                spreadsheet_error(format!(
                    "Spreadsheet shared-string clone source is missing index {index}."
                ))
            })?;
            let new_index = items
                .len()
                .checked_add(clone_indices.len())
                .ok_or_else(|| spreadsheet_error("Spreadsheet shared-string index overflowed."))?;
            clone_fragments.extend_from_slice(element_fragment(&clone_part, item)?);
            clone_indices.insert(index, new_index);
        }
    }

    if !clone_fragments.is_empty() {
        original_patches.push(XmlPatch::new(
            root.content_range.end..root.content_range.end,
            clone_fragments,
        ));
    }
    if !original_patches.is_empty() {
        let edited = apply_patches(&part, original_patches)?;
        let edited = if clone_indices.is_empty() {
            edited
        } else {
            remove_unique_count(edited)?
        };
        package.set_part(SHARED_STRINGS_PART, edited)?;
        accumulator.changed(SHARED_STRINGS_PART);
    }
    Ok(SharedReplacement { clone_indices })
}

fn remove_unique_count(bytes: Vec<u8>) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(SHARED_STRINGS_PART.to_string(), bytes)?;
    let root = index_xml(&part)?;
    if !root.qualified_attributes.contains_key("uniqueCount") {
        return Ok(part.raw().to_vec());
    }
    patch_start_tag_attributes(
        &part,
        &root,
        &BTreeMap::from([("uniqueCount".to_string(), None)]),
    )
}

fn replace_worksheet(
    package: &mut NativeOfficePackage,
    part_name: &str,
    scope: &SpreadsheetScope,
    clone_indices: &BTreeMap<usize, usize>,
    compiled: &CompiledTextReplacement,
    accumulator: &mut ReplacementAccumulator,
) -> UseResult<()> {
    let part = package.xml_part(part_name)?;
    let root = index_xml(&part)?;
    let mut patches = Vec::new();
    for (reference, cell) in worksheet_cells(&root)? {
        if !scope.selects(part_name, reference) {
            continue;
        }
        match cell.attributes.get("t").map_or("n", String::as_str) {
            "s" => {
                let values = direct_children(cell, "v");
                if values.len() != 1 {
                    return Err(spreadsheet_error(format!(
                        "Shared-string cell '{}' must contain exactly one direct value element.",
                        reference.a1()
                    )));
                }
                let index = shared_string_index(&part, cell)?;
                if let Some(new_index) = clone_indices.get(&index) {
                    patches.push(replace_element_text_patch(
                        values[0],
                        &new_index.to_string(),
                    ));
                }
            }
            "inlineStr" => {
                let inline = direct_children(cell, "is");
                ensure_at_most_one(reference, "is", inline.len())?;
                if let Some(inline) = inline.first() {
                    let mut text_elements = Vec::new();
                    collect_spreadsheet_text(inline, &mut text_elements);
                    transform_text_elements(
                        &part,
                        &text_elements,
                        compiled,
                        accumulator,
                        1,
                        &mut patches,
                    )?;
                }
            }
            "str" => {
                let values = direct_children(cell, "v");
                ensure_at_most_one(reference, "v", values.len())?;
                transform_text_elements(&part, &values, compiled, accumulator, 1, &mut patches)?;
            }
            _ => {}
        }
    }
    if !patches.is_empty() {
        package.set_part(part_name, apply_patches(&part, patches)?)?;
        accumulator.changed(part_name);
    }
    Ok(())
}

fn worksheet_parts(document: &NativeOfficeDocument) -> UseResult<Vec<String>> {
    document
        .root()
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Worksheet)
        .map(|node| {
            node.format
                .get("part")
                .map(|part| part.trim_start_matches('/').to_string())
                .ok_or_else(|| {
                    editor_error(
                        "use.office.text_scope_invalid",
                        format!("Spreadsheet worksheet '{}' has no source part.", node.path),
                    )
                })
        })
        .collect()
}

fn worksheet_cells(
    root: &IndexedXmlElement,
) -> UseResult<Vec<(CellReference, &IndexedXmlElement)>> {
    let sheet_data = direct_children(root, "sheetData");
    if sheet_data.len() > 1 {
        return Err(spreadsheet_error(
            "Spreadsheet worksheet contains more than one sheetData element.",
        ));
    }
    let Some(sheet_data) = sheet_data.first().copied() else {
        return Ok(Vec::new());
    };
    let mut cells = Vec::new();
    let mut inferred_row = 0_u32;
    for row in direct_children(sheet_data, "row") {
        let row_number = row
            .attributes
            .get("r")
            .map(|value| value.parse::<u32>())
            .transpose()
            .map_err(|error| spreadsheet_error(format!("Spreadsheet row is invalid: {error}")))?
            .unwrap_or_else(|| inferred_row.saturating_add(1));
        if row_number == 0 {
            return Err(spreadsheet_error("Spreadsheet row zero is invalid."));
        }
        inferred_row = row_number;
        let mut inferred_column = 0_u32;
        for cell in direct_children(row, "c") {
            let reference = if let Some(reference) = cell.attributes.get("r") {
                CellReference::parse(reference)?
            } else {
                inferred_column = inferred_column
                    .checked_add(1)
                    .ok_or_else(|| spreadsheet_error("Spreadsheet inferred column overflowed."))?;
                CellReference {
                    column: inferred_column,
                    row: row_number,
                }
            };
            inferred_column = reference.column;
            cells.push((reference, cell));
        }
    }
    Ok(cells)
}

fn shared_string_index(part: &LosslessXmlPart, cell: &IndexedXmlElement) -> UseResult<usize> {
    let values = direct_children(cell, "v");
    if values.len() != 1 {
        return Err(spreadsheet_error(
            "A shared-string cell must contain exactly one direct value element.",
        ));
    }
    let raw = decoded_element_text(part, values[0])?;
    raw.trim().parse::<usize>().map_err(|error| {
        spreadsheet_error(format!(
            "Spreadsheet shared-string index '{raw}' is invalid: {error}"
        ))
    })
}

fn collect_spreadsheet_text<'a>(
    element: &'a IndexedXmlElement,
    output: &mut Vec<&'a IndexedXmlElement>,
) {
    for child in &element.children {
        if child.namespace != element.namespace {
            continue;
        }
        match child.local_name.as_str() {
            "t" => output.push(child),
            "r" => collect_spreadsheet_text(child, output),
            "rPh" => {}
            _ => {}
        }
    }
}

fn direct_children<'a>(
    element: &'a IndexedXmlElement,
    local_name: &str,
) -> Vec<&'a IndexedXmlElement> {
    element
        .children
        .iter()
        .filter(|child| child.local_name == local_name && child.namespace == element.namespace)
        .collect()
}

fn ensure_at_most_one(reference: CellReference, local_name: &str, count: usize) -> UseResult<()> {
    if count <= 1 {
        return Ok(());
    }
    Err(spreadsheet_error(format!(
        "Spreadsheet cell '{}' contains more than one direct '{local_name}' element.",
        reference.a1()
    )))
}

fn spreadsheet_error(message: impl Into<String>) -> UseError {
    editor_error("use.office.text_spreadsheet_invalid", message)
}
