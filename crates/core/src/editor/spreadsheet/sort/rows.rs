use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::UseResult;

use super::super::{editor_error, indexed_cells_in_row, indexed_rows};
use crate::spreadsheet_reference::{CellRange, CellReference};
use crate::xml_edit::{apply_patches, element_fragment, IndexedXmlElement, XmlPatch};
use crate::LosslessXmlPart;

pub(super) fn rebuild(
    part: &LosslessXmlPart,
    sheet_data: &IndexedXmlElement,
    data_range: CellRange,
    old_to_new: &BTreeMap<u32, u32>,
) -> UseResult<Vec<u8>> {
    let rows = indexed_rows(sheet_data);
    validate_rows(&rows)?;
    let mut row_by_number = BTreeMap::new();
    for (number, row) in &rows {
        if row_by_number.insert(*number, *row).is_some() {
            return Err(row_error(format!(
                "Worksheet contains duplicate row {number}."
            )));
        }
        validate_cells(*number, row)?;
    }
    let new_to_old = old_to_new
        .iter()
        .map(|(old, new)| (*new, *old))
        .collect::<BTreeMap<_, _>>();
    if new_to_old.len() != old_to_new.len() {
        return Err(row_error(
            "Spreadsheet sort row permutation is not one-to-one.",
        ));
    }

    let mut replacements = BTreeMap::<u32, Vec<u8>>::new();
    let mut final_rows = row_by_number.keys().copied().collect::<BTreeSet<_>>();
    for target in data_range.start.row..=data_range.end.row {
        let source = *new_to_old.get(&target).ok_or_else(|| {
            row_error(format!(
                "Spreadsheet sort permutation has no source row for target {target}."
            ))
        })?;
        if let Some(fragment) = build_row(
            part,
            sheet_data,
            row_by_number.get(&target).copied(),
            row_by_number.get(&source).copied(),
            target,
            source,
            data_range,
        )? {
            replacements.insert(target, fragment);
            final_rows.insert(target);
        }
    }

    let bytes = part.parse_bytes();
    let mut prefixes = BTreeMap::<u32, Vec<u8>>::new();
    let mut cursor = sheet_data.content_range.start;
    for (number, row) in &rows {
        let prefix = bytes
            .get(cursor..row.full_range.start)
            .ok_or_else(|| row_error("Spreadsheet row byte ranges are invalid."))?;
        prefixes.insert(*number, prefix.to_vec());
        cursor = row.full_range.end;
    }
    let trailing = bytes
        .get(cursor..sheet_data.content_range.end)
        .ok_or_else(|| row_error("Spreadsheet sheetData byte range is invalid."))?;

    let mut content = Vec::new();
    if rows.is_empty() {
        content.extend_from_slice(trailing);
    }
    for number in final_rows {
        if let Some(prefix) = prefixes.get(&number) {
            content.extend_from_slice(prefix);
        }
        if let Some(replacement) = replacements.get(&number) {
            content.extend_from_slice(replacement);
        } else if let Some(row) = row_by_number.get(&number) {
            content.extend_from_slice(element_fragment(part, row)?);
        }
    }
    if !rows.is_empty() {
        content.extend_from_slice(trailing);
    }
    apply_patches(
        part,
        vec![XmlPatch::new(sheet_data.content_range.clone(), content)],
    )
}

fn build_row(
    part: &LosslessXmlPart,
    sheet_data: &IndexedXmlElement,
    destination: Option<&IndexedXmlElement>,
    source: Option<&IndexedXmlElement>,
    target_row: u32,
    source_row: u32,
    data_range: CellRange,
) -> UseResult<Option<Vec<u8>>> {
    if target_row == source_row {
        return destination
            .map(|row| element_fragment(part, row).map(ToOwned::to_owned))
            .transpose();
    }
    let destination_cells = destination
        .map(|row| cell_map(target_row, row))
        .transpose()?
        .unwrap_or_default();
    let source_cells = source
        .map(|row| cell_map(source_row, row))
        .transpose()?
        .unwrap_or_default();
    let destination_inside = destination_cells
        .range(data_range.start.column..=data_range.end.column)
        .next()
        .is_some();
    let source_inside = source_cells
        .range(data_range.start.column..=data_range.end.column)
        .next()
        .is_some();
    if !destination_inside && !source_inside {
        return destination
            .map(|row| element_fragment(part, row).map(ToOwned::to_owned))
            .transpose();
    }

    let mut cells = BTreeMap::<u32, Vec<u8>>::new();
    for (column, cell) in destination_cells {
        if column < data_range.start.column || column > data_range.end.column {
            cells.insert(column, element_fragment(part, cell)?.to_vec());
        }
    }
    for (column, cell) in source_cells {
        if (data_range.start.column..=data_range.end.column).contains(&column) {
            cells.insert(column, renumbered_cell(part, cell, source_row, target_row)?);
        }
    }

    let non_cells = destination
        .map(|row| non_cell_fragments(part, row))
        .transpose()?
        .unwrap_or_default();
    if destination.is_none() && cells.is_empty() && non_cells.is_empty() {
        return Ok(None);
    }
    let qualified_name = destination
        .or(source)
        .map(|row| row.qualified_name.clone())
        .unwrap_or_else(|| {
            let prefix = sheet_data
                .qualified_name
                .rsplit_once(':')
                .map(|(prefix, _)| prefix);
            prefix.map_or_else(|| "row".to_string(), |prefix| format!("{prefix}:row"))
        });
    let mut attributes = destination
        .map(|row| row.qualified_attributes.clone())
        .unwrap_or_default();
    attributes.insert("r".into(), target_row.to_string());
    let start = start_tag(&qualified_name, attributes, false);
    let mut fragment = start.into_bytes();
    for cell in cells.into_values() {
        fragment.extend_from_slice(&cell);
    }
    for child in non_cells {
        fragment.extend_from_slice(&child);
    }
    fragment.extend_from_slice(format!("</{qualified_name}>").as_bytes());
    Ok(Some(fragment))
}

fn validate_rows(rows: &[(u32, &IndexedXmlElement)]) -> UseResult<()> {
    let mut previous = 0_u32;
    for (number, row) in rows {
        if *number <= previous
            || row
                .attributes
                .get("r")
                .and_then(|value| value.parse::<u32>().ok())
                != Some(*number)
        {
            return Err(row_error(
                "Spreadsheet sorting requires unique ascending rows with explicit canonical r attributes.",
            ));
        }
        previous = *number;
    }
    Ok(())
}

fn validate_cells(row_number: u32, row: &IndexedXmlElement) -> UseResult<()> {
    let mut previous = 0_u32;
    let mut saw_non_cell = false;
    for child in &row.children {
        if child.local_name != "c" {
            saw_non_cell = true;
            continue;
        }
        if saw_non_cell {
            return Err(row_error(format!(
                "Spreadsheet row {row_number} contains a cell after non-cell content."
            )));
        }
        let reference = child
            .attributes
            .get("r")
            .ok_or_else(|| {
                row_error(format!(
                    "Spreadsheet row {row_number} has a cell without r."
                ))
            })
            .and_then(|reference| CellReference::parse(reference))?;
        if reference.row != row_number || reference.column <= previous {
            return Err(row_error(format!(
                "Spreadsheet row {row_number} contains non-canonical or duplicate cell references."
            )));
        }
        previous = reference.column;
    }
    Ok(())
}

fn cell_map(
    row_number: u32,
    row: &IndexedXmlElement,
) -> UseResult<BTreeMap<u32, &IndexedXmlElement>> {
    validate_cells(row_number, row)?;
    Ok(indexed_cells_in_row(row_number, row)
        .into_iter()
        .map(|(reference, cell)| (reference.column, cell))
        .collect())
}

fn non_cell_fragments(part: &LosslessXmlPart, row: &IndexedXmlElement) -> UseResult<Vec<Vec<u8>>> {
    row.children
        .iter()
        .filter(|child| child.local_name != "c")
        .map(|child| element_fragment(part, child).map(ToOwned::to_owned))
        .collect()
}

fn renumbered_cell(
    part: &LosslessXmlPart,
    cell: &IndexedXmlElement,
    source_row: u32,
    target_row: u32,
) -> UseResult<Vec<u8>> {
    let reference = cell
        .attributes
        .get("r")
        .ok_or_else(|| row_error("Spreadsheet sort source cell has no r attribute."))
        .and_then(|reference| CellReference::parse(reference))?;
    if reference.row != source_row {
        return Err(row_error(
            "Spreadsheet sort source cell row does not match its row element.",
        ));
    }
    let mut attributes = cell.qualified_attributes.clone();
    attributes.insert(
        "r".into(),
        CellReference {
            column: reference.column,
            row: target_row,
        }
        .a1(),
    );
    let start = start_tag(&cell.qualified_name, attributes, cell.empty);
    if cell.empty {
        return Ok(start.into_bytes());
    }
    let bytes = part.parse_bytes();
    let content = bytes
        .get(cell.content_range.clone())
        .ok_or_else(|| row_error("Spreadsheet sort cell content range is invalid."))?;
    let closing = bytes
        .get(cell.content_range.end..cell.full_range.end)
        .ok_or_else(|| row_error("Spreadsheet sort cell closing range is invalid."))?;
    let mut fragment = start.into_bytes();
    fragment.extend_from_slice(content);
    fragment.extend_from_slice(closing);
    Ok(fragment)
}

fn start_tag(qualified_name: &str, attributes: BTreeMap<String, String>, empty: bool) -> String {
    let attributes = attributes
        .into_iter()
        .map(|(name, value)| format!(" {name}=\"{}\"", quick_xml::escape::escape(&value)))
        .collect::<String>();
    let terminator = if empty { "/>" } else { ">" };
    format!("<{qualified_name}{attributes}{terminator}")
}

fn row_error(message: impl Into<String>) -> a3s_use_core::UseError {
    editor_error("use.office.spreadsheet_sort_rows_unsupported", message)
}
