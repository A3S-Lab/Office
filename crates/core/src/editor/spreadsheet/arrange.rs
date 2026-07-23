use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{
    editor_error, indexed_cells_in_row, indexed_rows, mark_workbook_for_recalculation,
    node_not_found, update_dimension, worksheet,
};
use crate::editor::{NativeOfficeInsertPosition, NativeOfficeSwapResult};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::{CellReference, MAX_ROWS};
use crate::xml_edit::{apply_patches, element_fragment, index_xml, IndexedXmlElement, XmlPatch};
use crate::{LosslessXmlPart, NativeOfficePackage};

pub(in crate::editor) fn move_node(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
) -> UseResult<String> {
    if parse_row_path(path).is_some() {
        return arrange_row(package, path, target_parent, position, RowOperation::Move);
    }
    require_worksheet_path(path)?;
    require_root_parent(target_parent)?;
    let sheets = worksheet_paths(package)?;
    let source = find_path(&sheets, path).ok_or_else(|| node_not_found(path))?;
    let source_index = sheets
        .iter()
        .position(|candidate| candidate.eq_ignore_ascii_case(source))
        .ok_or_else(|| node_not_found(path))?;
    let mut remaining = sheets.clone();
    remaining.remove(source_index);
    let slot = resolve_sheet_slot(
        &remaining,
        source,
        source_index,
        position,
        SheetOperation::Move,
    )?;
    worksheet::move_worksheet(package, source, slot + 1)
}

pub(in crate::editor) fn copy_node(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
    name: Option<&str>,
) -> UseResult<String> {
    if parse_row_path(path).is_some() {
        if name.is_some() {
            return Err(editor_error(
                "use.office.mutation_option_unsupported",
                "Spreadsheet row copy does not accept --name.",
            ));
        }
        return arrange_row(package, path, target_parent, position, RowOperation::Copy);
    }
    require_worksheet_path(path)?;
    require_root_parent(target_parent)?;
    let name = name.ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_copy_name_required",
            "Worksheet copy requires a distinct --name value.",
        )
    })?;
    let sheets = worksheet_paths(package)?;
    let source = find_path(&sheets, path).ok_or_else(|| node_not_found(path))?;
    let source_index = sheets
        .iter()
        .position(|candidate| candidate.eq_ignore_ascii_case(source))
        .ok_or_else(|| node_not_found(path))?;
    let slot = resolve_sheet_slot(
        &sheets,
        source,
        source_index,
        position,
        SheetOperation::Copy,
    )?;
    worksheet::copy_worksheet(package, source, name, Some(slot + 1))
}

pub(in crate::editor) fn swap_nodes(
    package: &mut NativeOfficePackage,
    path: &str,
    with: &str,
) -> UseResult<NativeOfficeSwapResult> {
    match (parse_row_path(path), parse_row_path(with)) {
        (Some(_), Some(_)) => swap_rows(package, path, with),
        (None, None) => swap_worksheets(package, path, with),
        _ => Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Spreadsheet swap requires two worksheets or two rows in the same worksheet.",
        )),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SheetOperation {
    Move,
    Copy,
}

fn resolve_sheet_slot(
    order: &[String],
    source: &str,
    source_index: usize,
    position: Option<&NativeOfficeInsertPosition>,
    operation: SheetOperation,
) -> UseResult<usize> {
    match position {
        Some(NativeOfficeInsertPosition::Index { index }) => {
            if *index > order.len() {
                return Err(editor_error(
                    "use.office.position_invalid",
                    format!(
                        "Worksheet insertion index {index} is outside 0-{}.",
                        order.len()
                    ),
                ));
            }
            Ok(*index)
        }
        Some(NativeOfficeInsertPosition::Before { path }) => {
            sheet_anchor_slot(order, source, source_index, path, false, operation)
        }
        Some(NativeOfficeInsertPosition::After { path }) => {
            sheet_anchor_slot(order, source, source_index, path, true, operation)
        }
        None if operation == SheetOperation::Copy => Ok(source_index + 1),
        None => Ok(order.len()),
    }
}

fn sheet_anchor_slot(
    order: &[String],
    source: &str,
    source_index: usize,
    anchor: &str,
    after: bool,
    operation: SheetOperation,
) -> UseResult<usize> {
    require_worksheet_path(anchor)?;
    if operation == SheetOperation::Move && anchor.eq_ignore_ascii_case(source) {
        return Ok(source_index.min(order.len()));
    }
    let slot = order
        .iter()
        .position(|candidate| candidate.eq_ignore_ascii_case(anchor))
        .ok_or_else(|| node_not_found(anchor))?;
    Ok(slot + usize::from(after))
}

fn swap_worksheets(
    package: &mut NativeOfficePackage,
    path: &str,
    with: &str,
) -> UseResult<NativeOfficeSwapResult> {
    require_worksheet_path(path)?;
    require_worksheet_path(with)?;
    let sheets = worksheet_paths(package)?;
    let first = find_path(&sheets, path).ok_or_else(|| node_not_found(path))?;
    let second = find_path(&sheets, with).ok_or_else(|| node_not_found(with))?;
    if first.eq_ignore_ascii_case(second) {
        return Ok(NativeOfficeSwapResult {
            first: first.to_string(),
            second: second.to_string(),
        });
    }
    let first_position = sheets
        .iter()
        .position(|candidate| candidate.eq_ignore_ascii_case(first))
        .ok_or_else(|| node_not_found(path))?
        + 1;
    let second_position = sheets
        .iter()
        .position(|candidate| candidate.eq_ignore_ascii_case(second))
        .ok_or_else(|| node_not_found(with))?
        + 1;
    worksheet::move_worksheet(package, first, second_position)?;
    worksheet::move_worksheet(package, second, first_position)?;
    Ok(NativeOfficeSwapResult {
        first: first.to_string(),
        second: second.to_string(),
    })
}

fn worksheet_paths(package: &NativeOfficePackage) -> UseResult<Vec<String>> {
    Ok(NativeOfficeDocument::from_package(package.clone())?
        .root()
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Worksheet)
        .map(|node| node.path.clone())
        .collect())
}

fn find_path<'a>(paths: &'a [String], requested: &str) -> Option<&'a str> {
    paths
        .iter()
        .find(|path| path.eq_ignore_ascii_case(requested))
        .map(String::as_str)
}

fn require_worksheet_path(path: &str) -> UseResult<()> {
    if path.starts_with('/') && !path[1..].contains('/') {
        Ok(())
    } else {
        Err(editor_error(
            "use.office.mutation_path_unsupported",
            format!("Spreadsheet worksheet operation requires '/SheetName', received '{path}'."),
        ))
    }
}

fn require_root_parent(target_parent: Option<&str>) -> UseResult<()> {
    if target_parent.is_none_or(|parent| matches!(parent, "/" | "/workbook")) {
        Ok(())
    } else {
        Err(editor_error(
            "use.office.mutation_parent_unsupported",
            "Spreadsheet worksheet move/copy requires the workbook root as its parent.",
        ))
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RowOperation {
    Move,
    Copy,
}

#[derive(Debug, Clone, Copy)]
struct RowToken<'a> {
    number: u32,
    element: &'a IndexedXmlElement,
}

fn arrange_row(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
    operation: RowOperation,
) -> UseResult<String> {
    let (sheet_path, source_number) = parse_row_path(path).ok_or_else(|| node_not_found(path))?;
    require_row_parent(&sheet_path, target_parent)?;
    let (canonical_sheet, part_name) = resolve_sheet_part(package, &sheet_path)?;
    let part = package.xml_part(&part_name)?;
    let index = index_xml(&part)?;
    let sheet_data = index.descendant("sheetData").ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_data_missing",
            "Spreadsheet worksheet has no sheetData element.",
        )
    })?;
    require_plain_dense_rows(package, &part_name, &part, &index, sheet_data)?;
    let rows = indexed_rows(sheet_data);
    let source_index = rows
        .iter()
        .position(|(number, _)| *number == source_number)
        .ok_or_else(|| node_not_found(path))?;
    if operation == RowOperation::Copy {
        require_copyable_row(rows[source_index].1, path)?;
        if rows.len() >= usize::try_from(MAX_ROWS).unwrap_or(usize::MAX) {
            return Err(editor_error(
                "use.office.spreadsheet_structure_overflow",
                "Spreadsheet row copy would exceed the worksheet row limit.",
            ));
        }
    }

    let mut order = rows
        .iter()
        .map(|(number, element)| RowToken {
            number: *number,
            element,
        })
        .collect::<Vec<_>>();
    if operation == RowOperation::Move {
        order.remove(source_index);
    }
    let slot = resolve_row_slot(
        &order,
        &canonical_sheet,
        source_number,
        source_index,
        position,
        operation,
    )?;
    let slot = slot.min(order.len());
    order.insert(
        slot,
        RowToken {
            number: source_number,
            element: rows[source_index].1,
        },
    );
    let edited = rebuild_rows(&part, sheet_data, &rows, &order)?;
    let edited = update_dimension(&part_name, edited)?;
    package.set_part(&part_name, edited)?;
    mark_workbook_for_recalculation(package)?;
    Ok(format!("{canonical_sheet}/row[{}]", slot + 1))
}

fn swap_rows(
    package: &mut NativeOfficePackage,
    path: &str,
    with: &str,
) -> UseResult<NativeOfficeSwapResult> {
    let (first_sheet, first_number) = parse_row_path(path).ok_or_else(|| node_not_found(path))?;
    let (second_sheet, second_number) = parse_row_path(with).ok_or_else(|| node_not_found(with))?;
    if !first_sheet.eq_ignore_ascii_case(&second_sheet) {
        return Err(editor_error(
            "use.office.mutation_parent_unsupported",
            "Spreadsheet row swap requires both rows to be in the same worksheet.",
        ));
    }
    let (canonical_sheet, part_name) = resolve_sheet_part(package, &first_sheet)?;
    let part = package.xml_part(&part_name)?;
    let index = index_xml(&part)?;
    let sheet_data = index.descendant("sheetData").ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_data_missing",
            "Spreadsheet worksheet has no sheetData element.",
        )
    })?;
    require_plain_dense_rows(package, &part_name, &part, &index, sheet_data)?;
    let rows = indexed_rows(sheet_data);
    let first_index = rows
        .iter()
        .position(|(number, _)| *number == first_number)
        .ok_or_else(|| node_not_found(path))?;
    let second_index = rows
        .iter()
        .position(|(number, _)| *number == second_number)
        .ok_or_else(|| node_not_found(with))?;
    if first_index == second_index {
        return Ok(NativeOfficeSwapResult {
            first: format!("{canonical_sheet}/row[{}]", first_index + 1),
            second: format!("{canonical_sheet}/row[{}]", second_index + 1),
        });
    }
    let mut order = rows
        .iter()
        .map(|(number, element)| RowToken {
            number: *number,
            element,
        })
        .collect::<Vec<_>>();
    order.swap(first_index, second_index);
    let edited = rebuild_rows(&part, sheet_data, &rows, &order)?;
    let edited = update_dimension(&part_name, edited)?;
    package.set_part(&part_name, edited)?;
    mark_workbook_for_recalculation(package)?;
    Ok(NativeOfficeSwapResult {
        first: format!("{canonical_sheet}/row[{}]", second_index + 1),
        second: format!("{canonical_sheet}/row[{}]", first_index + 1),
    })
}

fn parse_row_path(path: &str) -> Option<(String, u32)> {
    let (sheet, row) = path.rsplit_once("/row[")?;
    let row = row.strip_suffix(']')?.parse::<u32>().ok()?;
    let sheet_name = sheet.strip_prefix('/')?;
    if sheet_name.is_empty() || row == 0 || sheet_name.contains('/') {
        return None;
    }
    Some((sheet.to_string(), row))
}

fn require_row_parent(sheet: &str, target_parent: Option<&str>) -> UseResult<()> {
    if target_parent.is_none_or(|parent| parent.eq_ignore_ascii_case(sheet)) {
        Ok(())
    } else {
        Err(editor_error(
            "use.office.mutation_parent_unsupported",
            "Native Spreadsheet row move/copy currently requires the source worksheet; cross-sheet reference migration is not yet enabled.",
        ))
    }
}

fn resolve_sheet_part(
    package: &NativeOfficePackage,
    requested: &str,
) -> UseResult<(String, String)> {
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let sheet = snapshot
        .root()
        .children
        .iter()
        .find(|node| {
            node.node_type == OfficeNodeType::Worksheet && node.path.eq_ignore_ascii_case(requested)
        })
        .ok_or_else(|| node_not_found(requested))?;
    let part = sheet.format.get("part").cloned().ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_invalid",
            format!("Worksheet '{}' has no source part.", sheet.path),
        )
    })?;
    Ok((sheet.path.clone(), part))
}

fn resolve_row_slot(
    order: &[RowToken<'_>],
    sheet: &str,
    source_number: u32,
    source_index: usize,
    position: Option<&NativeOfficeInsertPosition>,
    operation: RowOperation,
) -> UseResult<usize> {
    match position {
        Some(NativeOfficeInsertPosition::Index { index }) => {
            if *index > order.len() {
                return Err(editor_error(
                    "use.office.position_invalid",
                    format!(
                        "Spreadsheet row insertion index {index} is outside 0-{}.",
                        order.len()
                    ),
                ));
            }
            Ok(*index)
        }
        Some(NativeOfficeInsertPosition::Before { path }) => row_anchor_slot(
            order,
            sheet,
            source_number,
            source_index,
            path,
            false,
            operation,
        ),
        Some(NativeOfficeInsertPosition::After { path }) => row_anchor_slot(
            order,
            sheet,
            source_number,
            source_index,
            path,
            true,
            operation,
        ),
        None if operation == RowOperation::Copy => Ok(source_index + 1),
        None => Ok(order.len()),
    }
}

fn row_anchor_slot(
    order: &[RowToken<'_>],
    sheet: &str,
    source_number: u32,
    source_index: usize,
    anchor: &str,
    after: bool,
    operation: RowOperation,
) -> UseResult<usize> {
    let (anchor_sheet, anchor_number) =
        parse_row_path(anchor).ok_or_else(|| node_not_found(anchor))?;
    if !anchor_sheet.eq_ignore_ascii_case(sheet) {
        return Err(editor_error(
            "use.office.position_parent_mismatch",
            format!("Spreadsheet row anchor '{anchor}' is not in '{sheet}'."),
        ));
    }
    if operation == RowOperation::Move && anchor_number == source_number {
        return Ok(source_index.min(order.len()));
    }
    let slot = order
        .iter()
        .position(|row| row.number == anchor_number)
        .ok_or_else(|| node_not_found(anchor))?;
    Ok(slot + usize::from(after))
}

fn require_plain_dense_rows(
    package: &NativeOfficePackage,
    part_name: &str,
    part: &LosslessXmlPart,
    index: &IndexedXmlElement,
    sheet_data: &IndexedXmlElement,
) -> UseResult<()> {
    let rows = indexed_rows(sheet_data);
    for (offset, (number, row)) in rows.iter().enumerate() {
        let expected = u32::try_from(offset + 1).map_err(|_| row_arrange_unsupported())?;
        if *number != expected || row.attributes.get("r") != Some(&expected.to_string()) {
            return Err(row_arrange_unsupported());
        }
        for (reference, cell) in indexed_cells_in_row(*number, row) {
            if cell.attributes.get("r") != Some(&reference.a1()) {
                return Err(row_arrange_unsupported());
            }
        }
    }
    require_whitespace_gaps(part, sheet_data, &rows)?;

    let relationship_part = worksheet::relationship_part(part_name);
    if package.contains_part(&relationship_part) {
        let relationships = package.xml_part(&relationship_part)?;
        if index_xml(&relationships)?
            .children
            .iter()
            .any(|child| child.local_name == "Relationship")
        {
            return Err(row_arrange_unsupported());
        }
    }

    let workbook = package.xml_part("xl/workbook.xml")?;
    let workbook_index = index_xml(&workbook)?;
    if workbook_index.descendant("definedName").is_some() {
        return Err(row_arrange_unsupported());
    }
    for candidate in package
        .part_names()
        .filter(|name| name.starts_with("xl/") && name.ends_with(".xml"))
        .map(str::to_string)
        .collect::<Vec<_>>()
    {
        let xml = package.xml_part(&candidate)?;
        let candidate_index = index_xml(&xml)?;
        if contains_formula(&candidate_index) {
            return Err(row_arrange_unsupported());
        }
    }
    if contains_row_reference_metadata(index, sheet_data) {
        return Err(row_arrange_unsupported());
    }
    Ok(())
}

fn contains_formula(element: &IndexedXmlElement) -> bool {
    matches!(
        element.local_name.as_str(),
        "f" | "formula" | "calculatedColumnFormula" | "totalsRowFormula"
    ) || element.children.iter().any(contains_formula)
}

fn contains_row_reference_metadata(
    element: &IndexedXmlElement,
    sheet_data: &IndexedXmlElement,
) -> bool {
    if element.full_range == sheet_data.full_range {
        return false;
    }
    if matches!(
        element.local_name.as_str(),
        "rowBreaks"
            | "mergeCells"
            | "conditionalFormatting"
            | "dataValidations"
            | "autoFilter"
            | "hyperlinks"
            | "sortState"
            | "protectedRanges"
            | "ignoredErrors"
            | "pivotTableDefinition"
    ) {
        return true;
    }
    let has_reference = element.qualified_attributes.keys().any(|attribute| {
        matches!(
            attribute
                .rsplit_once(':')
                .map_or(attribute.as_str(), |(_, local)| local),
            "sqref" | "activeCell" | "topLeftCell"
        ) || (attribute.ends_with("ref") && element.local_name != "dimension")
    });
    has_reference
        || element
            .children
            .iter()
            .any(|child| contains_row_reference_metadata(child, sheet_data))
}

fn require_whitespace_gaps(
    part: &LosslessXmlPart,
    sheet_data: &IndexedXmlElement,
    rows: &[(u32, &IndexedXmlElement)],
) -> UseResult<()> {
    let bytes = part.parse_bytes();
    let mut cursor = sheet_data.content_range.start;
    for (_, row) in rows {
        let gap = bytes
            .get(cursor..row.full_range.start)
            .ok_or_else(row_arrange_unsupported)?;
        if !gap.iter().all(u8::is_ascii_whitespace) {
            return Err(row_arrange_unsupported());
        }
        cursor = row.full_range.end;
    }
    let trailing = bytes
        .get(cursor..sheet_data.content_range.end)
        .ok_or_else(row_arrange_unsupported)?;
    if trailing.iter().all(u8::is_ascii_whitespace) {
        Ok(())
    } else {
        Err(row_arrange_unsupported())
    }
}

fn require_copyable_row(row: &IndexedXmlElement, path: &str) -> UseResult<()> {
    if row_contains_unsafe_copy_data(row) {
        Err(editor_error(
            "use.office.spreadsheet_row_copy_unsupported",
            format!(
                "Spreadsheet row '{path}' contains shared strings, metadata identities, extensions, or relationships that cannot yet be cloned losslessly."
            ),
        ))
    } else {
        Ok(())
    }
}

fn row_contains_unsafe_copy_data(element: &IndexedXmlElement) -> bool {
    element.local_name == "extLst"
        || (element.local_name == "c"
            && element.attributes.get("t").map(String::as_str) == Some("s"))
        || element.qualified_attributes.keys().any(|attribute| {
            matches!(
                attribute
                    .rsplit_once(':')
                    .map_or(attribute.as_str(), |(_, local)| local),
                "cm" | "vm" | "id" | "embed" | "link"
            )
        })
        || element.children.iter().any(row_contains_unsafe_copy_data)
}

fn rebuild_rows(
    part: &LosslessXmlPart,
    sheet_data: &IndexedXmlElement,
    original_rows: &[(u32, &IndexedXmlElement)],
    order: &[RowToken<'_>],
) -> UseResult<Vec<u8>> {
    let bytes = part.parse_bytes();
    let leading_end = original_rows
        .first()
        .map_or(sheet_data.content_range.end, |(_, row)| {
            row.full_range.start
        });
    let leading = bytes
        .get(sheet_data.content_range.start..leading_end)
        .ok_or_else(row_arrange_unsupported)?;
    let trailing_start = original_rows
        .last()
        .map_or(sheet_data.content_range.start, |(_, row)| {
            row.full_range.end
        });
    let trailing = bytes
        .get(trailing_start..sheet_data.content_range.end)
        .ok_or_else(row_arrange_unsupported)?;
    let gaps = original_rows
        .windows(2)
        .map(|pair| {
            bytes
                .get(pair[0].1.full_range.end..pair[1].1.full_range.start)
                .ok_or_else(row_arrange_unsupported)
        })
        .collect::<UseResult<Vec<_>>>()?;

    let mut content = Vec::new();
    content.extend_from_slice(leading);
    for (offset, token) in order.iter().enumerate() {
        let number = u32::try_from(offset + 1).map_err(|_| row_arrange_unsupported())?;
        content.extend_from_slice(&renumbered_row(part, token.element, token.number, number)?);
        if offset + 1 < order.len() {
            if let Some(gap) = gaps.get(offset) {
                content.extend_from_slice(gap);
            }
        }
    }
    content.extend_from_slice(trailing);
    apply_patches(
        part,
        vec![XmlPatch::new(sheet_data.content_range.clone(), content)],
    )
}

fn renumbered_row(
    part: &LosslessXmlPart,
    row: &IndexedXmlElement,
    old_number: u32,
    new_number: u32,
) -> UseResult<Vec<u8>> {
    let fragment = element_fragment(part, row)?.to_vec();
    if old_number == new_number {
        return Ok(fragment);
    }
    let mut patches = Vec::<(std::ops::Range<usize>, Vec<u8>)>::new();
    patches.push((
        relative_range(row, &row.start_tag_range)?,
        updated_start_tag(row, &BTreeMap::from([("r".into(), new_number.to_string())]))
            .into_bytes(),
    ));
    for (reference, cell) in indexed_cells_in_row(old_number, row) {
        let updated = CellReference {
            column: reference.column,
            row: new_number,
        };
        patches.push((
            relative_range(row, &cell.start_tag_range)?,
            updated_start_tag(cell, &BTreeMap::from([("r".into(), updated.a1())])).into_bytes(),
        ));
    }
    apply_fragment_patches(fragment, patches)
}

fn relative_range(
    owner: &IndexedXmlElement,
    range: &std::ops::Range<usize>,
) -> UseResult<std::ops::Range<usize>> {
    let start = range
        .start
        .checked_sub(owner.full_range.start)
        .ok_or_else(row_arrange_unsupported)?;
    let end = range
        .end
        .checked_sub(owner.full_range.start)
        .ok_or_else(row_arrange_unsupported)?;
    Ok(start..end)
}

fn apply_fragment_patches(
    fragment: Vec<u8>,
    mut patches: Vec<(std::ops::Range<usize>, Vec<u8>)>,
) -> UseResult<Vec<u8>> {
    patches.sort_by_key(|(range, _)| (range.start, range.end));
    let mut output = Vec::new();
    let mut cursor = 0_usize;
    for (range, replacement) in patches {
        if range.start < cursor || range.end < range.start || range.end > fragment.len() {
            return Err(row_arrange_unsupported());
        }
        output.extend_from_slice(&fragment[cursor..range.start]);
        output.extend_from_slice(&replacement);
        cursor = range.end;
    }
    output.extend_from_slice(&fragment[cursor..]);
    Ok(output)
}

fn updated_start_tag(element: &IndexedXmlElement, updates: &BTreeMap<String, String>) -> String {
    let mut attributes = element.qualified_attributes.clone();
    attributes.extend(updates.clone());
    let attributes = attributes
        .into_iter()
        .map(|(name, value)| format!(" {name}=\"{}\"", quick_xml::escape::escape(&value)))
        .collect::<String>();
    let terminator = if element.empty { "/>" } else { ">" };
    format!("<{}{attributes}{terminator}", element.qualified_name)
}

fn row_arrange_unsupported() -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_row_arrange_unsupported",
        "Native Spreadsheet row move/copy/swap currently requires dense plain rows without formulas, defined names, row-addressed metadata, or worksheet relationships; the mutation was not applied.",
    )
}
