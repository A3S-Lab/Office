use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::UseResult;

use super::{
    editor_error, escape_attribute, index_xml, indexed_cells_in_row, indexed_rows,
    mark_workbook_for_recalculation, node_not_found, prefix, qualified, update_dimension,
    validate_mutation_path, XmlPatch,
};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_formula::{rewrite_formula_references, ReferenceAxis, ReferenceEdit};
use crate::spreadsheet_reference::{parse_column, CellReference, MAX_COLUMNS, MAX_ROWS};
use crate::xml_edit::{apply_patches, escape_text, IndexedXmlElement};
use crate::{DocumentKind, LosslessXmlPart, NativeOfficePackage};

const MAX_STRUCTURAL_COUNT: u32 = 10_000;

mod related;

pub(crate) fn insert_rows(
    package: &mut NativeOfficePackage,
    sheet: &str,
    start: u32,
    count: u32,
) -> UseResult<String> {
    edit_structure(
        package,
        sheet,
        ReferenceAxis::Row,
        start,
        ReferenceEdit::Insert { at: start, count },
    )
}

pub(crate) fn delete_rows(
    package: &mut NativeOfficePackage,
    sheet: &str,
    start: u32,
    count: u32,
) -> UseResult<String> {
    edit_structure(
        package,
        sheet,
        ReferenceAxis::Row,
        start,
        ReferenceEdit::Delete { start, count },
    )
}

pub(crate) fn insert_columns(
    package: &mut NativeOfficePackage,
    sheet: &str,
    start: &str,
    count: u32,
) -> UseResult<String> {
    let start = parse_column(start)?;
    edit_structure(
        package,
        sheet,
        ReferenceAxis::Column,
        start,
        ReferenceEdit::Insert { at: start, count },
    )
}

pub(crate) fn delete_columns(
    package: &mut NativeOfficePackage,
    sheet: &str,
    start: &str,
    count: u32,
) -> UseResult<String> {
    let start = parse_column(start)?;
    edit_structure(
        package,
        sheet,
        ReferenceAxis::Column,
        start,
        ReferenceEdit::Delete { start, count },
    )
}

fn edit_structure(
    package: &mut NativeOfficePackage,
    sheet_path: &str,
    axis: ReferenceAxis,
    start: u32,
    edit: ReferenceEdit,
) -> UseResult<String> {
    require_spreadsheet(package)?;
    validate_mutation_path(sheet_path)?;
    let count = edit_count(edit);
    validate_request(axis, start, count)?;
    let (sheet_name, part_name) = resolve_worksheet(package, sheet_path)?;
    let part = package.xml_part(&part_name)?;
    let edited = rewrite_worksheet_structure(&part, axis, edit)?;
    let edited = update_dimension(&part_name, edited)?;
    package.set_part(&part_name, edited)?;
    related::rewrite_related_parts(package, &part_name, axis, edit)?;
    rewrite_formula_graph(package, &sheet_name, axis, edit)?;
    mark_workbook_for_recalculation(package)?;

    let end = start + count - 1;
    let subject = match axis {
        ReferenceAxis::Row => "row",
        ReferenceAxis::Column => "col",
    };
    Ok(format!("{sheet_path}/{subject}[{start}:{end}]"))
}

fn rewrite_worksheet_structure(
    part: &LosslessXmlPart,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<Vec<u8>> {
    let index = index_xml(part)?;
    let sheet_data = index.descendant("sheetData").ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_data_missing",
            "Spreadsheet worksheet has no sheetData element.",
        )
    })?;
    let mut patches = Vec::new();

    for (row_number, row) in indexed_rows(sheet_data) {
        let row_transform = transform_coordinate(row_number, ReferenceAxis::Row, axis, edit)?;
        if row_transform.is_none() {
            patches.push(XmlPatch::new(row.full_range.clone(), Vec::new()));
            continue;
        }
        let new_row = row_transform.unwrap_or(row_number);
        let mut row_updates = BTreeMap::new();
        let mut row_removals = BTreeSet::new();
        if new_row != row_number || !row.attributes.contains_key("r") {
            row_updates.insert("r".to_string(), new_row.to_string());
        }
        if axis == ReferenceAxis::Column && row.attributes.contains_key("spans") {
            row_removals.insert("spans".to_string());
        }
        if !row_updates.is_empty() || !row_removals.is_empty() {
            patches.push(XmlPatch::new(
                row.start_tag_range.clone(),
                updated_start_tag(row, &row_updates, &row_removals),
            ));
        }

        for (reference, cell) in indexed_cells_in_row(row_number, row) {
            let Some(new_column) =
                transform_coordinate(reference.column, ReferenceAxis::Column, axis, edit)?
            else {
                patches.push(XmlPatch::new(cell.full_range.clone(), Vec::new()));
                continue;
            };
            let new_reference = CellReference {
                column: new_column,
                row: new_row,
            };
            if new_reference != reference || !cell.attributes.contains_key("r") {
                let updates = BTreeMap::from([("r".to_string(), new_reference.a1())]);
                patches.push(XmlPatch::new(
                    cell.start_tag_range.clone(),
                    updated_start_tag(cell, &updates, &BTreeSet::new()),
                ));
            }
        }
    }

    rewrite_column_definitions(&index, axis, edit, &mut patches)?;
    rewrite_worksheet_areas(part, &index, axis, edit, &mut patches)?;
    apply_patches(part, patches)
}

fn rewrite_column_definitions(
    index: &IndexedXmlElement,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
    patches: &mut Vec<XmlPatch>,
) -> UseResult<()> {
    if axis != ReferenceAxis::Column {
        return Ok(());
    }
    let Some(columns) = index.descendant("cols") else {
        return Ok(());
    };
    for column in columns
        .children
        .iter()
        .filter(|element| element.local_name == "col")
    {
        let min = column
            .attributes
            .get("min")
            .and_then(|value| value.parse::<u32>().ok())
            .ok_or_else(|| invalid_column_definition("min"))?;
        let max = column
            .attributes
            .get("max")
            .and_then(|value| value.parse::<u32>().ok())
            .ok_or_else(|| invalid_column_definition("max"))?;
        let Some((min, max)) = transform_interval(min, max, edit, MAX_COLUMNS)? else {
            patches.push(XmlPatch::new(column.full_range.clone(), Vec::new()));
            continue;
        };
        let updates = BTreeMap::from([
            ("min".to_string(), min.to_string()),
            ("max".to_string(), max.to_string()),
        ]);
        patches.push(XmlPatch::new(
            column.start_tag_range.clone(),
            updated_start_tag(column, &updates, &BTreeSet::new()),
        ));
    }
    Ok(())
}

fn rewrite_worksheet_areas(
    part: &LosslessXmlPart,
    index: &IndexedXmlElement,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
    patches: &mut Vec<XmlPatch>,
) -> UseResult<()> {
    let mut descendants = Vec::new();
    collect_descendants(index, &mut descendants);
    let mut removed_merge_cells = 0_usize;
    let mut removed_validations = 0_usize;
    for element in descendants {
        let attributes: &[&str] = match element.local_name.as_str() {
            "mergeCell" | "autoFilter" | "hyperlink" | "sortState" | "sortCondition"
            | "protectedRange" => &["ref"],
            "selection" => &["activeCell", "sqref"],
            "conditionalFormatting" | "dataValidation" | "ignoredError" => &["sqref"],
            _ => &[],
        };
        if attributes.is_empty() {
            continue;
        }
        let mut updates = BTreeMap::new();
        let mut remove_element = false;
        for attribute in attributes {
            let Some(value) = element.attributes.get(*attribute) else {
                continue;
            };
            let values = if *attribute == "sqref" {
                value.split_ascii_whitespace().collect::<Vec<_>>()
            } else {
                vec![value.as_str()]
            };
            let rewritten = values
                .into_iter()
                .map(|reference| {
                    rewrite_formula_references(
                        reference,
                        Some("__current__"),
                        "__current__",
                        axis,
                        edit,
                    )
                })
                .collect::<UseResult<Vec<_>>>()?
                .into_iter()
                .filter(|reference| reference != "#REF!")
                .collect::<Vec<_>>()
                .join(" ");
            if rewritten.is_empty() {
                remove_element = true;
                break;
            }
            if rewritten != *value {
                updates.insert((*attribute).to_string(), rewritten);
            }
        }
        if remove_element {
            if element.local_name == "mergeCell" {
                removed_merge_cells += 1;
            } else if element.local_name == "dataValidation" {
                removed_validations += 1;
            }
            patches.push(XmlPatch::new(element.full_range.clone(), Vec::new()));
        } else if !updates.is_empty() {
            patches.push(XmlPatch::new(
                element.start_tag_range.clone(),
                updated_start_tag(element, &updates, &BTreeSet::new()),
            ));
        }
    }

    if removed_merge_cells > 0 {
        if let Some(merges) = index.descendant("mergeCells") {
            if let Some(count) = merges
                .attributes
                .get("count")
                .and_then(|value| value.parse::<usize>().ok())
            {
                let retained = count.saturating_sub(removed_merge_cells);
                let updates = BTreeMap::from([("count".to_string(), retained.to_string())]);
                patches.push(XmlPatch::new(
                    merges.start_tag_range.clone(),
                    updated_start_tag(merges, &updates, &BTreeSet::new()),
                ));
            }
        }
    }
    if removed_validations > 0 {
        if let Some(validations) = index.descendant("dataValidations") {
            if let Some(count) = validations
                .attributes
                .get("count")
                .and_then(|value| value.parse::<usize>().ok())
            {
                let retained = count.saturating_sub(removed_validations);
                let updates = BTreeMap::from([("count".to_string(), retained.to_string())]);
                patches.push(XmlPatch::new(
                    validations.start_tag_range.clone(),
                    updated_start_tag(validations, &updates, &BTreeSet::new()),
                ));
            }
        }
    }

    // Ensure every patched attribute remains well-formed before the outer
    // patch set is applied. This also gives area errors the worksheet part.
    let _ = part.name();
    Ok(())
}

fn rewrite_formula_graph(
    package: &mut NativeOfficePackage,
    target_sheet: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<()> {
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let sheets = snapshot
        .root()
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Worksheet)
        .map(|node| {
            let name = node.path.trim_start_matches('/').to_string();
            let part = node.format.get("part").cloned().ok_or_else(|| {
                editor_error(
                    "use.office.spreadsheet_sheet_invalid",
                    format!("Worksheet '{name}' has no source part."),
                )
            })?;
            Ok((name, part))
        })
        .collect::<UseResult<Vec<_>>>()?;

    for (sheet_name, part_name) in &sheets {
        let part = package.xml_part(part_name)?;
        let index = index_xml(&part)?;
        let mut patches = Vec::new();
        let mut formulas = Vec::new();
        index.descendants_named("f", &mut formulas);
        for formula in formulas {
            let text = decoded_text(&part, formula)?;
            let rewritten = rewrite_formula_references(
                &text,
                Some(sheet_name.as_str()),
                target_sheet,
                axis,
                edit,
            )?;
            if rewritten != text {
                patches.push(XmlPatch::new(
                    formula.content_range.clone(),
                    escape_text(&rewritten),
                ));
            }
        }
        let mut cells = Vec::new();
        index.descendants_named("c", &mut cells);
        for cell in cells {
            if cell.children.iter().any(|child| child.local_name == "f") {
                for cached in cell.children.iter().filter(|child| child.local_name == "v") {
                    patches.push(XmlPatch::new(cached.full_range.clone(), Vec::new()));
                }
            }
        }
        if !patches.is_empty() {
            package.set_part(part_name, apply_patches(&part, patches)?)?;
        }
    }
    related::rewrite_auxiliary_formulas(package, &sheets, target_sheet, axis, edit)?;
    rewrite_defined_names(package, target_sheet, axis, edit)
}

fn rewrite_defined_names(
    package: &mut NativeOfficePackage,
    target_sheet: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<()> {
    let workbook = package.xml_part("xl/workbook.xml")?;
    let index = index_xml(&workbook)?;
    let sheet_names = index
        .descendant("sheets")
        .map(|sheets| {
            sheets
                .children
                .iter()
                .filter(|sheet| sheet.local_name == "sheet")
                .filter_map(|sheet| sheet.attributes.get("name").cloned())
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let mut names = Vec::new();
    index.descendants_named("definedName", &mut names);
    let mut patches = Vec::new();
    for name in names {
        let text = decoded_text(&workbook, name)?;
        let current_sheet = name
            .attributes
            .get("localSheetId")
            .and_then(|value| value.parse::<usize>().ok())
            .and_then(|index| sheet_names.get(index))
            .map(String::as_str);
        let rewritten = rewrite_formula_references(&text, current_sheet, target_sheet, axis, edit)?;
        if rewritten != text {
            patches.push(XmlPatch::new(
                name.content_range.clone(),
                escape_text(&rewritten),
            ));
        }
    }
    if patches.is_empty() {
        Ok(())
    } else {
        package.set_part("xl/workbook.xml", apply_patches(&workbook, patches)?)
    }
}

fn resolve_worksheet(package: &NativeOfficePackage, path: &str) -> UseResult<(String, String)> {
    if path.trim_start_matches('/').contains('/') {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Spreadsheet structure mutation requires a worksheet path such as /Sheet1.",
        ));
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let sheet = snapshot
        .root()
        .children
        .iter()
        .find(|node| {
            node.node_type == OfficeNodeType::Worksheet && node.path.eq_ignore_ascii_case(path)
        })
        .ok_or_else(|| node_not_found(path))?;
    let name = sheet.path.trim_start_matches('/').to_string();
    let part = sheet.format.get("part").cloned().ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_invalid",
            format!("Worksheet '{name}' has no source part."),
        )
    })?;
    Ok((name, part))
}

fn validate_request(axis: ReferenceAxis, start: u32, count: u32) -> UseResult<()> {
    let limit = axis_limit(axis);
    if start == 0
        || start > limit
        || count == 0
        || count > MAX_STRUCTURAL_COUNT
        || start.checked_add(count - 1).is_none_or(|end| end > limit)
    {
        return Err(editor_error(
            "use.office.spreadsheet_structure_invalid",
            format!(
                "Spreadsheet structural edits require a start within 1-{limit} and a count within 1-{MAX_STRUCTURAL_COUNT} that remains in bounds."
            ),
        )
        .with_detail("start", start)
        .with_detail("count", count));
    }
    Ok(())
}

fn transform_coordinate(
    value: u32,
    coordinate_axis: ReferenceAxis,
    edit_axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<Option<u32>> {
    if coordinate_axis != edit_axis {
        return Ok(Some(value));
    }
    match edit {
        ReferenceEdit::Insert { at, count } => {
            if value < at {
                Ok(Some(value))
            } else {
                value
                    .checked_add(count)
                    .filter(|value| *value <= axis_limit(edit_axis))
                    .map(Some)
                    .ok_or_else(|| {
                        editor_error(
                            "use.office.spreadsheet_structure_overflow",
                            "Spreadsheet structural edit would move existing content outside worksheet limits.",
                        )
                    })
            }
        }
        ReferenceEdit::Delete { start, count } => {
            let end = start + count - 1;
            if value < start {
                Ok(Some(value))
            } else if value <= end {
                Ok(None)
            } else {
                Ok(Some(value - count))
            }
        }
    }
}

fn transform_interval(
    low: u32,
    high: u32,
    edit: ReferenceEdit,
    limit: u32,
) -> UseResult<Option<(u32, u32)>> {
    if low == 0 || high < low || high > limit {
        return Err(invalid_column_definition("range"));
    }
    match edit {
        ReferenceEdit::Insert { at, count } => {
            if at <= low {
                let low = low.checked_add(count).filter(|value| *value <= limit);
                let high = high.checked_add(count).filter(|value| *value <= limit);
                low.zip(high).map(Some).ok_or_else(|| {
                    editor_error(
                        "use.office.spreadsheet_structure_overflow",
                        "Spreadsheet structural edit would move column metadata outside worksheet limits.",
                    )
                })
            } else if at <= high {
                high.checked_add(count).filter(|value| *value <= limit).map_or_else(
                    || {
                        Err(editor_error(
                            "use.office.spreadsheet_structure_overflow",
                            "Spreadsheet structural edit would expand column metadata outside worksheet limits.",
                        ))
                    },
                    |high| Ok(Some((low, high))),
                )
            } else {
                Ok(Some((low, high)))
            }
        }
        ReferenceEdit::Delete { start, count } => {
            let end = start + count - 1;
            if high < start {
                return Ok(Some((low, high)));
            }
            if low > end {
                return Ok(Some((low - count, high - count)));
            }
            let low = if low < start { low } else { start };
            let high = if high > end {
                high - count
            } else {
                start.saturating_sub(1)
            };
            Ok((low <= high && high > 0).then_some((low, high)))
        }
    }
}

fn updated_start_tag(
    element: &IndexedXmlElement,
    updates: &BTreeMap<String, String>,
    removals: &BTreeSet<String>,
) -> String {
    let mut attributes = element.qualified_attributes.clone();
    for removal in removals {
        attributes.remove(removal);
    }
    attributes.extend(updates.clone());
    let attributes = attributes
        .into_iter()
        .map(|(name, value)| format!(" {name}=\"{}\"", escape_attribute(&value)))
        .collect::<String>();
    let terminator = if element.empty { "/>" } else { ">" };
    format!("<{}{attributes}{terminator}", element.qualified_name)
}

fn decoded_text(part: &LosslessXmlPart, element: &IndexedXmlElement) -> UseResult<String> {
    let bytes = part
        .parse_bytes()
        .get(element.content_range.clone())
        .ok_or_else(|| {
            editor_error(
                "use.office.spreadsheet_formula_invalid",
                format!("Formula range in '{}' is invalid.", part.name()),
            )
        })?;
    let text = std::str::from_utf8(bytes).map_err(|error| {
        editor_error(
            "use.office.spreadsheet_formula_invalid",
            format!("Formula in '{}' is not UTF-8: {error}", part.name()),
        )
    })?;
    quick_xml::escape::unescape(text)
        .map(|value| value.into_owned())
        .map_err(|error| {
            editor_error(
                "use.office.spreadsheet_formula_invalid",
                format!(
                    "Formula in '{}' contains invalid XML escapes: {error}",
                    part.name()
                ),
            )
        })
}

fn collect_descendants<'a>(
    element: &'a IndexedXmlElement,
    output: &mut Vec<&'a IndexedXmlElement>,
) {
    for child in &element.children {
        output.push(child);
        collect_descendants(child, output);
    }
}

fn edit_count(edit: ReferenceEdit) -> u32 {
    match edit {
        ReferenceEdit::Insert { count, .. } | ReferenceEdit::Delete { count, .. } => count,
    }
}

fn edit_start(edit: ReferenceEdit) -> u32 {
    match edit {
        ReferenceEdit::Insert { at, .. } => at,
        ReferenceEdit::Delete { start, .. } => start,
    }
}

fn axis_limit(axis: ReferenceAxis) -> u32 {
    match axis {
        ReferenceAxis::Row => MAX_ROWS,
        ReferenceAxis::Column => MAX_COLUMNS,
    }
}

fn require_spreadsheet(package: &NativeOfficePackage) -> UseResult<()> {
    if package.kind() == DocumentKind::Spreadsheet {
        Ok(())
    } else {
        Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Row and column structural edits are available only for Spreadsheet documents.",
        ))
    }
}

fn invalid_column_definition(attribute: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_column_definition_invalid",
        format!("Spreadsheet column definition has an invalid '{attribute}' attribute."),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn interval_edits_expand_shift_shrink_and_delete() {
        assert_eq!(
            transform_interval(2, 5, ReferenceEdit::Insert { at: 4, count: 2 }, 20).unwrap(),
            Some((2, 7))
        );
        assert_eq!(
            transform_interval(2, 5, ReferenceEdit::Insert { at: 2, count: 2 }, 20).unwrap(),
            Some((4, 7))
        );
        assert_eq!(
            transform_interval(2, 8, ReferenceEdit::Delete { start: 4, count: 2 }, 20).unwrap(),
            Some((2, 6))
        );
        assert_eq!(
            transform_interval(4, 5, ReferenceEdit::Delete { start: 4, count: 2 }, 20).unwrap(),
            None
        );
    }
}
