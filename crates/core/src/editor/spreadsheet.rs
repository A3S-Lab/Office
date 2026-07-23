use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{
    editor_error, escape_attribute, node_not_found, prefix, preserve_space_attribute, qualified,
    validate_mutation_path, SpreadsheetCellValue,
};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::{column_name, CellRange, CellReference};
use crate::xml_edit::{index_xml, insert_child, IndexedXmlElement, XmlPatch};
use crate::{DocumentKind, NativeOfficePackage};

const MAX_RANGE_MUTATION_CELLS: usize = 100_000;

mod arrange;
mod auto_filter;
mod conditional_formatting;
mod data_validation;
mod filter_xml;
mod formula;
mod import;
mod merge;
mod named_range;
mod sort;
mod structure;
mod style;
mod table;
mod view;
mod worksheet;

pub(super) use arrange::{copy_node, move_node, swap_nodes};
pub(super) use structure::{delete_columns, delete_rows, insert_columns, insert_rows};
pub(super) use worksheet::{copy_worksheet, move_worksheet, rename_worksheet};

pub(super) fn recalculate_formulas(
    package: &mut NativeOfficePackage,
) -> UseResult<crate::SpreadsheetFormulaCalculation> {
    formula::recalculate(package)
}

pub(super) fn add_auto_filter(
    package: &mut NativeOfficePackage,
    sheet: &str,
    filter: &super::NativeSpreadsheetAutoFilter,
) -> UseResult<String> {
    auto_filter::add(package, sheet, filter)
}

pub(super) fn set_auto_filter(
    package: &mut NativeOfficePackage,
    path: &str,
    filter: &super::NativeSpreadsheetAutoFilter,
) -> UseResult<String> {
    auto_filter::set(package, path, filter)
}

pub(super) fn sort_range(
    package: &mut NativeOfficePackage,
    path: &str,
    value: &super::NativeSpreadsheetSort,
) -> UseResult<String> {
    sort::sort(package, path, value)
}

pub(super) fn set_frozen_pane(
    package: &mut NativeOfficePackage,
    sheet: &str,
    pane: &super::NativeSpreadsheetFrozenPane,
) -> UseResult<String> {
    view::set(package, sheet, pane)
}

pub(super) fn import_delimited(
    package: &mut NativeOfficePackage,
    sheet: &str,
    import: &super::NativeSpreadsheetDelimitedImport,
) -> UseResult<super::NativeSpreadsheetImportResult> {
    import::apply(package, sheet, import)
}

pub(super) fn add_conditional_format(
    package: &mut NativeOfficePackage,
    sheet: &str,
    value: &super::NativeSpreadsheetConditionalFormat,
) -> UseResult<String> {
    conditional_formatting::add(package, sheet, value)
}

pub(super) fn set_conditional_format(
    package: &mut NativeOfficePackage,
    path: &str,
    value: &super::NativeSpreadsheetConditionalFormat,
) -> UseResult<String> {
    conditional_formatting::set(package, path, value)
}

pub(super) fn add_data_validation(
    package: &mut NativeOfficePackage,
    sheet: &str,
    validation: &super::NativeSpreadsheetDataValidation,
) -> UseResult<String> {
    data_validation::add(package, sheet, validation)
}

pub(super) fn set_data_validation(
    package: &mut NativeOfficePackage,
    path: &str,
    validation: &super::NativeSpreadsheetDataValidation,
) -> UseResult<String> {
    data_validation::set(package, path, validation)
}

pub(super) fn add_named_range(
    package: &mut NativeOfficePackage,
    named_range: &super::NativeSpreadsheetNamedRange,
) -> UseResult<String> {
    named_range::add(package, named_range)
}

pub(super) fn add_table(
    package: &mut NativeOfficePackage,
    sheet: &str,
    table: &super::NativeSpreadsheetTable,
) -> UseResult<String> {
    table::add(package, sheet, table)
}

pub(super) fn set_table(
    package: &mut NativeOfficePackage,
    path: &str,
    table: &super::NativeSpreadsheetTable,
) -> UseResult<String> {
    table::set(package, path, table)
}

pub(super) fn set_named_range(
    package: &mut NativeOfficePackage,
    path: &str,
    named_range: &super::NativeSpreadsheetNamedRange,
) -> UseResult<String> {
    named_range::set(package, path, named_range)
}

pub(super) fn merge_cells(package: &mut NativeOfficePackage, path: &str) -> UseResult<String> {
    merge::merge_cells(package, path)
}

pub(super) fn unmerge_cells(package: &mut NativeOfficePackage, path: &str) -> UseResult<String> {
    merge::unmerge_cells(package, path)
}

pub(super) fn set_text_format(
    package: &mut NativeOfficePackage,
    path: &str,
    format: &super::NativeOfficeTextFormat,
) -> UseResult<()> {
    style::set_text_format(package, path, format)
}

pub(super) fn set_cell_format(
    package: &mut NativeOfficePackage,
    path: &str,
    format: &super::NativeSpreadsheetCellFormat,
) -> UseResult<()> {
    style::set_cell_format(package, path, format)
}

pub(super) fn set_text(package: &mut NativeOfficePackage, path: &str, text: &str) -> UseResult<()> {
    set_cell_value(
        package,
        path,
        &SpreadsheetCellValue::Text {
            value: text.to_string(),
        },
    )
}

pub(super) fn set_cell_value(
    package: &mut NativeOfficePackage,
    path: &str,
    value: &SpreadsheetCellValue,
) -> UseResult<()> {
    if package.kind() != DocumentKind::Spreadsheet {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Typed cell values are available only for Spreadsheet documents.",
        ));
    }
    validate_mutation_path(path)?;
    let value = normalize_cell_value(value)?;
    let (sheet_path, reference) = path.rsplit_once('/').ok_or_else(|| node_not_found(path))?;
    let range = CellRange::parse(reference)?;
    validate_range_size(range)?;
    if sheet_path.is_empty() {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Spreadsheet cell mutation requires a single-cell path such as /Sheet1/A1.",
        ));
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let sheet = snapshot
        .root()
        .children
        .iter()
        .find(|node| {
            node.node_type == OfficeNodeType::Worksheet
                && node.path.eq_ignore_ascii_case(sheet_path)
        })
        .ok_or_else(|| node_not_found(path))?;
    let part_name = sheet.format.get("part").ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_invalid",
            "Spreadsheet semantic sheet has no source part.",
        )
    })?;
    let part = package.xml_part(part_name)?;
    let index = index_xml(&part)?;
    let sheet_data = index
        .descendant("sheetData")
        .ok_or_else(|| node_not_found(path))?;
    let prepared = formula::prepare_for_value_write(&part, sheet_data, sheet, range)?;
    let part = crate::LosslessXmlPart::parse(part_name.to_string(), prepared)?;
    let index = index_xml(&part)?;
    let sheet_data = index
        .descendant("sheetData")
        .ok_or_else(|| node_not_found(path))?;
    let edited = set_range(&part, sheet_data, range, &value)?;
    let edited = update_dimension(part_name, edited)?;
    package.set_part(part_name, edited)?;
    mark_workbook_for_recalculation(package)
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    if view::is_path(path) {
        view::remove(package, path)
    } else if sort::is_path(path) {
        sort::remove(package, path)
    } else if auto_filter::is_path(path) {
        auto_filter::remove(package, path)
    } else if table::is_path(path) {
        table::remove(package, path)
    } else if named_range::is_path(path) {
        named_range::remove(package, path)
    } else if conditional_formatting::is_path(path) {
        conditional_formatting::remove(package, path)
    } else if data_validation::is_path(path) {
        data_validation::remove(package, path)
    } else if path.trim_start_matches('/').contains('/') {
        remove_cell(package, path)
    } else {
        remove_worksheet(package, path)
    }
}

fn remove_cell(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let (sheet_path, reference) = path.rsplit_once('/').ok_or_else(|| node_not_found(path))?;
    let range = CellRange::parse(reference)?;
    validate_range_size(range)?;
    if sheet_path.is_empty() {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Spreadsheet remove requires a single-cell path such as /Sheet1/A1.",
        ));
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let sheet = snapshot
        .root()
        .children
        .iter()
        .find(|node| {
            node.node_type == OfficeNodeType::Worksheet
                && node.path.eq_ignore_ascii_case(sheet_path)
        })
        .ok_or_else(|| node_not_found(path))?;
    let part_name = sheet.format.get("part").ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_invalid",
            "Spreadsheet semantic sheet has no source part.",
        )
    })?;
    let part = package.xml_part(part_name)?;
    let index = index_xml(&part)?;
    let sheet_data = index
        .descendant("sheetData")
        .ok_or_else(|| node_not_found(path))?;
    let prepared = formula::prepare_for_remove(&part, sheet_data, sheet, range)?;
    package.set_part(part_name, prepared)?;
    super::comment::remove_spreadsheet_range_comments(package, path)?;
    super::hyperlink::remove_spreadsheet_range_links(package, path)?;
    let part = package.xml_part(part_name)?;
    let index = index_xml(&part)?;
    let sheet_data = index
        .descendant("sheetData")
        .ok_or_else(|| node_not_found(path))?;
    let cells = indexed_cells(sheet_data)
        .into_iter()
        .filter(|(reference, _, _)| range.contains(*reference))
        .collect::<Vec<_>>();
    if cells.is_empty() {
        return Err(node_not_found(path));
    }
    let edited = crate::xml_edit::apply_patches(
        &part,
        cells
            .into_iter()
            .map(|(_, _, cell)| XmlPatch::new(cell.full_range.clone(), Vec::new()))
            .collect(),
    )?;
    let edited = update_dimension(part_name, edited)?;
    package.set_part(part_name, edited)?;
    mark_workbook_for_recalculation(package)
}

fn set_range(
    part: &crate::LosslessXmlPart,
    sheet_data: &IndexedXmlElement,
    range: CellRange,
    value: &SpreadsheetCellValue,
) -> UseResult<Vec<u8>> {
    if sheet_data.empty {
        let cell_prefix = prefix(&sheet_data.qualified_name);
        let row_tag = qualified(cell_prefix, "row");
        let rows = (range.start.row..=range.end.row)
            .map(|row_number| {
                let cells = (range.start.column..=range.end.column)
                    .map(|column| {
                        let reference = CellReference {
                            column,
                            row: row_number,
                        }
                        .a1();
                        new_cell_fragment(cell_prefix, &reference, value)
                    })
                    .collect::<String>();
                format!("<{row_tag} r=\"{row_number}\">{cells}</{row_tag}>")
            })
            .collect::<String>();
        return insert_child(part, sheet_data, rows);
    }

    let rows = indexed_rows(sheet_data);
    let row_map = rows.iter().copied().collect::<BTreeMap<_, _>>();
    let mut patches = Vec::new();
    let mut insertions = BTreeMap::<usize, Vec<(u32, u32, String)>>::new();
    let cell_prefix = prefix(&sheet_data.qualified_name);

    for row_number in range.start.row..=range.end.row {
        let Some(row) = row_map.get(&row_number).copied() else {
            let cells = (range.start.column..=range.end.column)
                .map(|column| {
                    let reference = CellReference {
                        column,
                        row: row_number,
                    }
                    .a1();
                    new_cell_fragment(cell_prefix, &reference, value)
                })
                .collect::<String>();
            let row_tag = qualified(cell_prefix, "row");
            let fragment = format!("<{row_tag} r=\"{row_number}\">{cells}</{row_tag}>");
            let position = rows
                .iter()
                .find(|(existing, _)| *existing > row_number)
                .map_or(sheet_data.content_range.end, |(_, next)| {
                    next.full_range.start
                });
            insertions
                .entry(position)
                .or_default()
                .push((row_number, 0, fragment));
            continue;
        };

        if row.empty {
            let cells = (range.start.column..=range.end.column)
                .map(|column| {
                    let reference = CellReference {
                        column,
                        row: row_number,
                    }
                    .a1();
                    new_cell_fragment(prefix(&row.qualified_name), &reference, value)
                })
                .collect::<String>();
            patches.push(XmlPatch::new(
                row.full_range.clone(),
                expanded_element(row, &cells),
            ));
            continue;
        }

        let cells = indexed_cells_in_row(row_number, row);
        for column in range.start.column..=range.end.column {
            let reference = CellReference {
                column,
                row: row_number,
            };
            if let Some((_, cell)) = cells.iter().find(|(existing, _)| existing.column == column) {
                patches.push(XmlPatch::new(
                    cell.full_range.clone(),
                    cell_fragment(cell, &reference.a1(), value),
                ));
                continue;
            }
            let position = cells
                .iter()
                .find(|(existing, _)| existing.column > column)
                .map(|(_, next)| next.full_range.start)
                .or_else(|| {
                    row.children
                        .iter()
                        .find(|child| child.local_name != "c")
                        .map(|child| child.full_range.start)
                })
                .unwrap_or(row.content_range.end);
            insertions.entry(position).or_default().push((
                row_number,
                column,
                new_cell_fragment(prefix(&row.qualified_name), &reference.a1(), value),
            ));
        }
    }

    for (position, mut fragments) in insertions {
        fragments.sort_by_key(|(row, column, _)| (*row, *column));
        patches.push(XmlPatch::new(
            position..position,
            fragments
                .into_iter()
                .map(|(_, _, fragment)| fragment)
                .collect::<String>(),
        ));
    }
    crate::xml_edit::apply_patches(part, patches)
}

fn indexed_rows(sheet_data: &IndexedXmlElement) -> Vec<(u32, &IndexedXmlElement)> {
    let mut inferred = 0_u32;
    sheet_data
        .children
        .iter()
        .filter(|child| child.local_name == "row")
        .map(|row| {
            let number = row
                .attributes
                .get("r")
                .and_then(|value| value.parse::<u32>().ok())
                .unwrap_or_else(|| inferred.saturating_add(1));
            inferred = number;
            (number, row)
        })
        .collect()
}

fn indexed_cells(
    sheet_data: &IndexedXmlElement,
) -> Vec<(CellReference, &IndexedXmlElement, &IndexedXmlElement)> {
    indexed_rows(sheet_data)
        .into_iter()
        .flat_map(|(row_number, row)| {
            indexed_cells_in_row(row_number, row)
                .into_iter()
                .map(move |(reference, cell)| (reference, row, cell))
        })
        .collect()
}

fn indexed_cells_in_row(
    row_number: u32,
    row: &IndexedXmlElement,
) -> Vec<(CellReference, &IndexedXmlElement)> {
    let mut inferred_column = 0_u32;
    row.children
        .iter()
        .filter(|child| child.local_name == "c")
        .filter_map(|cell| {
            let reference = cell
                .attributes
                .get("r")
                .and_then(|reference| CellReference::parse(reference).ok())
                .unwrap_or_else(|| {
                    inferred_column = inferred_column.saturating_add(1);
                    CellReference {
                        column: inferred_column,
                        row: row_number,
                    }
                });
            inferred_column = reference.column;
            (reference.row == row_number).then_some((reference, cell))
        })
        .collect()
}

fn expanded_element(element: &IndexedXmlElement, content: &str) -> String {
    let attributes = element
        .qualified_attributes
        .iter()
        .map(|(name, value)| format!(" {name}=\"{}\"", escape_attribute(value)))
        .collect::<String>();
    format!(
        "<{}{attributes}>{content}</{}>",
        element.qualified_name, element.qualified_name
    )
}

fn validate_range_size(range: CellRange) -> UseResult<()> {
    let cells = range.cell_count()?;
    if cells > MAX_RANGE_MUTATION_CELLS {
        return Err(editor_error(
            "use.office.spreadsheet_range_too_large",
            format!(
                "Native Spreadsheet range mutations support at most {MAX_RANGE_MUTATION_CELLS} cells; '{}' contains {cells}.",
                range.a1()
            ),
        )
        .with_detail("cells", cells));
    }
    Ok(())
}

fn new_cell_fragment(
    prefix: Option<&str>,
    reference: &str,
    value: &SpreadsheetCellValue,
) -> String {
    let cell_tag = qualified(prefix, "c");
    let (cell_type, content) = cell_content(prefix, value);
    let cell_type = cell_type.map_or_else(String::new, |value_type| {
        format!(" t=\"{}\"", escape_attribute(value_type))
    });
    format!("<{cell_tag} r=\"{reference}\"{cell_type}>{content}</{cell_tag}>")
}

fn update_dimension(part_name: &str, bytes: Vec<u8>) -> UseResult<Vec<u8>> {
    let part = crate::LosslessXmlPart::parse(part_name.to_string(), bytes)?;
    let index = index_xml(&part)?;
    let sheet_data = index
        .descendant("sheetData")
        .ok_or_else(|| node_not_found(part_name))?;
    let mut bounds: Option<(u32, u32, u32, u32)> = None;
    for (row_number, row) in indexed_rows(sheet_data) {
        let mut inferred_column = 0_u32;
        for cell in row.children.iter().filter(|child| child.local_name == "c") {
            let (column, row_number, _) = cell
                .attributes
                .get("r")
                .and_then(|reference| cell_coordinates(reference).ok())
                .unwrap_or_else(|| {
                    inferred_column = inferred_column.saturating_add(1);
                    (inferred_column, row_number, String::new())
                });
            inferred_column = column;
            bounds = Some(match bounds {
                Some((min_column, min_row, max_column, max_row)) => (
                    min_column.min(column),
                    min_row.min(row_number),
                    max_column.max(column),
                    max_row.max(row_number),
                ),
                None => (column, row_number, column, row_number),
            });
        }
    }
    let dimension = bounds.map_or_else(
        || "A1".to_string(),
        |(min_column, min_row, max_column, max_row)| {
            let start = format!("{}{min_row}", column_name(min_column));
            let end = format!("{}{max_row}", column_name(max_column));
            if start == end {
                start
            } else {
                format!("{start}:{end}")
            }
        },
    );
    if let Some(existing) = index.child("dimension", 1) {
        let tag = &existing.qualified_name;
        return crate::xml_edit::apply_patches(
            &part,
            vec![XmlPatch::new(
                existing.full_range.clone(),
                format!("<{tag} ref=\"{dimension}\"/>"),
            )],
        );
    }
    let tag = qualified(prefix(&index.qualified_name), "dimension");
    let insertion = index
        .children
        .iter()
        .find(|child| child.local_name != "sheetPr")
        .map_or(index.content_range.end, |child| child.full_range.start);
    crate::xml_edit::apply_patches(
        &part,
        vec![XmlPatch::new(
            insertion..insertion,
            format!("<{tag} ref=\"{dimension}\"/>"),
        )],
    )
}

fn cell_fragment(
    cell: &IndexedXmlElement,
    reference: &str,
    value: &SpreadsheetCellValue,
) -> String {
    let prefix = prefix(&cell.qualified_name);
    let cell_tag = qualified(prefix, "c");
    let mut attributes = cell.qualified_attributes.clone();
    attributes.insert("r".into(), reference.to_ascii_uppercase());
    let (cell_type, content) = cell_content(prefix, value);
    if let Some(cell_type) = cell_type {
        attributes.insert("t".into(), cell_type.to_string());
    } else {
        attributes.remove("t");
    }
    let attributes = attributes
        .into_iter()
        .map(|(name, value)| format!(" {name}=\"{}\"", escape_attribute(&value)))
        .collect::<String>();
    format!("<{cell_tag}{attributes}>{content}</{cell_tag}>")
}

fn cell_content(
    prefix: Option<&str>,
    value: &SpreadsheetCellValue,
) -> (Option<&'static str>, String) {
    match value {
        SpreadsheetCellValue::Text { value } => {
            let inline_tag = qualified(prefix, "is");
            let text_tag = qualified(prefix, "t");
            let space = preserve_space_attribute(value);
            let value = crate::xml_edit::escape_text(value);
            (
                Some("inlineStr"),
                format!("<{inline_tag}><{text_tag}{space}>{value}</{text_tag}></{inline_tag}>"),
            )
        }
        SpreadsheetCellValue::Number { value } => {
            let value_tag = qualified(prefix, "v");
            (None, format!("<{value_tag}>{value}</{value_tag}>"))
        }
        SpreadsheetCellValue::Boolean { value } => {
            let value_tag = qualified(prefix, "v");
            let value = if *value { "1" } else { "0" };
            (Some("b"), format!("<{value_tag}>{value}</{value_tag}>"))
        }
        SpreadsheetCellValue::Formula { expression } => {
            let formula_tag = qualified(prefix, "f");
            let expression = crate::xml_edit::escape_text(expression);
            (None, format!("<{formula_tag}>{expression}</{formula_tag}>"))
        }
    }
}

fn normalize_cell_value(value: &SpreadsheetCellValue) -> UseResult<SpreadsheetCellValue> {
    match value {
        SpreadsheetCellValue::Text { value } => Ok(SpreadsheetCellValue::Text {
            value: value.clone(),
        }),
        SpreadsheetCellValue::Number { value } => {
            if value.is_empty()
                || value.len() > 128
                || value.trim() != value
                || !value.parse::<f64>().ok().is_some_and(f64::is_finite)
            {
                return Err(editor_error(
                    "use.office.spreadsheet_number_invalid",
                    "Spreadsheet numeric values must be bounded finite numbers without surrounding whitespace.",
                )
                .with_detail("length", value.len()));
            }
            Ok(SpreadsheetCellValue::Number {
                value: value.clone(),
            })
        }
        SpreadsheetCellValue::Boolean { value } => {
            Ok(SpreadsheetCellValue::Boolean { value: *value })
        }
        SpreadsheetCellValue::Formula { expression } => {
            let expression =
                crate::spreadsheet_formula::validate_and_normalize_formula(expression)?;
            Ok(SpreadsheetCellValue::Formula {
                expression: expression.to_string(),
            })
        }
    }
}

fn mark_workbook_for_recalculation(package: &mut NativeOfficePackage) -> UseResult<()> {
    remove_calculation_chain(package)?;
    let workbook = package.xml_part("xl/workbook.xml")?;
    let index = index_xml(&workbook)?;
    let edited = if let Some(calc) = index.child("calcPr", 1) {
        let mut attributes = calc.qualified_attributes.clone();
        attributes
            .entry("calcId".into())
            .or_insert_with(|| "0".into());
        attributes.insert("calcMode".into(), "auto".into());
        attributes.insert("fullCalcOnLoad".into(), "1".into());
        attributes.insert("forceFullCalc".into(), "1".into());
        let attributes = attributes
            .into_iter()
            .map(|(name, value)| format!(" {name}=\"{}\"", escape_attribute(&value)))
            .collect::<String>();
        let terminator = if calc.empty { "/>" } else { ">" };
        crate::xml_edit::apply_patches(
            &workbook,
            vec![XmlPatch::new(
                calc.start_tag_range.clone(),
                format!("<{}{attributes}{terminator}", calc.qualified_name),
            )],
        )?
    } else {
        let tag = qualified(prefix(&index.qualified_name), "calcPr");
        let fragment = format!(
            "<{tag} calcId=\"0\" calcMode=\"auto\" fullCalcOnLoad=\"1\" forceFullCalc=\"1\"/>"
        );
        let insertion = index
            .children
            .iter()
            .find(|child| {
                matches!(
                    child.local_name.as_str(),
                    "oleSize"
                        | "customWorkbookViews"
                        | "pivotCaches"
                        | "smartTagPr"
                        | "smartTagTypes"
                        | "webPublishing"
                        | "fileRecoveryPr"
                        | "webPublishObjects"
                        | "extLst"
                )
            })
            .map_or(index.content_range.end, |child| child.full_range.start);
        crate::xml_edit::apply_patches(
            &workbook,
            vec![XmlPatch::new(insertion..insertion, fragment)],
        )?
    };
    package.set_part("xl/workbook.xml", edited)
}

fn remove_calculation_chain(package: &mut NativeOfficePackage) -> UseResult<()> {
    let model = package.opc_model()?;
    let source = crate::RelationshipSource::Part {
        part_name: "xl/workbook.xml".into(),
    };
    let relationships = model
        .relationships()
        .relationships_from(&source)
        .iter()
        .filter(|relationship| relationship.relationship_type.ends_with("/calcChain"))
        .filter_map(|relationship| {
            relationship
                .target
                .internal_part_name()
                .map(|part| (relationship.id.clone(), part.to_string()))
        })
        .collect::<Vec<_>>();
    if relationships.is_empty() {
        return Ok(());
    }
    let targets = relationships
        .iter()
        .map(|(_, target)| target.clone())
        .collect::<std::collections::BTreeSet<_>>();
    for (id, _) in relationships {
        crate::opc_edit::remove_relationship(package, "xl/_rels/workbook.xml.rels", &id)?;
    }
    for target in targets {
        let relationships = worksheet::relationship_part(&target);
        if model
            .content_types()
            .override_for_part(&relationships)
            .is_some()
        {
            crate::opc_edit::remove_content_type_override(package, &relationships)?;
        }
        package.remove_part(&relationships)?;
        if model.content_types().override_for_part(&target).is_some() {
            crate::opc_edit::remove_content_type_override(package, &target)?;
        }
        package.remove_part(&target)?;
    }
    Ok(())
}

fn cell_coordinates(reference: &str) -> UseResult<(u32, u32, String)> {
    let reference = CellReference::parse(reference)?;
    Ok((reference.column, reference.row, reference.a1()))
}

pub(super) fn add_worksheet(package: &mut NativeOfficePackage, name: &str) -> UseResult<String> {
    if package.kind() != DocumentKind::Spreadsheet {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Native add-worksheet is available only for Spreadsheet documents.",
        ));
    }
    validate_worksheet_name(name)?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    if snapshot.root().children.iter().any(|sheet| {
        sheet.node_type == OfficeNodeType::Worksheet && sheet.path[1..].eq_ignore_ascii_case(name)
    }) {
        return Err(editor_error(
            "use.office.spreadsheet_sheet_exists",
            format!("Spreadsheet already contains a worksheet named '{name}'."),
        ));
    }
    let number = (1..=package.limits().max_entries.saturating_add(1))
        .find(|number| !package.contains_part(&format!("xl/worksheets/sheet{number}.xml")))
        .ok_or_else(|| {
            editor_error(
                "use.office.spreadsheet_sheet_limit",
                "Spreadsheet has no available native worksheet part number.",
            )
        })?;
    let sheet_part = format!("xl/worksheets/sheet{number}.xml");
    let workbook = package.xml_part("xl/workbook.xml")?;
    let index = index_xml(&workbook)?;
    let sheets = index.child("sheets", 1).ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheets_missing",
            "Spreadsheet workbook has no sheets collection.",
        )
    })?;
    let sheet_id = sheets
        .children
        .iter()
        .filter(|child| child.local_name == "sheet")
        .filter_map(|child| child.qualified_attributes.get("sheetId"))
        .filter_map(|id| id.parse::<u32>().ok())
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| {
            editor_error(
                "use.office.spreadsheet_sheet_limit",
                "Spreadsheet worksheet IDs are exhausted.",
            )
        })?;

    crate::opc_edit::add_content_type_override(
        package,
        &sheet_part,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
    )?;
    package.set_part(&sheet_part, blank_worksheet_xml().as_bytes().to_vec())?;
    let relationship_id = crate::opc_edit::add_relationship(
        package,
        "xl/_rels/workbook.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        &format!("worksheets/sheet{number}.xml"),
    )?;
    let tag = qualified(prefix(&sheets.qualified_name), "sheet");
    let fragment = format!(
        "<{tag} xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" name=\"{}\" sheetId=\"{sheet_id}\" r:id=\"{}\"/>",
        escape_attribute(name),
        escape_attribute(&relationship_id)
    );
    let edited = insert_child(&workbook, sheets, fragment)?;
    package.set_part("xl/workbook.xml", edited)?;
    Ok(format!("/{name}"))
}

pub(super) fn remove_worksheet(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    if package.kind() != DocumentKind::Spreadsheet {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Native worksheet removal is available only for Spreadsheet documents.",
        ));
    }
    let requested_name = path
        .strip_prefix('/')
        .filter(|name| !name.is_empty())
        .ok_or_else(|| {
            editor_error(
                "use.office.mutation_path_unsupported",
                "Native worksheet removal requires a path such as /Sheet2.",
            )
        })?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let worksheets = snapshot
        .root()
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Worksheet)
        .collect::<Vec<_>>();
    if worksheets.len() <= 1 {
        return Err(editor_error(
            "use.office.spreadsheet_last_sheet",
            "A Spreadsheet document must retain at least one worksheet.",
        ));
    }
    let requested = worksheets
        .into_iter()
        .find(|sheet| sheet.path[1..].eq_ignore_ascii_case(requested_name))
        .ok_or_else(|| {
            editor_error(
                "use.office.node_not_found",
                format!("Office semantic path '{path}' does not exist."),
            )
        })?;
    let part_name = requested.format.get("part").cloned().ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_invalid",
            "Spreadsheet semantic sheet has no source part.",
        )
    })?;
    let (owned_parts, owned_overrides) = worksheet::owned_worksheet_parts(package, &part_name)?;
    let workbook = package.xml_part("xl/workbook.xml")?;
    let index = index_xml(&workbook)?;
    let sheets = index.child("sheets", 1).ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheets_missing",
            "Spreadsheet workbook has no sheets collection.",
        )
    })?;
    let sheet = sheets
        .children
        .iter()
        .filter(|child| child.local_name == "sheet")
        .find(|child| {
            child
                .qualified_attributes
                .get("name")
                .is_some_and(|name| name.eq_ignore_ascii_case(requested_name))
        })
        .ok_or_else(|| {
            editor_error(
                "use.office.node_not_found",
                format!("Office semantic path '{path}' does not exist."),
            )
        })?;
    let relationship_id = sheet
        .qualified_attributes
        .iter()
        .find(|(name, _)| name.ends_with(":id"))
        .map(|(_, value)| value.clone())
        .ok_or_else(|| {
            editor_error(
                "use.office.spreadsheet_sheet_invalid",
                format!("Worksheet '{requested_name}' has no relationship ID."),
            )
        })?;
    let edited = worksheet::remove_workbook_sheet(&workbook, requested_name)?;
    worksheet::rewrite_deleted_worksheet_references(package, requested_name, &part_name)?;
    crate::opc_edit::remove_relationship(package, "xl/_rels/workbook.xml.rels", &relationship_id)?;
    for owned_part in &owned_parts {
        let relationships = worksheet::relationship_part(owned_part);
        if owned_overrides.contains(&relationships) {
            crate::opc_edit::remove_content_type_override(package, &relationships)?;
        }
        package.remove_part(&relationships)?;
        if owned_overrides.contains(owned_part) {
            crate::opc_edit::remove_content_type_override(package, owned_part)?;
        }
        package.remove_part(owned_part)?;
    }
    if package.contains_part(&part_name) {
        return Err(editor_error(
            "use.office.spreadsheet_sheet_invalid",
            format!("Worksheet part '{part_name}' could not be removed."),
        ));
    }
    package.set_part("xl/workbook.xml", edited)?;
    mark_workbook_for_recalculation(package)
}

fn validate_worksheet_name(name: &str) -> UseResult<()> {
    if name.is_empty()
        || name.chars().count() > 31
        || name.chars().any(char::is_control)
        || name
            .chars()
            .any(|character| matches!(character, '[' | ']' | ':' | '*' | '?' | '/' | '\\'))
        || name.starts_with('\'')
        || name.ends_with('\'')
    {
        return Err(editor_error(
            "use.office.spreadsheet_sheet_name_invalid",
            "Worksheet names must be 1-31 characters and exclude control characters, []:*?/\\, and edge apostrophes.",
        ));
    }
    Ok(())
}

fn blank_worksheet_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"#
}
