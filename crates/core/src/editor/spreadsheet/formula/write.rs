use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use crate::spreadsheet_reference::CellReference;
use crate::xml_edit::{insert_child, IndexedXmlElement, XmlPatch};
use crate::{NativeOfficePackage, SpreadsheetFormulaValue};

use super::super::{
    escape_attribute, expanded_element, indexed_cells_in_row, indexed_rows, prefix, qualified,
    remove_calculation_chain,
};
use super::{calculation_storage_error, CellWrite};

pub(super) fn apply_cell_writes(
    part: &crate::LosslessXmlPart,
    sheet_data: &IndexedXmlElement,
    writes: &BTreeMap<CellReference, CellWrite>,
) -> UseResult<Vec<u8>> {
    let mut by_row = BTreeMap::<u32, Vec<(CellReference, &CellWrite)>>::new();
    for (reference, write) in writes {
        by_row
            .entry(reference.row)
            .or_default()
            .push((*reference, write));
    }
    if sheet_data.empty {
        let rows = by_row
            .into_iter()
            .filter_map(|(row_number, writes)| {
                let cells = writes
                    .into_iter()
                    .filter_map(|(reference, write)| {
                        new_cell_fragment(prefix(&sheet_data.qualified_name), reference, write)
                            .transpose()
                    })
                    .collect::<UseResult<String>>();
                match cells {
                    Ok(cells) if cells.is_empty() => None,
                    Ok(cells) => {
                        let tag = qualified(prefix(&sheet_data.qualified_name), "row");
                        Some(Ok(format!("<{tag} r=\"{row_number}\">{cells}</{tag}>")))
                    }
                    Err(error) => Some(Err(error)),
                }
            })
            .collect::<UseResult<String>>()?;
        return insert_child(part, sheet_data, rows);
    }

    let rows = indexed_rows(sheet_data);
    let row_map = rows.iter().copied().collect::<BTreeMap<_, _>>();
    let mut patches = Vec::new();
    let mut insertions = BTreeMap::<usize, Vec<(u32, u32, String)>>::new();
    for (row_number, row_writes) in by_row {
        let Some(row) = row_map.get(&row_number).copied() else {
            let cells = row_writes
                .into_iter()
                .filter_map(|(reference, write)| {
                    new_cell_fragment(prefix(&sheet_data.qualified_name), reference, write)
                        .transpose()
                })
                .collect::<UseResult<String>>()?;
            if cells.is_empty() {
                continue;
            }
            let tag = qualified(prefix(&sheet_data.qualified_name), "row");
            let fragment = format!("<{tag} r=\"{row_number}\">{cells}</{tag}>");
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
            let cells = row_writes
                .into_iter()
                .filter_map(|(reference, write)| {
                    new_cell_fragment(prefix(&row.qualified_name), reference, write).transpose()
                })
                .collect::<UseResult<String>>()?;
            if !cells.is_empty() {
                patches.push(XmlPatch::new(
                    row.full_range.clone(),
                    expanded_element(row, &cells),
                ));
            }
            continue;
        }
        let cells = indexed_cells_in_row(row_number, row);
        for (reference, write) in row_writes {
            if let Some((_, cell)) = cells
                .iter()
                .find(|(existing, _)| existing.column == reference.column)
            {
                let replacement = existing_cell_fragment(part, cell, reference, write)?;
                patches.push(XmlPatch::new(
                    cell.full_range.clone(),
                    replacement.unwrap_or_default(),
                ));
                continue;
            }
            let Some(fragment) = new_cell_fragment(prefix(&row.qualified_name), reference, write)?
            else {
                continue;
            };
            let position = cells
                .iter()
                .find(|(existing, _)| existing.column > reference.column)
                .map(|(_, next)| next.full_range.start)
                .or_else(|| {
                    row.children
                        .iter()
                        .find(|child| child.local_name != "c")
                        .map(|child| child.full_range.start)
                })
                .unwrap_or(row.content_range.end);
            insertions
                .entry(position)
                .or_default()
                .push((row_number, reference.column, fragment));
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

fn new_cell_fragment(
    namespace_prefix: Option<&str>,
    reference: CellReference,
    write: &CellWrite,
) -> UseResult<Option<String>> {
    if matches!(write, CellWrite::Clear) {
        return Ok(None);
    }
    let tag = qualified(namespace_prefix, "c");
    let (value_type, content) = owned_cell_content(namespace_prefix, None, write)?;
    let value_type = value_type.map_or_else(String::new, |value_type| {
        format!(" t=\"{}\"", escape_attribute(value_type))
    });
    Ok(Some(format!(
        "<{tag} r=\"{}\"{value_type}>{content}</{tag}>",
        reference.a1()
    )))
}

fn existing_cell_fragment(
    part: &crate::LosslessXmlPart,
    cell: &IndexedXmlElement,
    reference: CellReference,
    write: &CellWrite,
) -> UseResult<Option<Vec<u8>>> {
    let preserved = preserved_cell_content(part, cell)?;
    let mut attributes = cell.qualified_attributes.clone();
    attributes.insert("r".into(), reference.a1());
    let (value_type, owned) =
        owned_cell_content(prefix(&cell.qualified_name), cell.child("f", 1), write)?;
    if let Some(value_type) = value_type {
        attributes.insert("t".into(), value_type.to_string());
    } else {
        attributes.remove("t");
    }
    if matches!(write, CellWrite::Clear)
        && attributes.len() == 1
        && attributes.contains_key("r")
        && preserved.iter().all(u8::is_ascii_whitespace)
    {
        return Ok(None);
    }
    let attributes = attributes
        .into_iter()
        .map(|(name, value)| format!(" {name}=\"{}\"", escape_attribute(&value)))
        .collect::<String>();
    let mut content = owned.into_bytes();
    content.extend_from_slice(&preserved);
    if content.is_empty() {
        return Ok(Some(
            format!("<{}{attributes}/>", cell.qualified_name).into_bytes(),
        ));
    }
    let mut output = format!("<{}{attributes}>", cell.qualified_name).into_bytes();
    output.extend_from_slice(&content);
    output.extend_from_slice(format!("</{}>", cell.qualified_name).as_bytes());
    Ok(Some(output))
}

fn preserved_cell_content(
    part: &crate::LosslessXmlPart,
    cell: &IndexedXmlElement,
) -> UseResult<Vec<u8>> {
    if cell.empty {
        return Ok(Vec::new());
    }
    let bytes = part.parse_bytes();
    let mut output = Vec::new();
    let mut cursor = cell.content_range.start;
    for child in &cell.children {
        if matches!(child.local_name.as_str(), "f" | "v" | "is") {
            output.extend_from_slice(bytes.get(cursor..child.full_range.start).ok_or_else(
                || {
                    calculation_storage_error(
                        "use.office.spreadsheet_formula_storage_invalid",
                        "Spreadsheet cell child range is invalid.",
                    )
                },
            )?);
            cursor = child.full_range.end;
        }
    }
    output.extend_from_slice(bytes.get(cursor..cell.content_range.end).ok_or_else(|| {
        calculation_storage_error(
            "use.office.spreadsheet_formula_storage_invalid",
            "Spreadsheet cell content range is invalid.",
        )
    })?);
    Ok(output)
}

fn owned_cell_content(
    namespace_prefix: Option<&str>,
    existing_formula: Option<&IndexedXmlElement>,
    write: &CellWrite,
) -> UseResult<(Option<&'static str>, String)> {
    match write {
        CellWrite::Clear => Ok((None, String::new())),
        CellWrite::Cached(value) => cached_value_content(namespace_prefix, value),
        CellWrite::Formula {
            expression,
            value,
            spill_range,
        } => {
            let formula_tag = qualified(namespace_prefix, "f");
            let mut attributes = existing_formula
                .map(|formula| formula.qualified_attributes.clone())
                .unwrap_or_default();
            let existing_type = attributes.get("t").cloned();
            attributes.remove("t");
            attributes.remove("ref");
            if let Some(spill_range) = spill_range {
                attributes.insert("t".into(), "array".into());
                attributes.insert("ref".into(), spill_range.clone());
            } else if existing_type.is_some_and(|value| value.eq_ignore_ascii_case("normal")) {
                attributes.insert("t".into(), "normal".into());
            }
            let attributes = attributes
                .into_iter()
                .map(|(name, value)| format!(" {name}=\"{}\"", escape_attribute(&value)))
                .collect::<String>();
            let formula = format!(
                "<{formula_tag}{attributes}>{}</{formula_tag}>",
                crate::xml_edit::escape_text(expression)
            );
            let (value_type, value) = cached_value_content(namespace_prefix, value)?;
            Ok((value_type, format!("{formula}{value}")))
        }
    }
}

fn cached_value_content(
    namespace_prefix: Option<&str>,
    value: &SpreadsheetFormulaValue,
) -> UseResult<(Option<&'static str>, String)> {
    let value_tag = qualified(namespace_prefix, "v");
    let (value_type, value) = match value {
        SpreadsheetFormulaValue::Blank => (None, String::new()),
        SpreadsheetFormulaValue::Number { value } => (None, value.clone()),
        SpreadsheetFormulaValue::Text { value } => (Some("str"), value.clone()),
        SpreadsheetFormulaValue::Boolean { value } => {
            (Some("b"), if *value { "1".into() } else { "0".into() })
        }
        SpreadsheetFormulaValue::Error { error } => (Some("e"), error.as_str().into()),
        SpreadsheetFormulaValue::Array { .. } => {
            return Err(calculation_storage_error(
                "use.office.spreadsheet_formula_storage_invalid",
                "Nested Spreadsheet formula arrays cannot be written to OOXML cells.",
            ))
        }
    };
    Ok((
        value_type,
        format!(
            "<{value_tag}>{}</{value_tag}>",
            crate::xml_edit::escape_text(&value)
        ),
    ))
}

pub(super) fn mark_workbook_calculated(package: &mut NativeOfficePackage) -> UseResult<()> {
    remove_calculation_chain(package)?;
    let workbook = package.xml_part("xl/workbook.xml")?;
    let index = crate::xml_edit::index_xml(&workbook)?;
    let edited = if let Some(calc) = index.child("calcPr", 1) {
        let mut attributes = calc.qualified_attributes.clone();
        attributes.insert("calcMode".into(), "auto".into());
        attributes.insert("calcCompleted".into(), "1".into());
        attributes.insert("fullCalcOnLoad".into(), "0".into());
        attributes.insert("forceFullCalc".into(), "0".into());
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
            "<{tag} calcId=\"0\" calcMode=\"auto\" calcCompleted=\"1\" fullCalcOnLoad=\"0\" forceFullCalc=\"0\"/>"
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
