use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::metadata_error;
use crate::spreadsheet_reference::{CellRange, CellReference};
use crate::xml_edit::{apply_patches, index_xml, IndexedXmlElement, XmlPatch};
use crate::LosslessXmlPart;

const MAX_REWRITTEN_REFERENCES: usize = 4_096;
const MAX_REWRITTEN_REFERENCE_BYTES: usize = 32_767;

pub(super) fn rewrite_worksheet(
    part: &LosslessXmlPart,
    data_range: CellRange,
    old_to_new: &BTreeMap<u32, u32>,
) -> UseResult<Vec<u8>> {
    let root = index_xml(part)?;
    let mut elements = Vec::new();
    collect_descendants(&root, &mut elements);
    let mut patches = Vec::new();
    for element in elements {
        match element.local_name.as_str() {
            "hyperlink" => {
                rewrite_single_attribute(element, "ref", data_range, old_to_new, &mut patches)?
            }
            "dataValidation" | "conditionalFormatting" | "protectedRange" | "ignoredError" => {
                rewrite_list_attribute(element, "sqref", data_range, old_to_new, &mut patches)?
            }
            _ => {}
        }
    }
    if patches.is_empty() {
        Ok(part.parse_bytes().to_vec())
    } else {
        apply_patches(part, patches)
    }
}

pub(super) fn rewrite_single_attribute(
    element: &IndexedXmlElement,
    attribute: &str,
    data_range: CellRange,
    old_to_new: &BTreeMap<u32, u32>,
    patches: &mut Vec<XmlPatch>,
) -> UseResult<()> {
    let Some(value) = element.attributes.get(attribute) else {
        return Ok(());
    };
    let transformed = transform_reference(value, data_range, old_to_new)?;
    if transformed.len() != 1 {
        return Err(metadata_error(
            "use.office.spreadsheet_sort_metadata_unsupported",
            format!(
                "Spreadsheet {}@{} cannot represent a non-contiguous sorted range.",
                element.local_name, attribute
            ),
        ));
    }
    if transformed[0] != *value {
        patches.push(attribute_patch(element, attribute, &transformed[0])?);
    }
    Ok(())
}

fn rewrite_list_attribute(
    element: &IndexedXmlElement,
    attribute: &str,
    data_range: CellRange,
    old_to_new: &BTreeMap<u32, u32>,
    patches: &mut Vec<XmlPatch>,
) -> UseResult<()> {
    let Some(value) = element.attributes.get(attribute) else {
        return Ok(());
    };
    let mut transformed = Vec::new();
    for token in value.split_ascii_whitespace() {
        transformed.extend(transform_reference(token, data_range, old_to_new)?);
    }
    if transformed.len() > MAX_REWRITTEN_REFERENCES {
        return Err(metadata_error(
            "use.office.spreadsheet_sort_metadata_limit",
            format!(
                "Spreadsheet sorting would expand {}@{} beyond {MAX_REWRITTEN_REFERENCES} references.",
                element.local_name, attribute
            ),
        ));
    }
    let transformed = transformed.join(" ");
    if transformed.len() > MAX_REWRITTEN_REFERENCE_BYTES {
        return Err(metadata_error(
            "use.office.spreadsheet_sort_metadata_limit",
            format!(
                "Spreadsheet sorting would expand {}@{} beyond {MAX_REWRITTEN_REFERENCE_BYTES} bytes.",
                element.local_name, attribute
            ),
        ));
    }
    if transformed != *value {
        patches.push(attribute_patch(element, attribute, &transformed)?);
    }
    Ok(())
}

fn transform_reference(
    value: &str,
    data_range: CellRange,
    old_to_new: &BTreeMap<u32, u32>,
) -> UseResult<Vec<String>> {
    let reference = CellRange::parse(value).map_err(|error| {
        metadata_error(
            "use.office.spreadsheet_sort_metadata_unsupported",
            format!("Spreadsheet metadata reference '{value}' is invalid: {error}"),
        )
    })?;
    if !reference.intersects(data_range) || contains(reference, data_range) {
        return Ok(vec![value.to_string()]);
    }
    if !contains(data_range, reference) {
        return Err(metadata_error(
            "use.office.spreadsheet_sort_metadata_unsupported",
            format!(
                "Spreadsheet metadata reference '{value}' partially crosses the sorted data range '{}'.",
                data_range.a1()
            ),
        ));
    }
    let mut target_rows = (reference.start.row..=reference.end.row)
        .map(|row| {
            old_to_new.get(&row).copied().ok_or_else(|| {
                metadata_error(
                    "use.office.spreadsheet_sort_metadata_unsupported",
                    format!("Spreadsheet sort has no target for metadata row {row}."),
                )
            })
        })
        .collect::<UseResult<Vec<_>>>()?;
    target_rows.sort_unstable();
    target_rows.dedup();
    let mut ranges = Vec::new();
    let mut start = target_rows[0];
    let mut end = start;
    for row in target_rows.into_iter().skip(1) {
        if end.checked_add(1) == Some(row) {
            end = row;
        } else {
            ranges.push(reference_with_rows(reference, start, end));
            start = row;
            end = row;
        }
    }
    ranges.push(reference_with_rows(reference, start, end));
    Ok(ranges)
}

fn reference_with_rows(reference: CellRange, start_row: u32, end_row: u32) -> String {
    CellRange {
        start: CellReference {
            column: reference.start.column,
            row: start_row,
        },
        end: CellReference {
            column: reference.end.column,
            row: end_row,
        },
    }
    .a1()
}

fn contains(outer: CellRange, inner: CellRange) -> bool {
    outer.start.column <= inner.start.column
        && outer.end.column >= inner.end.column
        && outer.start.row <= inner.start.row
        && outer.end.row >= inner.end.row
}

fn attribute_patch(
    element: &IndexedXmlElement,
    attribute: &str,
    value: &str,
) -> UseResult<XmlPatch> {
    let mut attributes = element.qualified_attributes.clone();
    let key = if attributes.contains_key(attribute) {
        attribute.to_string()
    } else {
        attributes
            .keys()
            .find(|name| {
                name.rsplit_once(':')
                    .map_or(name.as_str(), |(_, local)| local)
                    == attribute
            })
            .cloned()
            .ok_or_else(|| {
                metadata_error(
                    "use.office.spreadsheet_sort_metadata_unsupported",
                    format!(
                        "Spreadsheet {} has no {attribute} attribute.",
                        element.local_name
                    ),
                )
            })?
    };
    attributes.insert(key, value.to_string());
    let attributes = attributes
        .into_iter()
        .map(|(name, value)| format!(" {name}=\"{}\"", quick_xml::escape::escape(&value)))
        .collect::<String>();
    let terminator = if element.empty { "/>" } else { ">" };
    Ok(XmlPatch::new(
        element.start_tag_range.clone(),
        format!("<{}{attributes}{terminator}", element.qualified_name),
    ))
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
