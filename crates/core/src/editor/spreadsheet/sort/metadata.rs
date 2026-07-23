use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::super::editor_error;
use crate::spreadsheet_reference::CellRange;
use crate::xml_edit::{apply_patches, index_xml, IndexedXmlElement, XmlPatch};
use crate::{
    LosslessXmlPart, NativeOfficePackage, OpcPackageModel, RelationshipSource, RelationshipTarget,
};

mod references;

pub(super) fn rewrite_worksheet(
    part: &LosslessXmlPart,
    data_range: CellRange,
    old_to_new: &BTreeMap<u32, u32>,
) -> UseResult<Vec<u8>> {
    references::rewrite_worksheet(part, data_range, old_to_new)
}

pub(super) fn rewrite_related(
    package: &mut NativeOfficePackage,
    worksheet_part: &str,
    data_range: CellRange,
    old_to_new: &BTreeMap<u32, u32>,
) -> UseResult<()> {
    let opc = OpcPackageModel::read(package)?;
    let source = RelationshipSource::Part {
        part_name: worksheet_part.to_string(),
    };
    let relationships = opc.relationships().relationships_from(&source).to_vec();
    for relationship in relationships {
        let RelationshipTarget::Internal { part_name, .. } = relationship.target else {
            continue;
        };
        if relationship.relationship_type.ends_with("/comments")
            || relationship.relationship_type.ends_with("/threadedComment")
            || relationship
                .relationship_type
                .ends_with("/threadedComments")
        {
            rewrite_comment_part(package, &part_name, data_range, old_to_new)?;
        } else if relationship.relationship_type.ends_with("/vmlDrawing") {
            rewrite_vml_part(package, &part_name, data_range, old_to_new)?;
        } else if relationship.relationship_type.ends_with("/drawing") {
            rewrite_drawing_part(package, &part_name, data_range, old_to_new)?;
        } else if relationship.relationship_type.ends_with("/pivotTable") {
            return Err(metadata_error(
                "use.office.spreadsheet_sort_pivot_unsupported",
                "Native Spreadsheet sorting does not yet mutate a worksheet that owns a pivot table.",
            )
            .with_detail("part", part_name));
        }
    }
    Ok(())
}

pub(super) fn clear_chart_caches(package: &mut NativeOfficePackage) -> UseResult<()> {
    let parts = package
        .part_names()
        .filter(|part| part.starts_with("xl/charts/") && part.ends_with(".xml"))
        .map(str::to_string)
        .collect::<Vec<_>>();
    for part_name in parts {
        let part = package.xml_part(&part_name)?;
        let index = index_xml(&part)?;
        let mut patches = Vec::new();
        for name in ["numCache", "strCache", "multiLvlStrCache"] {
            let mut caches = Vec::new();
            index.descendants_named(name, &mut caches);
            patches.extend(
                caches
                    .into_iter()
                    .map(|cache| XmlPatch::new(cache.full_range.clone(), Vec::new())),
            );
        }
        if !patches.is_empty() {
            package.set_part(&part_name, apply_patches(&part, patches)?)?;
        }
    }
    Ok(())
}

fn rewrite_comment_part(
    package: &mut NativeOfficePackage,
    part_name: &str,
    data_range: CellRange,
    old_to_new: &BTreeMap<u32, u32>,
) -> UseResult<()> {
    let part = package.xml_part(part_name)?;
    let index = index_xml(&part)?;
    let mut elements = Vec::new();
    for name in ["comment", "threadedComment"] {
        index.descendants_named(name, &mut elements);
    }
    let mut patches = Vec::new();
    for element in elements {
        references::rewrite_single_attribute(element, "ref", data_range, old_to_new, &mut patches)?;
    }
    if !patches.is_empty() {
        package.set_part(part_name, apply_patches(&part, patches)?)?;
    }
    Ok(())
}

fn rewrite_vml_part(
    package: &mut NativeOfficePackage,
    part_name: &str,
    data_range: CellRange,
    old_to_new: &BTreeMap<u32, u32>,
) -> UseResult<()> {
    let part = package.xml_part(part_name)?;
    let index = index_xml(&part)?;
    let mut clients = Vec::new();
    index.descendants_named("ClientData", &mut clients);
    let mut patches = Vec::new();
    for client in clients {
        if client.attributes.get("ObjectType").map(String::as_str) != Some("Note") {
            continue;
        }
        let Some(column) = client.descendant("Column") else {
            continue;
        };
        let Some(row) = client.descendant("Row") else {
            continue;
        };
        let column_number = one_based_marker(
            parse_zero_based_text(&part, column, "VML comment column")?,
            "VML comment column",
        )?;
        let row_number = one_based_marker(
            parse_zero_based_text(&part, row, "VML comment row")?,
            "VML comment row",
        )?;
        if !(data_range.start.column..=data_range.end.column).contains(&column_number)
            || !(data_range.start.row..=data_range.end.row).contains(&row_number)
        {
            continue;
        }
        let target = *old_to_new.get(&row_number).ok_or_else(|| {
            metadata_error(
                "use.office.spreadsheet_sort_metadata_unsupported",
                "Spreadsheet sort has no target for a VML comment row.",
            )
        })?;
        patches.push(XmlPatch::new(
            row.content_range.clone(),
            (target - 1).to_string(),
        ));
        if let Some(anchor) = client.descendant("Anchor") {
            let text = element_text(&part, anchor)?;
            let mut values = text
                .split(',')
                .map(str::trim)
                .map(str::parse::<i64>)
                .collect::<Result<Vec<_>, _>>()
                .map_err(|_| {
                    metadata_error(
                        "use.office.spreadsheet_sort_metadata_unsupported",
                        format!("VML comment anchor in '{part_name}' is invalid."),
                    )
                })?;
            if values.len() == 8 {
                let delta = i64::from(target) - i64::from(row_number);
                for index in [2_usize, 6] {
                    values[index] = values[index].checked_add(delta).ok_or_else(|| {
                        metadata_error(
                            "use.office.spreadsheet_sort_metadata_unsupported",
                            "VML comment anchor row overflowed during sorting.",
                        )
                    })?;
                    if values[index] < 0 {
                        return Err(metadata_error(
                            "use.office.spreadsheet_sort_metadata_unsupported",
                            "VML comment anchor would move before row zero.",
                        ));
                    }
                }
                patches.push(XmlPatch::new(
                    anchor.content_range.clone(),
                    values
                        .into_iter()
                        .map(|value| value.to_string())
                        .collect::<Vec<_>>()
                        .join(", "),
                ));
            }
        }
    }
    if !patches.is_empty() {
        package.set_part(part_name, apply_patches(&part, patches)?)?;
    }
    Ok(())
}

fn rewrite_drawing_part(
    package: &mut NativeOfficePackage,
    part_name: &str,
    data_range: CellRange,
    old_to_new: &BTreeMap<u32, u32>,
) -> UseResult<()> {
    let part = package.xml_part(part_name)?;
    let index = index_xml(&part)?;
    let mut patches = Vec::new();
    for anchor_name in ["oneCellAnchor", "twoCellAnchor"] {
        let mut anchors = Vec::new();
        index.descendants_named(anchor_name, &mut anchors);
        for anchor in anchors {
            let from = anchor.child("from", 1).ok_or_else(|| {
                metadata_error(
                    "use.office.spreadsheet_sort_drawing_unsupported",
                    format!("Drawing anchor in '{part_name}' has no from marker."),
                )
            })?;
            let from_move = marker_move(&part, from, data_range, old_to_new)?;
            let to_move = if anchor_name == "twoCellAnchor" {
                let to = anchor.child("to", 1).ok_or_else(|| {
                    metadata_error(
                        "use.office.spreadsheet_sort_drawing_unsupported",
                        format!("Drawing anchor in '{part_name}' has no to marker."),
                    )
                })?;
                marker_move(&part, to, data_range, old_to_new)?
            } else {
                None
            };
            match (from_move, to_move) {
                (None, None) => {}
                (Some((row, target, element)), None) if anchor_name == "oneCellAnchor" => {
                    patches.push(XmlPatch::new(
                        element.content_range.clone(),
                        (target - 1).to_string(),
                    ));
                    let _ = row;
                }
                (
                    Some((from_row, from_target, from_element)),
                    Some((to_row, to_target, to_element)),
                ) if i64::from(from_target) - i64::from(from_row)
                    == i64::from(to_target) - i64::from(to_row) =>
                {
                    patches.push(XmlPatch::new(
                        from_element.content_range.clone(),
                        (from_target - 1).to_string(),
                    ));
                    patches.push(XmlPatch::new(
                        to_element.content_range.clone(),
                        (to_target - 1).to_string(),
                    ));
                }
                _ => {
                    return Err(metadata_error(
                        "use.office.spreadsheet_sort_drawing_unsupported",
                        "A drawing crosses independently permuted rows and cannot move losslessly with this sort.",
                    )
                    .with_detail("part", part_name));
                }
            }
        }
    }
    if !patches.is_empty() {
        package.set_part(part_name, apply_patches(&part, patches)?)?;
    }
    Ok(())
}

fn marker_move<'a>(
    part: &LosslessXmlPart,
    marker: &'a IndexedXmlElement,
    data_range: CellRange,
    old_to_new: &BTreeMap<u32, u32>,
) -> UseResult<Option<(u32, u32, &'a IndexedXmlElement)>> {
    let column = marker.child("col", 1).ok_or_else(|| {
        metadata_error(
            "use.office.spreadsheet_sort_drawing_unsupported",
            "Spreadsheet drawing marker has no column.",
        )
    })?;
    let row = marker.child("row", 1).ok_or_else(|| {
        metadata_error(
            "use.office.spreadsheet_sort_drawing_unsupported",
            "Spreadsheet drawing marker has no row.",
        )
    })?;
    let column_number = one_based_marker(
        parse_zero_based_text(part, column, "drawing column")?,
        "drawing column",
    )?;
    let row_number = one_based_marker(
        parse_zero_based_text(part, row, "drawing row")?,
        "drawing row",
    )?;
    if !(data_range.start.column..=data_range.end.column).contains(&column_number)
        || !(data_range.start.row..=data_range.end.row).contains(&row_number)
    {
        return Ok(None);
    }
    let target = *old_to_new.get(&row_number).ok_or_else(|| {
        metadata_error(
            "use.office.spreadsheet_sort_drawing_unsupported",
            "Spreadsheet sort has no target for a drawing row.",
        )
    })?;
    Ok(Some((row_number, target, row)))
}

fn parse_zero_based_text(
    part: &LosslessXmlPart,
    element: &IndexedXmlElement,
    label: &str,
) -> UseResult<u32> {
    element_text(part, element)?.parse::<u32>().map_err(|_| {
        metadata_error(
            "use.office.spreadsheet_sort_metadata_unsupported",
            format!("Spreadsheet {label} is invalid."),
        )
    })
}

fn one_based_marker(value: u32, label: &str) -> UseResult<u32> {
    value.checked_add(1).ok_or_else(|| {
        metadata_error(
            "use.office.spreadsheet_sort_metadata_unsupported",
            format!("Spreadsheet {label} exceeds the supported coordinate range."),
        )
    })
}

fn element_text(part: &LosslessXmlPart, element: &IndexedXmlElement) -> UseResult<String> {
    let bytes = part
        .parse_bytes()
        .get(element.content_range.clone())
        .ok_or_else(|| {
            metadata_error(
                "use.office.spreadsheet_sort_metadata_unsupported",
                "Spreadsheet metadata text range is invalid.",
            )
        })?;
    std::str::from_utf8(bytes)
        .map(str::trim)
        .map(str::to_string)
        .map_err(|_| {
            metadata_error(
                "use.office.spreadsheet_sort_metadata_unsupported",
                "Spreadsheet metadata text is not UTF-8.",
            )
        })
}

fn metadata_error(code: &str, message: impl Into<String>) -> a3s_use_core::UseError {
    editor_error(code, message)
}
