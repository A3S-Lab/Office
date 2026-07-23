use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::UseResult;

use super::{
    collect_descendants, decoded_text, edit_start, editor_error, escape_attribute, index_xml,
    prefix, qualified, transform_coordinate, updated_start_tag, XmlPatch,
};
use crate::spreadsheet_formula::{rewrite_formula_references, ReferenceAxis, ReferenceEdit};
use crate::spreadsheet_reference::CellRange;
use crate::xml_edit::{apply_patches, escape_text, IndexedXmlElement};
use crate::{NativeOfficePackage, OpcPackageModel, RelationshipSource, RelationshipTarget};

pub(super) fn rewrite_related_parts(
    package: &mut NativeOfficePackage,
    worksheet_part: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
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
        if relationship.relationship_type.ends_with("/table") {
            rewrite_table_part(package, &part_name, axis, edit)?;
        } else if relationship.relationship_type.ends_with("/comments") {
            rewrite_comments_part(package, &part_name, axis, edit)?;
        } else if relationship.relationship_type.ends_with("/vmlDrawing") {
            rewrite_vml_part(package, &part_name, axis, edit)?;
        } else if relationship.relationship_type.ends_with("/drawing") {
            rewrite_drawing_part(package, &part_name, axis, edit)?;
        } else if relationship.relationship_type.ends_with("/pivotTable") {
            return Err(editor_error(
                "use.office.spreadsheet_pivot_structure_unsupported",
                "Row and column structural edits on worksheets containing pivot tables are not yet lossless.",
            )
            .with_detail("part", part_name));
        }
    }
    Ok(())
}

pub(super) fn rewrite_auxiliary_formulas(
    package: &mut NativeOfficePackage,
    sheets: &[(String, String)],
    target_sheet: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<()> {
    let opc = OpcPackageModel::read(package)?;
    let mut table_owners = BTreeMap::new();
    for (sheet_name, sheet_part) in sheets {
        let source = RelationshipSource::Part {
            part_name: sheet_part.clone(),
        };
        for relationship in opc.relationships().relationships_from(&source) {
            if relationship.relationship_type.ends_with("/table") {
                if let Some(part) = relationship.target.internal_part_name() {
                    table_owners.insert(part.to_string(), sheet_name.as_str());
                }
            }
        }
    }
    let parts = package
        .part_names()
        .filter(|part| part.starts_with("xl/charts/") || part.starts_with("xl/tables/"))
        .filter(|part| part.ends_with(".xml"))
        .map(str::to_string)
        .collect::<Vec<_>>();
    for part_name in parts {
        let part = package.xml_part(&part_name)?;
        let index = index_xml(&part)?;
        let mut formulas = Vec::new();
        collect_formula_elements(&index, &mut formulas);
        let mut patches = Vec::new();
        let mut changed = false;
        for formula in formulas {
            let text = decoded_text(&part, formula)?;
            let rewritten = rewrite_formula_references(
                &text,
                table_owners.get(&part_name).copied(),
                target_sheet,
                axis,
                edit,
            )?;
            if rewritten != text {
                changed = true;
                patches.push(XmlPatch::new(
                    formula.content_range.clone(),
                    escape_text(&rewritten),
                ));
            }
        }
        if changed && part_name.starts_with("xl/charts/") {
            for cache_name in ["numCache", "strCache"] {
                let mut caches = Vec::new();
                index.descendants_named(cache_name, &mut caches);
                patches.extend(
                    caches
                        .into_iter()
                        .map(|cache| XmlPatch::new(cache.full_range.clone(), Vec::new())),
                );
            }
        }
        if !patches.is_empty() {
            package.set_part(&part_name, apply_patches(&part, patches)?)?;
        }
    }
    Ok(())
}

fn rewrite_table_part(
    package: &mut NativeOfficePackage,
    part_name: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<()> {
    let part = package.xml_part(part_name)?;
    let index = index_xml(&part)?;
    if index.local_name != "table" {
        return Err(editor_error(
            "use.office.spreadsheet_table_invalid",
            format!("Spreadsheet table part '{part_name}' has an invalid root."),
        ));
    }
    let table_reference = index.attributes.get("ref").ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_table_invalid",
            format!("Spreadsheet table part '{part_name}' has no range."),
        )
    })?;
    let table_range = CellRange::parse(table_reference)?;
    validate_table_edit(&index, table_range, axis, edit, part_name)?;
    let mut descendants = Vec::new();
    collect_descendants(&index, &mut descendants);
    let mut patches = Vec::new();
    for element in descendants.into_iter().chain(std::iter::once(&index)) {
        let Some(reference) = element.attributes.get("ref") else {
            continue;
        };
        let rewritten =
            rewrite_formula_references(reference, Some("__current__"), "__current__", axis, edit)?;
        if rewritten == "#REF!" {
            return Err(editor_error(
                "use.office.spreadsheet_table_deleted",
                "A structural edit cannot delete an entire Spreadsheet table range; remove the table first.",
            )
            .with_detail("part", part_name)
            .with_detail("reference", reference.as_str()));
        }
        if rewritten != *reference {
            let updates = BTreeMap::from([("ref".to_string(), rewritten)]);
            patches.push(XmlPatch::new(
                element.start_tag_range.clone(),
                updated_start_tag(element, &updates, &BTreeSet::new()),
            ));
        }
    }
    rewrite_table_columns(&index, table_range, axis, edit, part_name, &mut patches)?;
    if patches.is_empty() {
        Ok(())
    } else {
        package.set_part(part_name, apply_patches(&part, patches)?)
    }
}

fn validate_table_edit(
    table: &IndexedXmlElement,
    range: CellRange,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
    part_name: &str,
) -> UseResult<()> {
    let ReferenceEdit::Delete { start, count } = edit else {
        return Ok(());
    };
    let end = start + count - 1;
    match axis {
        ReferenceAxis::Row => {
            if (start..=end).contains(&range.start.row) {
                return Err(editor_error(
                    "use.office.spreadsheet_table_header_deleted",
                    "A structural edit cannot delete a Spreadsheet table header row; remove or resize the table first.",
                )
                .with_detail("part", part_name));
            }
            let has_totals = table
                .attributes
                .get("totalsRowCount")
                .and_then(|value| value.parse::<u32>().ok())
                .is_some_and(|count| count > 0);
            if has_totals && (start..=end).contains(&range.end.row) {
                return Err(editor_error(
                    "use.office.spreadsheet_table_totals_deleted",
                    "A structural edit cannot delete a Spreadsheet table totals row; disable totals or resize the table first.",
                )
                .with_detail("part", part_name));
            }
        }
        ReferenceAxis::Column => {
            let overlap_start = range.start.column.max(start);
            let overlap_end = range.end.column.min(end);
            if overlap_start == range.start.column && overlap_end == range.end.column {
                return Err(editor_error(
                    "use.office.spreadsheet_table_deleted",
                    "A structural edit cannot delete every column in a Spreadsheet table; remove the table first.",
                )
                .with_detail("part", part_name));
            }
        }
    }
    Ok(())
}

fn rewrite_table_columns(
    table: &IndexedXmlElement,
    range: CellRange,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
    part_name: &str,
    patches: &mut Vec<XmlPatch>,
) -> UseResult<()> {
    if axis != ReferenceAxis::Column {
        return Ok(());
    }
    let columns = table.child("tableColumns", 1).ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_table_invalid",
            format!("Spreadsheet table part '{part_name}' has no tableColumns collection."),
        )
    })?;
    let column_elements = columns
        .children
        .iter()
        .filter(|column| column.local_name == "tableColumn")
        .collect::<Vec<_>>();
    let expected = usize::try_from(range.end.column - range.start.column + 1).unwrap_or(usize::MAX);
    if column_elements.len() != expected {
        return Err(editor_error(
            "use.office.spreadsheet_table_invalid",
            format!(
                "Spreadsheet table part '{part_name}' declares {expected} columns but contains {}.",
                column_elements.len()
            ),
        ));
    }

    match edit {
        ReferenceEdit::Insert { at, count }
            if at > range.start.column && at <= range.end.column =>
        {
            let offset = usize::try_from(at - range.start.column).unwrap_or(usize::MAX);
            let mut next_id = column_elements
                .iter()
                .filter_map(|column| column.attributes.get("id"))
                .filter_map(|id| id.parse::<u32>().ok())
                .max()
                .unwrap_or(0);
            let mut names = column_elements
                .iter()
                .filter_map(|column| column.attributes.get("name"))
                .map(|name| name.to_ascii_lowercase())
                .collect::<BTreeSet<_>>();
            let tag = column_elements.first().map_or_else(
                || qualified(prefix(&columns.qualified_name), "tableColumn"),
                |column| column.qualified_name.clone(),
            );
            let mut fragments = String::new();
            for _ in 0..count {
                next_id = next_id.checked_add(1).ok_or_else(|| {
                    editor_error(
                        "use.office.spreadsheet_table_invalid",
                        format!("Spreadsheet table part '{part_name}' exhausted column IDs."),
                    )
                })?;
                let name = unique_table_column_name(&mut names, next_id)?;
                fragments.push_str(&format!(
                    "<{tag} id=\"{next_id}\" name=\"{}\"/>",
                    escape_attribute(&name)
                ));
            }
            let insertion = column_elements
                .get(offset)
                .map_or(columns.content_range.end, |column| column.full_range.start);
            patches.push(XmlPatch::new(insertion..insertion, fragments));
            let new_count = expected
                .checked_add(usize::try_from(count).unwrap_or(usize::MAX))
                .ok_or_else(|| {
                    editor_error(
                        "use.office.spreadsheet_table_invalid",
                        format!("Spreadsheet table part '{part_name}' column count overflowed."),
                    )
                })?;
            patches.push(XmlPatch::new(
                columns.start_tag_range.clone(),
                updated_start_tag(
                    columns,
                    &BTreeMap::from([("count".to_string(), new_count.to_string())]),
                    &BTreeSet::new(),
                ),
            ));
        }
        ReferenceEdit::Delete { start, count } => {
            let end = start + count - 1;
            let overlap_start = range.start.column.max(start);
            let overlap_end = range.end.column.min(end);
            if overlap_start > overlap_end {
                return Ok(());
            }
            let first = usize::try_from(overlap_start - range.start.column).unwrap_or(usize::MAX);
            let removed = usize::try_from(overlap_end - overlap_start + 1).unwrap_or(usize::MAX);
            for column in column_elements.iter().skip(first).take(removed) {
                patches.push(XmlPatch::new(column.full_range.clone(), Vec::new()));
            }
            let retained = expected - removed;
            patches.push(XmlPatch::new(
                columns.start_tag_range.clone(),
                updated_start_tag(
                    columns,
                    &BTreeMap::from([("count".to_string(), retained.to_string())]),
                    &BTreeSet::new(),
                ),
            ));
        }
        _ => {}
    }
    Ok(())
}

fn unique_table_column_name(names: &mut BTreeSet<String>, mut suffix: u32) -> UseResult<String> {
    loop {
        let candidate = format!("Column{suffix}");
        if names.insert(candidate.to_ascii_lowercase()) {
            return Ok(candidate);
        }
        suffix = suffix.checked_add(1).ok_or_else(|| {
            editor_error(
                "use.office.spreadsheet_table_invalid",
                "Spreadsheet table exhausted generated column names.",
            )
        })?;
    }
}

fn rewrite_comments_part(
    package: &mut NativeOfficePackage,
    part_name: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<()> {
    let part = package.xml_part(part_name)?;
    let index = index_xml(&part)?;
    let mut comments = Vec::new();
    index.descendants_named("comment", &mut comments);
    let mut patches = Vec::new();
    for comment in comments {
        let Some(reference) = comment.attributes.get("ref") else {
            continue;
        };
        let rewritten =
            rewrite_formula_references(reference, Some("__current__"), "__current__", axis, edit)?;
        if rewritten == "#REF!" {
            patches.push(XmlPatch::new(comment.full_range.clone(), Vec::new()));
        } else if rewritten != *reference {
            let updates = BTreeMap::from([("ref".to_string(), rewritten)]);
            patches.push(XmlPatch::new(
                comment.start_tag_range.clone(),
                updated_start_tag(comment, &updates, &BTreeSet::new()),
            ));
        }
    }
    if patches.is_empty() {
        Ok(())
    } else {
        package.set_part(part_name, apply_patches(&part, patches)?)
    }
}

fn rewrite_vml_part(
    package: &mut NativeOfficePackage,
    part_name: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<()> {
    let part = package.xml_part(part_name)?;
    let index = index_xml(&part)?;
    let mut shapes = Vec::new();
    index.descendants_named("shape", &mut shapes);
    let mut patches = Vec::new();
    for shape in shapes {
        let Some(client_data) = shape.descendant("ClientData") else {
            continue;
        };
        if client_data.attributes.get("ObjectType").map(String::as_str) != Some("Note") {
            continue;
        }
        let target_name = match axis {
            ReferenceAxis::Row => "Row",
            ReferenceAxis::Column => "Column",
        };
        let Some(position) = client_data.descendant(target_name) else {
            continue;
        };
        let zero_based = decoded_text(&part, position)?.parse::<u32>().map_err(|_| {
            editor_error(
                "use.office.spreadsheet_vml_invalid",
                format!("VML comment anchor in '{part_name}' has an invalid {target_name}."),
            )
        })?;
        let one_based = zero_based.checked_add(1).ok_or_else(|| {
            editor_error(
                "use.office.spreadsheet_vml_invalid",
                format!("VML comment anchor in '{part_name}' overflows {target_name}."),
            )
        })?;
        match transform_coordinate(one_based, axis, axis, edit)? {
            Some(transformed) => {
                if transformed != one_based {
                    patches.push(XmlPatch::new(
                        position.content_range.clone(),
                        (transformed - 1).to_string(),
                    ));
                }
            }
            None => patches.push(XmlPatch::new(shape.full_range.clone(), Vec::new())),
        }
    }
    if patches.is_empty() {
        Ok(())
    } else {
        package.set_part(part_name, apply_patches(&part, patches)?)
    }
}

fn rewrite_drawing_part(
    package: &mut NativeOfficePackage,
    part_name: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<()> {
    let part = package.xml_part(part_name)?;
    let index = index_xml(&part)?;
    let target_name = match axis {
        ReferenceAxis::Row => "row",
        ReferenceAxis::Column => "col",
    };
    let mut markers = Vec::new();
    for marker_name in ["from", "to"] {
        let mut candidates = Vec::new();
        index.descendants_named(marker_name, &mut candidates);
        for marker in candidates {
            if let Some(coordinate) = marker.child(target_name, 1) {
                markers.push(coordinate);
            }
        }
    }
    let mut patches = Vec::new();
    for coordinate in markers {
        let zero_based = decoded_text(&part, coordinate)?
            .parse::<u32>()
            .map_err(|_| {
                editor_error(
                    "use.office.spreadsheet_drawing_invalid",
                    format!("Drawing anchor in '{part_name}' has an invalid {target_name}."),
                )
            })?;
        let one_based = zero_based.checked_add(1).ok_or_else(|| {
            editor_error(
                "use.office.spreadsheet_drawing_invalid",
                format!("Drawing anchor in '{part_name}' overflows {target_name}."),
            )
        })?;
        let transformed =
            transform_coordinate(one_based, axis, axis, edit)?.unwrap_or_else(|| edit_start(edit));
        if transformed != one_based {
            patches.push(XmlPatch::new(
                coordinate.content_range.clone(),
                (transformed - 1).to_string(),
            ));
        }
    }
    if patches.is_empty() {
        Ok(())
    } else {
        package.set_part(part_name, apply_patches(&part, patches)?)
    }
}

fn collect_formula_elements<'a>(
    element: &'a IndexedXmlElement,
    output: &mut Vec<&'a IndexedXmlElement>,
) {
    for child in &element.children {
        if matches!(
            child.local_name.as_str(),
            "f" | "formula" | "calculatedColumnFormula" | "totalsRowFormula"
        ) {
            output.push(child);
        }
        collect_formula_elements(child, output);
    }
}
