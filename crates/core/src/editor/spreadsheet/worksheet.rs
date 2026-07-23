use std::collections::{BTreeMap, BTreeSet, VecDeque};

use a3s_use_core::UseResult;

use super::{
    editor_error, escape_attribute, index_xml, node_not_found, validate_mutation_path,
    validate_worksheet_name, XmlPatch,
};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_formula::{rewrite_formula_deleted_sheet, rewrite_formula_sheet_name};
use crate::xml_edit::{apply_patches, escape_text, IndexedXmlElement};
use crate::{
    DocumentKind, LosslessXmlPart, NativeOfficePackage, RelationshipSource, RelationshipTarget,
};

mod copy;

pub(crate) use copy::copy_worksheet;

pub(super) fn owned_worksheet_parts(
    package: &NativeOfficePackage,
    root_part: &str,
) -> UseResult<(BTreeSet<String>, BTreeSet<String>)> {
    let model = package.opc_model()?;
    let mut candidates = BTreeSet::from([root_part.to_string()]);
    let mut pending = VecDeque::from([root_part.to_string()]);
    while let Some(source_part) = pending.pop_front() {
        let source = RelationshipSource::Part {
            part_name: source_part,
        };
        for relationship in model.relationships().relationships_from(&source) {
            let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
                continue;
            };
            if copy::is_shared_relationship(&relationship.relationship_type, part_name) {
                continue;
            }
            if candidates.insert(part_name.clone()) {
                pending.push_back(part_name.clone());
            }
        }
    }

    let mut protected = BTreeSet::new();
    for (source, relationship) in model.relationships().relationships() {
        let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
            continue;
        };
        if part_name == root_part || !candidates.contains(part_name) {
            continue;
        }
        let inbound_is_owned = source
            .part_name()
            .is_some_and(|source| candidates.contains(source));
        if !inbound_is_owned {
            protected.insert(part_name.clone());
        }
    }

    let mut pending = protected.iter().cloned().collect::<VecDeque<_>>();
    while let Some(source_part) = pending.pop_front() {
        let source = RelationshipSource::Part {
            part_name: source_part,
        };
        for relationship in model.relationships().relationships_from(&source) {
            let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
                continue;
            };
            if candidates.contains(part_name) && protected.insert(part_name.clone()) {
                pending.push_back(part_name.clone());
            }
        }
    }
    candidates.retain(|part| part == root_part || !protected.contains(part));
    let overrides = model
        .content_types()
        .overrides()
        .map(|(part, _)| part.to_string())
        .filter(|part| {
            candidates.contains(part)
                || candidates
                    .iter()
                    .any(|source| relationship_part(source) == *part)
        })
        .collect();
    Ok((candidates, overrides))
}

pub(super) fn relationship_part(part_name: &str) -> String {
    part_name.rsplit_once('/').map_or_else(
        || format!("_rels/{part_name}.rels"),
        |(directory, file_name)| format!("{directory}/_rels/{file_name}.rels"),
    )
}

pub(super) fn rewrite_deleted_worksheet_references(
    package: &mut NativeOfficePackage,
    deleted_name: &str,
    deleted_part: &str,
) -> UseResult<()> {
    let parts = package
        .part_names()
        .filter(|part| {
            *part != deleted_part
                && part.ends_with(".xml")
                && (part.starts_with("xl/worksheets/")
                    || part.starts_with("xl/charts/")
                    || part.starts_with("xl/tables/"))
        })
        .map(str::to_string)
        .collect::<Vec<_>>();
    for part_name in parts {
        let part = package.xml_part(&part_name)?;
        let index = index_xml(&part)?;
        let mut formulas = Vec::new();
        collect_formula_elements(&index, &mut formulas);
        let mut patches = Vec::new();
        for formula in formulas {
            let text = decoded_text(&part, formula)?;
            let rewritten = rewrite_formula_deleted_sheet(&text, deleted_name)?;
            if rewritten != text {
                patches.push(XmlPatch::new(
                    formula.content_range.clone(),
                    escape_text(&rewritten),
                ));
            }
        }
        if patches.is_empty() {
            continue;
        }
        if part_name.starts_with("xl/worksheets/") {
            let mut cells = Vec::new();
            index.descendants_named("c", &mut cells);
            for cell in cells {
                if cell.children.iter().any(|child| child.local_name == "f") {
                    if let Some(value) = cell.children.iter().find(|child| child.local_name == "v")
                    {
                        patches.push(XmlPatch::new(value.full_range.clone(), Vec::new()));
                    }
                }
            }
        } else if part_name.starts_with("xl/charts/") {
            for cache_name in ["numCache", "strCache", "multiLvlStrCache"] {
                let mut caches = Vec::new();
                index.descendants_named(cache_name, &mut caches);
                patches.extend(
                    caches
                        .into_iter()
                        .map(|cache| XmlPatch::new(cache.full_range.clone(), Vec::new())),
                );
            }
        }
        package.set_part(&part_name, apply_patches(&part, patches)?)?;
    }
    Ok(())
}

pub(super) fn remove_workbook_sheet(
    workbook: &LosslessXmlPart,
    deleted_name: &str,
) -> UseResult<Vec<u8>> {
    let index = index_xml(workbook)?;
    let sheets = index.child("sheets", 1).ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheets_missing",
            "Spreadsheet workbook has no sheets collection.",
        )
    })?;
    let sheet_elements = sheets
        .children
        .iter()
        .filter(|sheet| sheet.local_name == "sheet")
        .collect::<Vec<_>>();
    let removed_index = sheet_elements
        .iter()
        .position(|sheet| {
            sheet
                .attributes
                .get("name")
                .is_some_and(|name| name.eq_ignore_ascii_case(deleted_name))
        })
        .ok_or_else(|| node_not_found(&format!("/{deleted_name}")))?;
    let mut patches = vec![XmlPatch::new(
        sheet_elements[removed_index].full_range.clone(),
        Vec::new(),
    )];

    let mut defined_names = Vec::new();
    index.descendants_named("definedName", &mut defined_names);
    for defined_name in defined_names {
        let local_sheet_id = defined_name
            .attributes
            .get("localSheetId")
            .and_then(|value| value.parse::<usize>().ok());
        if local_sheet_id == Some(removed_index) {
            patches.push(XmlPatch::new(defined_name.full_range.clone(), Vec::new()));
            continue;
        }
        if local_sheet_id.is_some_and(|index| index >= sheet_elements.len()) {
            return Err(editor_error(
                "use.office.spreadsheet_defined_name_invalid",
                format!(
                    "Defined name localSheetId {} has no worksheet.",
                    local_sheet_id.unwrap_or_default()
                ),
            ));
        }
        if let Some(local_sheet_id) = local_sheet_id.filter(|index| *index > removed_index) {
            patches.push(XmlPatch::new(
                defined_name.start_tag_range.clone(),
                updated_start_tag(
                    defined_name,
                    &BTreeMap::from([(
                        "localSheetId".to_string(),
                        (local_sheet_id - 1).to_string(),
                    )]),
                ),
            ));
        }
        let formula = decoded_text(workbook, defined_name)?;
        let rewritten = rewrite_formula_deleted_sheet(&formula, deleted_name)?;
        if rewritten != formula {
            patches.push(XmlPatch::new(
                defined_name.content_range.clone(),
                escape_text(&rewritten),
            ));
        }
    }

    let remaining_sheets = sheet_elements.len() - 1;
    let mut views = Vec::new();
    index.descendants_named("workbookView", &mut views);
    for view in views {
        let mut updates = BTreeMap::new();
        for attribute in ["activeTab", "firstSheet"] {
            let Some(old_index) = view
                .attributes
                .get(attribute)
                .and_then(|value| value.parse::<usize>().ok())
            else {
                continue;
            };
            let new_index = if old_index > removed_index {
                old_index - 1
            } else if old_index == removed_index {
                removed_index.min(remaining_sheets - 1)
            } else {
                old_index
            };
            if new_index != old_index {
                updates.insert(attribute.to_string(), new_index.to_string());
            }
        }
        if !updates.is_empty() {
            patches.push(XmlPatch::new(
                view.start_tag_range.clone(),
                updated_start_tag(view, &updates),
            ));
        }
    }
    apply_patches(workbook, patches)
}

pub(crate) fn rename_worksheet(
    package: &mut NativeOfficePackage,
    path: &str,
    name: &str,
) -> UseResult<String> {
    require_spreadsheet(package)?;
    validate_mutation_path(path)?;
    validate_worksheet_name(name)?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let worksheets = snapshot
        .root()
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Worksheet)
        .collect::<Vec<_>>();
    let requested = worksheets
        .iter()
        .copied()
        .find(|sheet| sheet.path.eq_ignore_ascii_case(path))
        .ok_or_else(|| node_not_found(path))?;
    let old_name = requested.path.trim_start_matches('/');
    if old_name == name {
        return Ok(format!("/{name}"));
    }
    if worksheets.iter().any(|sheet| {
        sheet
            .path
            .trim_start_matches('/')
            .eq_ignore_ascii_case(name)
    }) {
        return Err(editor_error(
            "use.office.spreadsheet_sheet_exists",
            format!("Spreadsheet already contains a worksheet named '{name}'."),
        ));
    }

    rewrite_formula_part_names(package, old_name, name)?;
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
        .filter(|sheet| sheet.local_name == "sheet")
        .find(|sheet| {
            sheet
                .attributes
                .get("name")
                .is_some_and(|value| value.eq_ignore_ascii_case(old_name))
        })
        .ok_or_else(|| node_not_found(path))?;
    let updates = BTreeMap::from([("name".to_string(), name.to_string())]);
    let mut patches = vec![XmlPatch::new(
        sheet.start_tag_range.clone(),
        updated_start_tag(sheet, &updates),
    )];
    let mut defined_names = Vec::new();
    index.descendants_named("definedName", &mut defined_names);
    for defined_name in defined_names {
        let text = decoded_text(&workbook, defined_name)?;
        let rewritten = rewrite_formula_sheet_name(&text, old_name, name)?;
        if rewritten != text {
            patches.push(XmlPatch::new(
                defined_name.content_range.clone(),
                escape_text(&rewritten),
            ));
        }
    }
    package.set_part("xl/workbook.xml", apply_patches(&workbook, patches)?)?;
    Ok(format!("/{name}"))
}

pub(crate) fn move_worksheet(
    package: &mut NativeOfficePackage,
    path: &str,
    position: usize,
) -> UseResult<String> {
    require_spreadsheet(package)?;
    validate_mutation_path(path)?;
    let workbook = package.xml_part("xl/workbook.xml")?;
    let index = index_xml(&workbook)?;
    let sheets = index.child("sheets", 1).ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheets_missing",
            "Spreadsheet workbook has no sheets collection.",
        )
    })?;
    let sheet_elements = sheets
        .children
        .iter()
        .filter(|sheet| sheet.local_name == "sheet")
        .collect::<Vec<_>>();
    if position == 0 || position > sheet_elements.len() {
        return Err(editor_error(
            "use.office.spreadsheet_sheet_position_invalid",
            format!(
                "Worksheet position must be within 1-{}.",
                sheet_elements.len()
            ),
        )
        .with_detail("position", position));
    }
    let old_index = sheet_elements
        .iter()
        .position(|sheet| {
            sheet
                .attributes
                .get("name")
                .is_some_and(|name| format!("/{name}").eq_ignore_ascii_case(path))
        })
        .ok_or_else(|| node_not_found(path))?;
    let new_index = position - 1;
    if old_index == new_index {
        return Ok(path.to_string());
    }

    let mut order = (0..sheet_elements.len()).collect::<Vec<_>>();
    let moved = order.remove(old_index);
    order.insert(new_index, moved);
    let parse_bytes = workbook.parse_bytes();
    let content = order
        .iter()
        .map(|old| {
            parse_bytes
                .get(sheet_elements[*old].full_range.clone())
                .ok_or_else(|| {
                    editor_error(
                        "use.office.spreadsheet_sheet_invalid",
                        "Worksheet XML range is invalid.",
                    )
                })
        })
        .collect::<UseResult<Vec<_>>>()?
        .concat();
    let mut patches = vec![XmlPatch::new(sheets.content_range.clone(), content)];
    let new_index_for_old = |old: usize| order.iter().position(|candidate| *candidate == old);

    let mut defined_names = Vec::new();
    index.descendants_named("definedName", &mut defined_names);
    for name in defined_names {
        let Some(old) = name
            .attributes
            .get("localSheetId")
            .and_then(|value| value.parse::<usize>().ok())
        else {
            continue;
        };
        let Some(new) = new_index_for_old(old) else {
            return Err(editor_error(
                "use.office.spreadsheet_defined_name_invalid",
                format!("Defined name localSheetId {old} has no worksheet."),
            ));
        };
        if new != old {
            patches.push(XmlPatch::new(
                name.start_tag_range.clone(),
                updated_start_tag(
                    name,
                    &BTreeMap::from([("localSheetId".to_string(), new.to_string())]),
                ),
            ));
        }
    }

    let mut views = Vec::new();
    index.descendants_named("workbookView", &mut views);
    for view in views {
        let mut updates = BTreeMap::new();
        for attribute in ["activeTab", "firstSheet"] {
            let Some(old) = view
                .attributes
                .get(attribute)
                .and_then(|value| value.parse::<usize>().ok())
            else {
                continue;
            };
            if let Some(new) = new_index_for_old(old) {
                if new != old {
                    updates.insert(attribute.to_string(), new.to_string());
                }
            }
        }
        if !updates.is_empty() {
            patches.push(XmlPatch::new(
                view.start_tag_range.clone(),
                updated_start_tag(view, &updates),
            ));
        }
    }

    package.set_part("xl/workbook.xml", apply_patches(&workbook, patches)?)?;
    Ok(path.to_string())
}

fn rewrite_formula_part_names(
    package: &mut NativeOfficePackage,
    old: &str,
    new: &str,
) -> UseResult<()> {
    let parts = package
        .part_names()
        .filter(|part| {
            part.starts_with("xl/")
                && part.ends_with(".xml")
                && *part != "xl/workbook.xml"
                && (part.starts_with("xl/worksheets/")
                    || part.starts_with("xl/charts/")
                    || part.starts_with("xl/tables/"))
        })
        .map(str::to_string)
        .collect::<Vec<_>>();
    for part_name in parts {
        let part = package.xml_part(&part_name)?;
        let index = index_xml(&part)?;
        let mut candidates = Vec::new();
        collect_formula_elements(&index, &mut candidates);
        let mut patches = Vec::new();
        for formula in candidates {
            let text = decoded_text(&part, formula)?;
            let rewritten = rewrite_formula_sheet_name(&text, old, new)?;
            if rewritten != text {
                patches.push(XmlPatch::new(
                    formula.content_range.clone(),
                    escape_text(&rewritten),
                ));
            }
        }
        if !patches.is_empty() {
            package.set_part(&part_name, apply_patches(&part, patches)?)?;
        }
    }
    Ok(())
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

fn updated_start_tag(element: &IndexedXmlElement, updates: &BTreeMap<String, String>) -> String {
    let mut attributes = element.qualified_attributes.clone();
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

fn require_spreadsheet(package: &NativeOfficePackage) -> UseResult<()> {
    if package.kind() == DocumentKind::Spreadsheet {
        Ok(())
    } else {
        Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Worksheet rename and move are available only for Spreadsheet documents.",
        ))
    }
}
