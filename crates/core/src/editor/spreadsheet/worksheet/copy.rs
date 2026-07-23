use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{decoded_text, require_spreadsheet, updated_start_tag};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_formula::rewrite_formula_sheet_name;
use crate::xml_edit::{apply_patches, escape_text, index_xml, IndexedXmlElement, XmlPatch};
use crate::NativeOfficePackage;

mod graph;
mod table;

use graph::{relative_target, ClonePlan};
use table::TableIdentityPlan;

const WORKSHEET_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const WORKSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

pub(crate) fn copy_worksheet(
    package: &mut NativeOfficePackage,
    path: &str,
    name: &str,
    position: Option<usize>,
) -> UseResult<String> {
    require_spreadsheet(package)?;
    super::super::validate_mutation_path(path)?;
    super::super::validate_worksheet_name(name)?;

    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let worksheets = snapshot
        .root()
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Worksheet)
        .collect::<Vec<_>>();
    if worksheets.iter().any(|sheet| {
        sheet
            .path
            .trim_start_matches('/')
            .eq_ignore_ascii_case(name)
    }) {
        return Err(super::super::editor_error(
            "use.office.spreadsheet_sheet_exists",
            format!("Spreadsheet already contains a worksheet named '{name}'."),
        ));
    }
    let source_index = worksheets
        .iter()
        .position(|sheet| sheet.path.eq_ignore_ascii_case(path))
        .ok_or_else(|| super::super::node_not_found(path))?;
    let source = worksheets[source_index];
    let source_name = source.path.trim_start_matches('/');
    let source_part = source.format.get("part").ok_or_else(|| {
        super::super::editor_error(
            "use.office.spreadsheet_sheet_invalid",
            format!("Worksheet '{source_name}' has no source part."),
        )
    })?;
    let position = position.unwrap_or(source_index + 2);
    if position == 0 || position > worksheets.len() + 1 {
        return Err(super::super::editor_error(
            "use.office.spreadsheet_sheet_position_invalid",
            format!(
                "Copied worksheet position must be within 1-{}.",
                worksheets.len() + 1
            ),
        )
        .with_detail("position", position));
    }

    let model = package.opc_model()?;
    let plan = ClonePlan::build(package, &model, source_part)?;
    let target_part = plan
        .target(source_part)
        .ok_or_else(|| copy_error("Worksheet copy plan has no target worksheet part."))?
        .to_string();
    let table_identities = TableIdentityPlan::build(package, &plan)?;
    plan.apply(package, &model, source_name, name, &table_identities)?;

    crate::opc_edit::add_content_type_override(package, &target_part, WORKSHEET_CONTENT_TYPE)?;
    let workbook_relationship = crate::opc_edit::add_relationship(
        package,
        "xl/_rels/workbook.xml.rels",
        WORKSHEET_RELATIONSHIP_TYPE,
        &relative_target("xl/workbook.xml", &target_part, None),
    )?;
    insert_workbook_sheet(package, source_name, name, position, &workbook_relationship)?;
    super::super::mark_workbook_for_recalculation(package)?;
    Ok(format!("/{name}"))
}

pub(super) fn is_shared_relationship(relationship_type: &str, part_name: &str) -> bool {
    graph::is_shared_relationship(relationship_type, part_name)
}

fn insert_workbook_sheet(
    package: &mut NativeOfficePackage,
    source_name: &str,
    target_name: &str,
    position: usize,
    relationship_id: &str,
) -> UseResult<()> {
    let workbook = package.xml_part("xl/workbook.xml")?;
    let index = index_xml(&workbook)?;
    let sheets = index.child("sheets", 1).ok_or_else(|| {
        super::super::editor_error(
            "use.office.spreadsheet_sheets_missing",
            "Spreadsheet workbook has no sheets collection.",
        )
    })?;
    let sheet_elements = sheets
        .children
        .iter()
        .filter(|sheet| sheet.local_name == "sheet")
        .collect::<Vec<_>>();
    let source_index = sheet_elements
        .iter()
        .position(|sheet| {
            sheet
                .attributes
                .get("name")
                .is_some_and(|name| name.eq_ignore_ascii_case(source_name))
        })
        .ok_or_else(|| super::super::node_not_found(&format!("/{source_name}")))?;
    let source_sheet = sheet_elements[source_index];
    let sheet_id = sheet_elements
        .iter()
        .filter_map(|sheet| sheet.attributes.get("sheetId"))
        .filter_map(|id| id.parse::<u32>().ok())
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| copy_error("Spreadsheet worksheet IDs are exhausted."))?;
    let relationship_attribute = source_sheet
        .qualified_attributes
        .keys()
        .find(|name| {
            name.rsplit_once(':')
                .is_some_and(|(_, local)| local == "id")
        })
        .cloned()
        .unwrap_or_else(|| "r:id".into());
    let sheet_fragment = updated_start_tag(
        source_sheet,
        &BTreeMap::from([
            ("name".to_string(), target_name.to_string()),
            ("sheetId".to_string(), sheet_id.to_string()),
            (relationship_attribute, relationship_id.to_string()),
        ]),
    );
    let new_index = position - 1;
    let insertion = sheet_elements
        .get(new_index)
        .map_or(sheets.content_range.end, |sheet| sheet.full_range.start);
    let mut patches = vec![XmlPatch::new(insertion..insertion, sheet_fragment)];

    let mut cloned_names = String::new();
    if let Some(defined_names) = index.child("definedNames", 1) {
        for defined_name in defined_names
            .children
            .iter()
            .filter(|name| name.local_name == "definedName")
        {
            let Some(old_index) = defined_name
                .attributes
                .get("localSheetId")
                .and_then(|value| value.parse::<usize>().ok())
            else {
                continue;
            };
            if old_index == source_index {
                let formula = decoded_text(&workbook, defined_name)?;
                let formula = rewrite_formula_sheet_name(&formula, source_name, target_name)?;
                cloned_names.push_str(&defined_name_fragment(defined_name, new_index, &formula));
            }
            if old_index >= new_index {
                patches.push(XmlPatch::new(
                    defined_name.start_tag_range.clone(),
                    updated_start_tag(
                        defined_name,
                        &BTreeMap::from([(
                            "localSheetId".to_string(),
                            (old_index + 1).to_string(),
                        )]),
                    ),
                ));
            }
        }
        if !cloned_names.is_empty() {
            patches.push(XmlPatch::new(
                defined_names.content_range.end..defined_names.content_range.end,
                cloned_names,
            ));
        }
    }

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
            if old_index >= new_index {
                updates.insert(attribute.to_string(), (old_index + 1).to_string());
            }
        }
        if !updates.is_empty() {
            patches.push(XmlPatch::new(
                view.start_tag_range.clone(),
                updated_start_tag(view, &updates),
            ));
        }
    }

    package.set_part("xl/workbook.xml", apply_patches(&workbook, patches)?)
}

fn defined_name_fragment(
    source: &IndexedXmlElement,
    local_sheet_id: usize,
    formula: &str,
) -> String {
    let start = updated_start_tag(
        source,
        &BTreeMap::from([("localSheetId".to_string(), local_sheet_id.to_string())]),
    );
    if source.empty {
        return start;
    }
    format!(
        "{start}{}</{}>",
        escape_text(formula),
        source.qualified_name
    )
}

pub(super) fn copy_error(message: impl Into<String>) -> a3s_use_core::UseError {
    super::super::editor_error("use.office.spreadsheet_copy_failed", message)
}
