use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{editor_error, validate_mutation_path, SpreadsheetCellValue};
use crate::editor::part::{dialect, relationship_part, relative_target};
use crate::editor::NativeSpreadsheetTable;
use crate::semantic::{DocumentNode, NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_formula::StructuredReferenceRewritePlan;
use crate::spreadsheet_reference::{CellRange, CellReference};
use crate::xml_edit::index_xml;
use crate::{DocumentKind, NativeOfficePackage, RelationshipSource, RelationshipTarget};

mod formula;
mod xml;

const MAX_SPREADSHEET_TABLES: usize = 65_536;

struct ResolvedSheet {
    path: String,
    part: String,
}

struct ResolvedTable {
    sheet: ResolvedSheet,
    path: String,
    relationship_id: String,
    part: String,
}

pub(super) fn is_path(path: &str) -> bool {
    path.rsplit_once('/')
        .is_some_and(|(_, segment)| parse_table_segment(segment).is_some())
}

pub(super) fn add(
    package: &mut NativeOfficePackage,
    sheet: &str,
    table: &NativeSpreadsheetTable,
) -> UseResult<String> {
    let resolved = resolve_sheet(package, sheet)?;
    let mut table = table.clone();
    let range = table.validate()?;
    table.range = range.a1();
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let workbook_table_count = spreadsheet_tables(&snapshot).count();
    if workbook_table_count >= MAX_SPREADSHEET_TABLES {
        return Err(editor_error(
            "use.office.spreadsheet_table_limit",
            format!(
                "Spreadsheet already has the maximum {MAX_SPREADSHEET_TABLES} native table definitions."
            ),
        ));
    }
    let sheet_table_count = snapshot
        .get(&resolved.path, 1)?
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Table)
        .count();
    validate_identity(&snapshot, &table, None)?;
    validate_range(package, &snapshot, &resolved, range, None)?;

    let table_id = next_table_id(package)?;
    let table_part = allocate_table_part(package)?;
    let table_xml = xml::new_table_xml(dialect(package)?, table_id, &table, range)?;
    crate::LosslessXmlPart::parse(table_part.clone(), table_xml.clone())?;
    crate::opc_edit::add_content_type_override(package, &table_part, xml::content_type())?;
    package.set_part(&table_part, table_xml)?;
    let office_dialect = dialect(package)?;
    let relationship_id = crate::opc_edit::add_relationship(
        package,
        &relationship_part(&resolved.part),
        &office_dialect.relationship_type("table"),
        &relative_target(&resolved.part, &table_part),
    )?;
    let worksheet = package.xml_part(&resolved.part)?;
    let worksheet = xml::add_table_part_reference(
        &worksheet,
        office_dialect.relationship_namespace(),
        &relationship_id,
    )?;
    package.set_part(&resolved.part, worksheet)?;
    stamp_headers(package, &resolved, &table, range)?;
    super::mark_workbook_for_recalculation(package)?;
    Ok(format!(
        "{}/table[{}]",
        resolved.path,
        sheet_table_count + 1
    ))
}

pub(super) fn set(
    package: &mut NativeOfficePackage,
    path: &str,
    table: &NativeSpreadsheetTable,
) -> UseResult<String> {
    let resolved = resolve_table(package, path)?;
    validate_relationship_graph(package, &resolved)?;
    let mut table = table.clone();
    let range = table.validate()?;
    table.range = range.a1();
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let node = snapshot.get(&resolved.path, 3)?;
    if node.format.get("nativeMutable").map(String::as_str) != Some("true") {
        return Err(editor_error(
            "use.office.spreadsheet_table_unknown_content",
            format!(
                "Spreadsheet table '{}' contains formulas, totals metadata, a custom style, or other unsupported content.",
                resolved.path
            ),
        )
        .with_suggestion(
            "Keep the imported table unchanged or inspect its OOXML before replacing it through the typed table contract.",
        ));
    }
    let old_table = NativeSpreadsheetTable::from_semantic_node(&node)?;
    let old_range = CellRange::parse(&old_table.range)?;
    validate_identity(&snapshot, &table, Some(&resolved.path))?;
    validate_range(
        package,
        &snapshot,
        &resolved.sheet,
        range,
        Some(&resolved.path),
    )?;
    let (plan, formula_rewrite_required) = table_formula_rewrite_plan(
        &old_table,
        &table,
        old_range,
        range,
        resolved.sheet.path.trim_start_matches('/'),
    );
    let mut candidate = package.clone();
    if formula_rewrite_required {
        formula::rewrite_table_references(
            &mut candidate,
            &resolved.sheet.part,
            &resolved.part,
            old_range,
            &plan,
        )?;
    }
    let part = candidate.xml_part(&resolved.part)?;
    let edited = xml::replace_table(&part, &table, range)?;
    candidate.set_part(&resolved.part, edited)?;
    stamp_headers(&mut candidate, &resolved.sheet, &table, range)?;
    super::mark_workbook_for_recalculation(&mut candidate)?;
    *package = candidate;
    Ok(resolved.path)
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let resolved = resolve_table(package, path)?;
    validate_relationship_graph(package, &resolved)?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let node = snapshot.get(&resolved.path, 0)?;
    let name = required_table_format(&node, "name")?.to_string();
    let display_name = required_table_format(&node, "displayName")?.to_string();
    let range = CellRange::parse(required_table_format(&node, "ref")?)?;
    let plan = StructuredReferenceRewritePlan::removal(
        name.clone(),
        resolved.sheet.path.trim_start_matches('/'),
        [name, display_name],
    );
    let mut candidate = package.clone();
    formula::rewrite_table_references(
        &mut candidate,
        &resolved.sheet.part,
        &resolved.part,
        range,
        &plan,
    )?;
    let worksheet = candidate.xml_part(&resolved.sheet.part)?;
    let worksheet = xml::remove_table_part_reference(&worksheet, &resolved.relationship_id)?;
    candidate.set_part(&resolved.sheet.part, worksheet)?;
    crate::opc_edit::remove_relationship(
        &mut candidate,
        &relationship_part(&resolved.sheet.part),
        &resolved.relationship_id,
    )?;
    let content_types = candidate.opc_model()?.content_types().clone();
    if content_types.override_for_part(&resolved.part).is_some() {
        crate::opc_edit::remove_content_type_override(&mut candidate, &resolved.part)?;
    }
    candidate.remove_part(&resolved.part)?;
    let table_relationships = relationship_part(&resolved.part);
    if candidate.contains_part(&table_relationships) {
        candidate.remove_part(&table_relationships)?;
    }
    super::mark_workbook_for_recalculation(&mut candidate)?;
    *package = candidate;
    Ok(())
}

fn table_formula_rewrite_plan(
    old: &NativeSpreadsheetTable,
    new: &NativeSpreadsheetTable,
    old_range: CellRange,
    new_range: CellRange,
    sheet: &str,
) -> (StructuredReferenceRewritePlan, bool) {
    let old_display_name = old.display_name.as_deref().unwrap_or(&old.name);
    let new_display_name = new.display_name.as_deref().unwrap_or(&new.name);
    let mut aliases = BTreeMap::new();
    aliases.insert(old.name.to_lowercase(), new.name.clone());
    aliases.insert(
        old_display_name.to_lowercase(),
        new_display_name.to_string(),
    );

    let mut columns = BTreeMap::new();
    for (index, old_column) in old.columns.iter().enumerate() {
        let replacement = new.columns.get(index).map(|column| column.name.clone());
        if replacement.as_deref() != Some(old_column.name.as_str()) {
            columns.insert(old_column.name.to_lowercase(), replacement);
        }
    }

    let aliases_changed = if old.name.eq_ignore_ascii_case(old_display_name) {
        old_display_name != new_display_name
    } else {
        old.name != new.name || old_display_name != new_display_name
    };
    let geometry_changed = old_range != new_range
        || old.header_row != new.header_row
        || old.totals_row != new.totals_row;
    let rewrite_required = aliases_changed || !columns.is_empty() || geometry_changed;
    (
        StructuredReferenceRewritePlan::rename(
            old.name.clone(),
            sheet,
            aliases,
            columns,
            geometry_changed,
        ),
        rewrite_required,
    )
}

fn resolve_sheet(package: &NativeOfficePackage, requested: &str) -> UseResult<ResolvedSheet> {
    if package.kind() != DocumentKind::Spreadsheet {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Spreadsheet table operations are available only for Spreadsheet documents.",
        ));
    }
    validate_mutation_path(requested)?;
    if requested.trim_start_matches('/').contains('/') {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Adding a Spreadsheet table requires a worksheet path such as /Sheet1.",
        ));
    }
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
    Ok(ResolvedSheet {
        path: sheet.path.clone(),
        part,
    })
}

fn resolve_table(package: &NativeOfficePackage, requested: &str) -> UseResult<ResolvedTable> {
    validate_mutation_path(requested)?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let node = snapshot.get(requested, 0)?;
    if node.node_type != OfficeNodeType::Table {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Spreadsheet table updates require a path such as /Sheet1/table[1].",
        ));
    }
    let (sheet_path, segment) = node.path.rsplit_once('/').ok_or_else(|| {
        editor_error(
            "use.office.mutation_path_unsupported",
            "Spreadsheet table updates require a path such as /Sheet1/table[1].",
        )
    })?;
    let index = parse_table_segment(segment).ok_or_else(|| {
        editor_error(
            "use.office.mutation_path_unsupported",
            "Spreadsheet table updates require a path such as /Sheet1/table[1].",
        )
    })?;
    let sheet = resolve_sheet(package, sheet_path)?;
    let worksheet = package.xml_part(&sheet.part)?;
    let worksheet_index = index_xml(&worksheet)?;
    let collections = worksheet_index
        .children
        .iter()
        .filter(|child| {
            child.local_name == "tableParts" && child.namespace == worksheet_index.namespace
        })
        .collect::<Vec<_>>();
    if collections.len() != 1 {
        return Err(table_error(
            &sheet.part,
            "does not contain exactly one tableParts collection",
        ));
    }
    let entries = collections[0]
        .children
        .iter()
        .filter(|child| {
            child.local_name == "tablePart" && child.namespace == worksheet_index.namespace
        })
        .collect::<Vec<_>>();
    let entry = entries
        .get(index - 1)
        .ok_or_else(|| node_not_found(&node.path))?;
    let relationship_ids = entry
        .qualified_attributes
        .iter()
        .filter(|(name, _)| name.ends_with(":id"))
        .map(|(_, value)| value)
        .collect::<Vec<_>>();
    if relationship_ids.len() != 1 {
        return Err(table_error(
            &sheet.part,
            "contains a tablePart without exactly one qualified relationship ID",
        ));
    }
    let relationship_id = relationship_ids[0].clone();
    let source = RelationshipSource::Part {
        part_name: sheet.part.clone(),
    };
    let model = package.opc_model()?;
    let relationship = model
        .relationships()
        .relationship(&source, &relationship_id)
        .ok_or_else(|| {
            table_error(
                &sheet.part,
                format!("references missing table relationship '{relationship_id}'"),
            )
        })?;
    if !relationship.relationship_type.ends_with("/table") {
        return Err(table_error(
            &sheet.part,
            format!("relationship '{relationship_id}' is not a table relationship"),
        ));
    }
    let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
        return Err(table_error(
            &sheet.part,
            "uses an external table relationship",
        ));
    };
    Ok(ResolvedTable {
        sheet,
        path: node.path,
        relationship_id,
        part: part_name.clone(),
    })
}

fn parse_table_segment(segment: &str) -> Option<usize> {
    let (name, index) = segment.split_once('[')?;
    if !matches!(name.to_ascii_lowercase().as_str(), "table" | "listobject") {
        return None;
    }
    index
        .strip_suffix(']')?
        .parse::<usize>()
        .ok()
        .filter(|index| *index > 0)
}

fn validate_identity(
    snapshot: &NativeOfficeDocument,
    table: &NativeSpreadsheetTable,
    excluded_path: Option<&str>,
) -> UseResult<()> {
    let requested = [
        table.name.as_str(),
        table.display_name.as_deref().unwrap_or(&table.name),
    ];
    for existing in spreadsheet_tables(snapshot) {
        if excluded_path.is_some_and(|path| existing.path.eq_ignore_ascii_case(path)) {
            continue;
        }
        for key in ["name", "displayName"] {
            if let Some(existing_name) = existing.format.get(key) {
                if requested
                    .iter()
                    .any(|candidate| candidate.eq_ignore_ascii_case(existing_name))
                {
                    return Err(identity_collision(
                        requested
                            .iter()
                            .find(|candidate| candidate.eq_ignore_ascii_case(existing_name))
                            .copied()
                            .unwrap_or(&table.name),
                        &existing.path,
                    ));
                }
            }
        }
    }
    for collection in snapshot
        .root()
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::NamedRangeCollection)
    {
        for named_range in &collection.children {
            if let Some(name) = named_range.format.get("name") {
                if requested
                    .iter()
                    .any(|candidate| candidate.eq_ignore_ascii_case(name))
                {
                    return Err(identity_collision(name, &named_range.path));
                }
            }
        }
    }
    Ok(())
}

fn validate_range(
    package: &NativeOfficePackage,
    snapshot: &NativeOfficeDocument,
    sheet: &ResolvedSheet,
    range: CellRange,
    excluded_path: Option<&str>,
) -> UseResult<()> {
    let sheet_node = snapshot.get(&sheet.path, 1)?;
    for existing in sheet_node
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Table)
    {
        if excluded_path.is_some_and(|path| existing.path.eq_ignore_ascii_case(path)) {
            continue;
        }
        if let Some(existing_range) = existing
            .format
            .get("ref")
            .and_then(|value| CellRange::parse(value).ok())
        {
            if range.intersects(existing_range) {
                return Err(editor_error(
                    "use.office.spreadsheet_table_overlap",
                    format!(
                        "Spreadsheet table range '{}' overlaps table '{}' at '{}'.",
                        range.a1(),
                        existing.path,
                        existing_range.a1()
                    ),
                )
                .with_detail("requested", range.a1())
                .with_detail("existing", existing_range.a1()));
            }
        }
    }
    let worksheet = package.xml_part(&sheet.part)?;
    let index = index_xml(&worksheet)?;
    for merges in index
        .children
        .iter()
        .filter(|child| child.local_name == "mergeCells" && child.namespace == index.namespace)
    {
        for merged in merges
            .children
            .iter()
            .filter(|child| child.local_name == "mergeCell" && child.namespace == index.namespace)
        {
            if let Some(reference) = merged.attributes.get("ref") {
                let merged_range = CellRange::parse(reference)?;
                if range.intersects(merged_range) {
                    return Err(editor_error(
                        "use.office.spreadsheet_table_merge_overlap",
                        format!(
                            "Spreadsheet table range '{}' overlaps merged range '{}'.",
                            range.a1(),
                            merged_range.a1()
                        ),
                    ));
                }
            }
        }
    }
    for filter in index
        .children
        .iter()
        .filter(|child| child.local_name == "autoFilter" && child.namespace == index.namespace)
    {
        if let Some(reference) = filter.attributes.get("ref") {
            let filter_range = CellRange::parse(reference)?;
            if range.intersects(filter_range) {
                return Err(editor_error(
                    "use.office.spreadsheet_table_filter_overlap",
                    format!(
                        "Spreadsheet table range '{}' overlaps worksheet AutoFilter range '{}'.",
                        range.a1(),
                        filter_range.a1()
                    ),
                )
                .with_suggestion(
                    "Remove or move the worksheet AutoFilter before creating the table; the table owns its own header filter.",
                ));
            }
        }
    }
    Ok(())
}

fn stamp_headers(
    package: &mut NativeOfficePackage,
    sheet: &ResolvedSheet,
    table: &NativeSpreadsheetTable,
    range: CellRange,
) -> UseResult<()> {
    if !table.header_row {
        return Ok(());
    }
    for (offset, column) in table.columns.iter().enumerate() {
        let offset = u32::try_from(offset).map_err(|_| {
            editor_error(
                "use.office.spreadsheet_table_column_count_invalid",
                "Spreadsheet table column index does not fit OOXML limits.",
            )
        })?;
        let reference = CellReference {
            column: range.start.column + offset,
            row: range.start.row,
        };
        super::set_cell_value(
            package,
            &format!("{}/{}", sheet.path, reference.a1()),
            &SpreadsheetCellValue::Text {
                value: column.name.clone(),
            },
        )?;
    }
    Ok(())
}

fn next_table_id(package: &NativeOfficePackage) -> UseResult<u32> {
    let opc = package.opc_model()?;
    let mut maximum = 0_u32;
    for part_name in package
        .part_names()
        .filter(|name| opc.content_types().content_type(name) == Some(xml::content_type()))
    {
        let part = package.xml_part(part_name)?;
        let root = index_xml(&part)?;
        if root.local_name != "table" {
            return Err(table_error(part_name, "has an unexpected root element"));
        }
        let id = root
            .attributes
            .get("id")
            .ok_or_else(|| table_error(part_name, "has no table ID"))?
            .parse::<u32>()
            .map_err(|error| table_error(part_name, format!("has invalid table ID: {error}")))?;
        maximum = maximum.max(id);
    }
    maximum.checked_add(1).ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_table_id_exhausted",
            "Spreadsheet table IDs are exhausted.",
        )
    })
}

fn allocate_table_part(package: &NativeOfficePackage) -> UseResult<String> {
    (1..=package.limits().max_entries.saturating_add(1))
        .map(|number| format!("xl/tables/table{number}.xml"))
        .find(|candidate| !package.contains_part(candidate))
        .ok_or_else(|| {
            editor_error(
                "use.office.spreadsheet_table_part_exhausted",
                "No available Spreadsheet table part name remains.",
            )
        })
}

fn validate_relationship_graph(
    package: &NativeOfficePackage,
    resolved: &ResolvedTable,
) -> UseResult<()> {
    let model = package.opc_model()?;
    let table_source = RelationshipSource::Part {
        part_name: resolved.part.clone(),
    };
    if !model
        .relationships()
        .relationships_from(&table_source)
        .is_empty()
    {
        return Err(editor_error(
            "use.office.spreadsheet_table_relationships_unsupported",
            "The Spreadsheet table owns relationships outside the native typed table lifecycle.",
        )
        .with_detail("part", resolved.part.clone()));
    }
    let expected_source = RelationshipSource::Part {
        part_name: resolved.sheet.part.clone(),
    };
    for (source, relationship) in model.relationships().relationships() {
        if relationship.target.internal_part_name() == Some(resolved.part.as_str())
            && !(*source == expected_source && relationship.id == resolved.relationship_id)
        {
            return Err(editor_error(
                "use.office.spreadsheet_table_relationships_unsupported",
                "The Spreadsheet table has an unexpected additional inbound relationship.",
            )
            .with_detail("part", resolved.part.clone())
            .with_detail("relationshipId", relationship.id.clone()));
        }
    }
    Ok(())
}

fn spreadsheet_tables(document: &NativeOfficeDocument) -> impl Iterator<Item = &DocumentNode> {
    document
        .root()
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Worksheet)
        .flat_map(|sheet| sheet.children.iter())
        .filter(|node| node.node_type == OfficeNodeType::Table)
}

fn required_table_format<'a>(node: &'a DocumentNode, key: &str) -> UseResult<&'a str> {
    node.format.get(key).map(String::as_str).ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_table_invalid",
            format!("Spreadsheet table '{}' has no '{key}' property.", node.path),
        )
        .with_detail("path", node.path.clone())
    })
}

fn identity_collision(name: &str, owner: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_table_name_collision",
        format!("Spreadsheet table identity '{name}' collides with '{owner}'."),
    )
    .with_detail("name", name)
    .with_detail("owner", owner)
}

fn node_not_found(path: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.node_not_found",
        format!("Office semantic path '{path}' does not exist."),
    )
    .with_detail("path", path)
}

pub(super) fn table_error(part_name: &str, reason: impl Into<String>) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_table_invalid",
        format!("Spreadsheet table part '{part_name}' {}.", reason.into()),
    )
    .with_detail("part", part_name)
}
