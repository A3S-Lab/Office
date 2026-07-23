use a3s_use_core::UseResult;

use super::{editor_error, mark_workbook_for_recalculation, prefix, qualified};
use crate::editor::{NativeSpreadsheetNamedRange, NativeSpreadsheetNamedRangeScope};
use crate::spreadsheet_named_range::{
    canonical_named_range_path, is_protected_named_range, named_range_scope_label,
    parse_named_range_path, protected_named_range_error, quote_sheet_name,
    validate_named_range_comment, validate_named_range_name, validate_named_range_reference,
    NamedRangePathSelector, MAX_SPREADSHEET_NAMED_RANGES,
};
use crate::spreadsheet_reference::CellRange;
use crate::xml_edit::{
    apply_patches, decoded_element_text, escape_attribute, escape_text, index_xml,
    insert_ordered_child, IndexedXmlElement, XmlPatch,
};
use crate::{DocumentKind, LosslessXmlPart, NativeOfficePackage};

const WORKBOOK_CHILDREN_AFTER_DEFINED_NAMES: &[&str] = &[
    "calcPr",
    "oleSize",
    "customWorkbookViews",
    "pivotCaches",
    "smartTagPr",
    "smartTagTypes",
    "webPublishing",
    "fileRecoveryPr",
    "webPublishObjects",
    "extLst",
];

struct ExistingName<'a> {
    element: &'a IndexedXmlElement,
    name: String,
    scope: String,
    local_sheet_id: Option<usize>,
}

struct NormalizedName {
    value: NativeSpreadsheetNamedRange,
    scope: String,
    scope_label: String,
    local_sheet_id: Option<usize>,
}

pub(super) fn is_path(path: &str) -> bool {
    parse_named_range_path(path).ok().flatten().is_some()
}

pub(super) fn add(
    package: &mut NativeOfficePackage,
    named_range: &NativeSpreadsheetNamedRange,
) -> UseResult<String> {
    require_spreadsheet(package)?;
    let workbook = package.xml_part("xl/workbook.xml")?;
    let index = index_xml(&workbook)?;
    let sheet_names = sheet_names(&index)?;
    let normalized = normalize(named_range, &sheet_names)?;
    let container = defined_names_container(&index)?;
    let existing = existing_names(&index, container, &sheet_names)?;
    if existing.len() >= MAX_SPREADSHEET_NAMED_RANGES {
        return Err(editor_error(
            "use.office.spreadsheet_named_range_limit",
            format!(
                "Spreadsheet workbook already has the maximum {MAX_SPREADSHEET_NAMED_RANGES} defined names."
            ),
        ));
    }
    reject_duplicate(&normalized, &existing, None)?;
    reject_table_collision(package, &normalized.value.name)?;

    let fragment = name_fragment(prefix(&index.qualified_name), None, &normalized);
    let edited = if let Some(container) = container {
        validate_container_content(&workbook, &index, container)?;
        apply_patches(
            &workbook,
            vec![XmlPatch::new(
                container.content_range.end..container.content_range.end,
                fragment,
            )],
        )?
    } else {
        let tag = qualified(prefix(&index.qualified_name), "definedNames");
        insert_ordered_child(
            &workbook,
            &index,
            format!("<{tag}>{fragment}</{tag}>"),
            WORKBOOK_CHILDREN_AFTER_DEFINED_NAMES,
        )?
    };
    package.set_part("xl/workbook.xml", edited)?;
    mark_workbook_for_recalculation(package)?;
    Ok(canonical_named_range_path(
        &normalized.value.name,
        &normalized.scope_label,
    ))
}

pub(super) fn set(
    package: &mut NativeOfficePackage,
    path: &str,
    named_range: &NativeSpreadsheetNamedRange,
) -> UseResult<String> {
    require_spreadsheet(package)?;
    let workbook = package.xml_part("xl/workbook.xml")?;
    let index = index_xml(&workbook)?;
    let sheet_names = sheet_names(&index)?;
    let container = defined_names_container(&index)?.ok_or_else(|| not_found(path))?;
    validate_container_content(&workbook, &index, container)?;
    let existing = existing_names(&index, Some(container), &sheet_names)?;
    let target_index = resolve_target(path, &existing)?;
    let target = &existing[target_index];
    if is_protected_named_range(&target.name) {
        return Err(protected_named_range_error(&target.name));
    }
    validate_editable_name_content(&workbook, target.element)?;

    let normalized = normalize(named_range, &sheet_names)?;
    reject_duplicate(&normalized, &existing, Some(target_index))?;
    reject_table_collision(package, &normalized.value.name)?;
    let replacement = name_fragment(
        prefix(&target.element.qualified_name),
        Some(target.element),
        &normalized,
    );
    let edited = apply_patches(
        &workbook,
        vec![XmlPatch::new(
            target.element.full_range.clone(),
            replacement,
        )],
    )?;
    package.set_part("xl/workbook.xml", edited)?;
    mark_workbook_for_recalculation(package)?;
    Ok(canonical_named_range_path(
        &normalized.value.name,
        &normalized.scope_label,
    ))
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    require_spreadsheet(package)?;
    let workbook = package.xml_part("xl/workbook.xml")?;
    let index = index_xml(&workbook)?;
    let sheet_names = sheet_names(&index)?;
    let container = defined_names_container(&index)?.ok_or_else(|| not_found(path))?;
    validate_container_content(&workbook, &index, container)?;
    let existing = existing_names(&index, Some(container), &sheet_names)?;
    let target = &existing[resolve_target(path, &existing)?];
    if is_protected_named_range(&target.name) {
        return Err(protected_named_range_error(&target.name));
    }
    validate_editable_name_content(&workbook, target.element)?;

    let edited = if existing.len() == 1 {
        if !container.qualified_attributes.is_empty() {
            return Err(editor_error(
                "use.office.spreadsheet_named_range_unknown_content",
                "The final named range cannot be removed without discarding unknown definedNames attributes.",
            )
            .with_detail("part", "xl/workbook.xml")
            .with_detail("path", path));
        }
        apply_patches(
            &workbook,
            vec![XmlPatch::new(container.full_range.clone(), Vec::new())],
        )?
    } else {
        apply_patches(
            &workbook,
            vec![XmlPatch::new(target.element.full_range.clone(), Vec::new())],
        )?
    };
    package.set_part("xl/workbook.xml", edited)?;
    mark_workbook_for_recalculation(package)
}

fn require_spreadsheet(package: &NativeOfficePackage) -> UseResult<()> {
    if package.kind() != DocumentKind::Spreadsheet {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Named-range operations are available only for Spreadsheet documents.",
        ));
    }
    Ok(())
}

fn normalize(
    named_range: &NativeSpreadsheetNamedRange,
    sheet_names: &[String],
) -> UseResult<NormalizedName> {
    validate_named_range_name(&named_range.name)?;
    if is_protected_named_range(&named_range.name) {
        return Err(protected_named_range_error(&named_range.name));
    }
    validate_named_range_reference(&named_range.reference)?;
    validate_named_range_comment(named_range.comment.as_deref())?;

    let (scope, local_sheet_id) = match &named_range.scope {
        NativeSpreadsheetNamedRangeScope::Workbook => ("workbook".to_string(), None),
        NativeSpreadsheetNamedRangeScope::Worksheet(requested) => {
            let (index, canonical) = sheet_names
                .iter()
                .enumerate()
                .find(|(_, name)| name.eq_ignore_ascii_case(requested))
                .ok_or_else(|| {
                    editor_error(
                        "use.office.spreadsheet_named_range_scope_invalid",
                        format!(
                            "Spreadsheet named-range scope worksheet '{requested}' does not exist."
                        ),
                    )
                    .with_detail("scope", requested.as_str())
                })?;
            (canonical.clone(), Some(index))
        }
    };
    let mut value = named_range.clone();
    value.scope = local_sheet_id.map_or(NativeSpreadsheetNamedRangeScope::Workbook, |_| {
        NativeSpreadsheetNamedRangeScope::worksheet(scope.clone())
    });
    if is_unqualified_a1_range(&value.reference) {
        let Some(_) = local_sheet_id else {
            return Err(editor_error(
                "use.office.spreadsheet_named_range_ref_invalid",
                "Workbook-scoped named ranges require a sheet-qualified A1 ref.",
            )
            .with_suggestion("Use a ref such as 'Sheet1'!$A$1:$A$10."));
        };
        value.reference = format!("{}!{}", quote_sheet_name(&scope), value.reference);
    }
    validate_simple_qualified_sheet(&value.reference, sheet_names)?;
    validate_named_range_reference(&value.reference)?;
    let scope_label = named_range_scope_label(&scope, local_sheet_id.is_some());
    Ok(NormalizedName {
        value,
        scope,
        scope_label,
        local_sheet_id,
    })
}

fn sheet_names(workbook: &IndexedXmlElement) -> UseResult<Vec<String>> {
    let sheets = workbook
        .children
        .iter()
        .find(|child| child.local_name == "sheets" && child.namespace == workbook.namespace)
        .ok_or_else(|| {
            editor_error(
                "use.office.spreadsheet_sheets_missing",
                "Spreadsheet workbook has no sheets collection.",
            )
        })?;
    sheets
        .children
        .iter()
        .filter(|sheet| sheet.local_name == "sheet" && sheet.namespace == workbook.namespace)
        .map(|sheet| {
            sheet
                .qualified_attributes
                .get("name")
                .cloned()
                .ok_or_else(|| {
                    editor_error(
                        "use.office.spreadsheet_sheet_invalid",
                        "Spreadsheet sheet is missing its name.",
                    )
                })
        })
        .collect()
}

fn defined_names_container(workbook: &IndexedXmlElement) -> UseResult<Option<&IndexedXmlElement>> {
    let containers = workbook
        .children
        .iter()
        .filter(|child| child.local_name == "definedNames")
        .collect::<Vec<_>>();
    if containers.len() > 1 {
        return Err(invalid_collection(
            "contains multiple definedNames collections",
        ));
    }
    let Some(container) = containers.first().copied() else {
        return Ok(None);
    };
    if container.namespace != workbook.namespace {
        return Err(invalid_collection(
            "uses definedNames in an unexpected namespace",
        ));
    }
    Ok(Some(container))
}

fn existing_names<'a>(
    workbook: &'a IndexedXmlElement,
    container: Option<&'a IndexedXmlElement>,
    sheet_names: &[String],
) -> UseResult<Vec<ExistingName<'a>>> {
    let Some(container) = container else {
        return Ok(Vec::new());
    };
    let names = container
        .children
        .iter()
        .filter(|child| child.local_name == "definedName" && child.namespace == workbook.namespace)
        .map(|element| {
            let name = element
                .qualified_attributes
                .get("name")
                .cloned()
                .ok_or_else(|| invalid_collection("contains a definedName without name"))?;
            let local_sheet_id = element
                .qualified_attributes
                .get("localSheetId")
                .map(|value| {
                    value
                        .parse::<usize>()
                        .ok()
                        .filter(|index| *index < sheet_names.len())
                        .ok_or_else(|| {
                            invalid_collection(format!(
                                "contains defined name '{name}' with invalid localSheetId='{value}'"
                            ))
                        })
                })
                .transpose()?;
            let scope = local_sheet_id.map_or_else(
                || "workbook".to_string(),
                |index| sheet_names[index].clone(),
            );
            let scope = named_range_scope_label(&scope, local_sheet_id.is_some());
            Ok(ExistingName {
                element,
                name,
                scope,
                local_sheet_id,
            })
        })
        .collect::<UseResult<Vec<_>>>()?;
    if names.len() > MAX_SPREADSHEET_NAMED_RANGES {
        return Err(editor_error(
            "use.office.spreadsheet_named_range_limit",
            format!(
                "Spreadsheet workbook contains {} defined names; the limit is {MAX_SPREADSHEET_NAMED_RANGES}.",
                names.len()
            ),
        ));
    }
    Ok(names)
}

fn resolve_target(path: &str, existing: &[ExistingName<'_>]) -> UseResult<usize> {
    let selector = parse_named_range_path(path)?.ok_or_else(|| not_found(path))?;
    let matches = match selector {
        NamedRangePathSelector::Collection => return Err(not_found(path)),
        NamedRangePathSelector::Position(position) => {
            return existing
                .get(position - 1)
                .map(|_| position - 1)
                .ok_or_else(|| not_found(path));
        }
        NamedRangePathSelector::Name { name, scope } => existing
            .iter()
            .enumerate()
            .filter(|(_, candidate)| {
                candidate.name.eq_ignore_ascii_case(&name)
                    && scope
                        .as_deref()
                        .is_none_or(|scope| candidate.scope.eq_ignore_ascii_case(scope))
            })
            .map(|(index, _)| index)
            .collect::<Vec<_>>(),
    };
    match matches.as_slice() {
        [index] => Ok(*index),
        [] => Err(not_found(path)),
        _ => Err(editor_error(
            "use.office.spreadsheet_named_range_ambiguous",
            format!("Spreadsheet named-range path '{path}' matches multiple scopes."),
        )
        .with_suggestion(
            "Add [@scope=workbook] or [@scope=SheetName] to select one stable identity.",
        )
        .with_detail("path", path)
        .with_detail("matches", matches.len())),
    }
}

fn reject_duplicate(
    requested: &NormalizedName,
    existing: &[ExistingName<'_>],
    excluded: Option<usize>,
) -> UseResult<()> {
    if existing.iter().enumerate().any(|(index, candidate)| {
        Some(index) != excluded
            && candidate.name.eq_ignore_ascii_case(&requested.value.name)
            && candidate.local_sheet_id == requested.local_sheet_id
    }) {
        return Err(editor_error(
            "use.office.spreadsheet_named_range_duplicate",
            format!(
                "Spreadsheet defined name '{}' already exists in scope '{}'.",
                requested.value.name, requested.scope
            ),
        )
        .with_detail("name", requested.value.name.clone())
        .with_detail("scope", requested.scope.clone()));
    }
    Ok(())
}

fn validate_container_content(
    part: &LosslessXmlPart,
    workbook: &IndexedXmlElement,
    container: &IndexedXmlElement,
) -> UseResult<()> {
    let bytes = part.parse_bytes();
    let mut cursor = container.content_range.start;
    for child in &container.children {
        let gap = bytes
            .get(cursor..child.full_range.start)
            .ok_or_else(|| invalid_collection("has invalid definedNames child ranges"))?;
        if !gap.iter().all(u8::is_ascii_whitespace)
            || child.local_name != "definedName"
            || child.namespace != workbook.namespace
        {
            return Err(unknown_content(&child.qualified_name));
        }
        cursor = child.full_range.end;
    }
    let trailing = bytes
        .get(cursor..container.content_range.end)
        .ok_or_else(|| invalid_collection("has invalid definedNames trailing content"))?;
    if !trailing.iter().all(u8::is_ascii_whitespace) {
        return Err(unknown_content("non-element content"));
    }
    Ok(())
}

fn validate_editable_name_content(
    part: &LosslessXmlPart,
    element: &IndexedXmlElement,
) -> UseResult<()> {
    if !element.children.is_empty() {
        return Err(unknown_content(&element.children[0].qualified_name));
    }
    decoded_element_text(part, element)
        .map(|_| ())
        .map_err(|_| unknown_content("CDATA, comment, or processing-instruction content"))
}

fn name_fragment(
    namespace_prefix: Option<&str>,
    existing: Option<&IndexedXmlElement>,
    named_range: &NormalizedName,
) -> String {
    let mut attributes = existing
        .map(|element| element.qualified_attributes.clone())
        .unwrap_or_default();
    for known in ["name", "localSheetId", "comment", "function"] {
        attributes.remove(known);
    }
    attributes.insert("name".into(), named_range.value.name.clone());
    if let Some(local_sheet_id) = named_range.local_sheet_id {
        attributes.insert("localSheetId".into(), local_sheet_id.to_string());
    }
    if let Some(comment) = &named_range.value.comment {
        attributes.insert("comment".into(), comment.clone());
    }
    if named_range.value.volatile {
        attributes.insert("function".into(), "1".into());
    }
    let attributes = attributes
        .into_iter()
        .map(|(name, value)| format!(" {name}=\"{}\"", escape_attribute(&value)))
        .collect::<String>();
    let tag = existing.map_or_else(
        || qualified(namespace_prefix, "definedName"),
        |element| element.qualified_name.clone(),
    );
    format!(
        "<{tag}{attributes}>{}</{tag}>",
        escape_text(&named_range.value.reference)
    )
}

fn reject_table_collision(package: &NativeOfficePackage, name: &str) -> UseResult<()> {
    let table_parts = package
        .part_names()
        .filter(|part| part.starts_with("xl/tables/") && part.ends_with(".xml"))
        .map(str::to_string)
        .collect::<Vec<_>>();
    for part_name in table_parts {
        let part = package.xml_part(&part_name)?;
        let table = index_xml(&part)?;
        if ["name", "displayName"].iter().any(|attribute| {
            table
                .qualified_attributes
                .get(*attribute)
                .is_some_and(|candidate| candidate.eq_ignore_ascii_case(name))
        }) {
            return Err(editor_error(
                "use.office.spreadsheet_named_range_name_collision",
                format!("Spreadsheet defined name '{name}' collides with an existing table name."),
            )
            .with_detail("name", name)
            .with_detail("part", part_name));
        }
    }
    Ok(())
}

fn is_unqualified_a1_range(reference: &str) -> bool {
    !reference.contains('!') && CellRange::parse(&reference.replace('$', "")).is_ok()
}

fn validate_simple_qualified_sheet(reference: &str, sheet_names: &[String]) -> UseResult<()> {
    let Some((qualifier, area)) = reference.split_once('!') else {
        return Ok(());
    };
    if !is_unqualified_a1_range(area) {
        return Ok(());
    }
    let sheet = if qualifier.starts_with('\'') && qualifier.ends_with('\'') && qualifier.len() >= 2
    {
        qualifier[1..qualifier.len() - 1].replace("''", "'")
    } else {
        qualifier.to_string()
    };
    if sheet_names
        .iter()
        .any(|candidate| candidate.eq_ignore_ascii_case(&sheet))
    {
        Ok(())
    } else {
        Err(editor_error(
            "use.office.spreadsheet_named_range_ref_invalid",
            format!(
                "Spreadsheet named-range ref '{reference}' targets missing worksheet '{sheet}'."
            ),
        )
        .with_detail("ref", reference)
        .with_detail("sheet", sheet))
    }
}

fn unknown_content(child: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_named_range_unknown_content",
        "The named-range mutation cannot proceed without risking unknown definedNames content.",
    )
    .with_detail("part", "xl/workbook.xml")
    .with_detail("child", child)
}

fn invalid_collection(reason: impl Into<String>) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_named_range_invalid",
        format!("Spreadsheet workbook.xml {}.", reason.into()),
    )
    .with_detail("part", "xl/workbook.xml")
}

fn not_found(path: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.node_not_found",
        format!("Office semantic path '{path}' does not exist."),
    )
    .with_detail("path", path)
}
