use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::UseResult;

use super::{editor_error, prefix, qualified, validate_mutation_path};
use crate::editor::NativeSpreadsheetConditionalFormat;
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::CellRange;
use crate::xml_edit::{
    apply_patches, index_xml, insert_ordered_child, patch_start_tag_attributes, IndexedXmlElement,
    XmlPatch,
};
use crate::{DocumentKind, LosslessXmlPart, NativeOfficePackage};

mod normalize;
mod xml;

use normalize::{normalize, MAX_CONDITIONAL_FORMATS, MAX_CONDITIONAL_FORMAT_RANGES};
use xml::{reject_unknown_rule_children, rule_fragment};

const WORKSHEET_CHILDREN_AFTER_CONDITIONAL_FORMATTING: &[&str] = &[
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
];

struct ResolvedSheet {
    path: String,
    part: String,
}

struct ExistingRule<'a> {
    container: &'a IndexedXmlElement,
    element: &'a IndexedXmlElement,
    ranges: Vec<String>,
    priority: u32,
    rules_in_container: usize,
}

pub(super) fn is_path(path: &str) -> bool {
    path.rsplit_once('/')
        .is_some_and(|(_, segment)| parse_rule_segment(segment).is_some())
}

pub(super) fn add(
    package: &mut NativeOfficePackage,
    sheet: &str,
    value: &NativeSpreadsheetConditionalFormat,
) -> UseResult<String> {
    let resolved = resolve_sheet(package, sheet)?;
    let value = normalize(value)?;
    let part = package.xml_part(&resolved.part)?;
    let index = index_xml(&part)?;
    let rules = existing_rules(&index, &resolved.part)?;
    if rules.len() >= MAX_CONDITIONAL_FORMATS {
        return Err(editor_error(
            "use.office.spreadsheet_conditional_format_limit",
            format!(
                "Worksheet '{}' already has the maximum {MAX_CONDITIONAL_FORMATS} conditional formats.",
                resolved.path
            ),
        ));
    }
    let priority = next_priority(&rules)?;
    let dxf_id = value
        .rule
        .differential_format()
        .map(|format| super::style::find_or_append_differential_format(package, format))
        .transpose()?
        .flatten();
    let rule = rule_fragment(
        prefix(&index.qualified_name),
        None,
        &value,
        priority,
        dxf_id,
    )?;
    let container = qualified(prefix(&index.qualified_name), "conditionalFormatting");
    let fragment = format!(
        "<{container} sqref=\"{}\">{rule}</{container}>",
        crate::xml_edit::escape_attribute(&value.ranges.join(" "))
    );
    let edited = insert_ordered_child(
        &part,
        &index,
        fragment,
        WORKSHEET_CHILDREN_AFTER_CONDITIONAL_FORMATTING,
    )?;
    package.set_part(&resolved.part, edited)?;
    Ok(format!("{}/cf[{}]", resolved.path, rules.len() + 1))
}

pub(super) fn set(
    package: &mut NativeOfficePackage,
    path: &str,
    value: &NativeSpreadsheetConditionalFormat,
) -> UseResult<String> {
    let (resolved, requested_index) = resolve_rule(package, path)?;
    let value = normalize(value)?;
    let part = package.xml_part(&resolved.part)?;
    let index = index_xml(&part)?;
    let rules = existing_rules(&index, &resolved.part)?;
    let existing = rules
        .get(requested_index - 1)
        .ok_or_else(|| rule_not_found(&resolved.path, requested_index))?;
    reject_unknown_rule_children(&resolved.part, &part, existing.element)?;
    let ranges_changed = existing.ranges != value.ranges;
    if ranges_changed && existing.rules_in_container != 1 {
        return Err(editor_error(
            "use.office.spreadsheet_conditional_format_shared_range",
            "The selected conditional format shares its sqref container with another rule, so its ranges cannot be changed independently.",
        )
        .with_suggestion("Keep the existing ranges or recreate the shared rules as separate conditional formats.")
        .with_detail("path", path));
    }
    if ranges_changed && !container_has_only_rule(&part, existing.container, existing.element) {
        return Err(editor_error(
            "use.office.spreadsheet_conditional_format_unknown_content",
            "The conditional-format range cannot be changed without affecting unknown shared container content.",
        )
        .with_detail("part", resolved.part.clone())
        .with_detail("path", path));
    }
    let dxf_id = value
        .rule
        .differential_format()
        .map(|format| super::style::find_or_append_differential_format(package, format))
        .transpose()?
        .flatten();
    let replacement = rule_fragment(
        prefix(&existing.element.qualified_name),
        Some(existing.element),
        &value,
        existing.priority,
        dxf_id,
    )?;
    let mut edited = apply_patches(
        &part,
        vec![XmlPatch::new(
            existing.element.full_range.clone(),
            replacement,
        )],
    )?;
    if ranges_changed {
        let updated = LosslessXmlPart::parse(resolved.part.clone(), edited)?;
        let updated_index = index_xml(&updated)?;
        let updated_rules = existing_rules(&updated_index, &resolved.part)?;
        let updated_rule = updated_rules
            .get(requested_index - 1)
            .ok_or_else(|| rule_not_found(&resolved.path, requested_index))?;
        edited = patch_start_tag_attributes(
            &updated,
            updated_rule.container,
            &BTreeMap::from([("sqref".to_string(), Some(value.ranges.join(" ")))]),
        )?;
    }
    package.set_part(&resolved.part, edited)?;
    Ok(format!("{}/cf[{requested_index}]", resolved.path))
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let (resolved, requested_index) = resolve_rule(package, path)?;
    let part = package.xml_part(&resolved.part)?;
    let index = index_xml(&part)?;
    let rules = existing_rules(&index, &resolved.part)?;
    let existing = rules
        .get(requested_index - 1)
        .ok_or_else(|| rule_not_found(&resolved.path, requested_index))?;
    reject_unknown_rule_children(&resolved.part, &part, existing.element)?;
    let range = if existing.rules_in_container == 1 {
        if !container_is_removable(&part, existing.container, existing.element) {
            return Err(editor_error(
                "use.office.spreadsheet_conditional_format_unknown_content",
                "The final rule cannot be removed without discarding unknown conditionalFormatting container content.",
            )
            .with_detail("part", resolved.part)
            .with_detail("path", path));
        }
        existing.container.full_range.clone()
    } else {
        existing.element.full_range.clone()
    };
    let edited = apply_patches(&part, vec![XmlPatch::new(range, Vec::new())])?;
    package.set_part(&resolved.part, edited)
}

fn resolve_sheet(package: &NativeOfficePackage, requested: &str) -> UseResult<ResolvedSheet> {
    if package.kind() != DocumentKind::Spreadsheet {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Conditional-formatting operations are available only for Spreadsheet documents.",
        ));
    }
    validate_mutation_path(requested)?;
    if requested.trim_start_matches('/').contains('/') {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Adding a conditional format requires a worksheet path such as /Sheet1.",
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
        .ok_or_else(|| super::node_not_found(requested))?;
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

fn resolve_rule(
    package: &NativeOfficePackage,
    requested: &str,
) -> UseResult<(ResolvedSheet, usize)> {
    validate_mutation_path(requested)?;
    let (sheet, segment) = requested.rsplit_once('/').ok_or_else(|| {
        editor_error(
            "use.office.mutation_path_unsupported",
            "Conditional-format updates require a path such as /Sheet1/cf[1].",
        )
    })?;
    let index = parse_rule_segment(segment).ok_or_else(|| {
        editor_error(
            "use.office.mutation_path_unsupported",
            "Conditional-format updates require a path such as /Sheet1/cf[1].",
        )
    })?;
    Ok((resolve_sheet(package, sheet)?, index))
}

fn parse_rule_segment(segment: &str) -> Option<usize> {
    let (name, index) = segment.split_once('[')?;
    if !matches!(
        name.to_ascii_lowercase().replace(['-', '_'], "").as_str(),
        "cf" | "conditionalformat" | "conditionalformatting"
    ) {
        return None;
    }
    index
        .strip_suffix(']')?
        .parse::<usize>()
        .ok()
        .filter(|index| *index > 0)
}

fn existing_rules<'a>(
    worksheet: &'a IndexedXmlElement,
    part_name: &str,
) -> UseResult<Vec<ExistingRule<'a>>> {
    let mut result = Vec::new();
    let mut priorities = BTreeSet::new();
    for container in worksheet.children.iter().filter(|child| {
        child.local_name == "conditionalFormatting" && child.namespace == worksheet.namespace
    }) {
        let ranges = parse_sqref(
            container.attributes.get("sqref").ok_or_else(|| {
                invalid(part_name, "contains conditionalFormatting without sqref")
            })?,
            part_name,
        )?;
        let rules = container
            .children
            .iter()
            .filter(|child| child.local_name == "cfRule" && child.namespace == worksheet.namespace)
            .collect::<Vec<_>>();
        for element in &rules {
            let priority = element
                .attributes
                .get("priority")
                .and_then(|value| value.parse::<u32>().ok())
                .filter(|value| *value > 0)
                .ok_or_else(|| invalid(part_name, "contains cfRule with invalid priority"))?;
            if !priorities.insert(priority) {
                return Err(invalid(
                    part_name,
                    format!("contains duplicate cfRule priority {priority}"),
                ));
            }
            result.push(ExistingRule {
                container,
                element,
                ranges: ranges.clone(),
                priority,
                rules_in_container: rules.len(),
            });
        }
    }
    if result.len() > MAX_CONDITIONAL_FORMATS {
        return Err(editor_error(
            "use.office.spreadsheet_conditional_format_limit",
            format!(
                "Worksheet part '{part_name}' contains {} conditional formats; the limit is {MAX_CONDITIONAL_FORMATS}.",
                result.len()
            ),
        ));
    }
    Ok(result)
}

fn parse_sqref(reference: &str, part_name: &str) -> UseResult<Vec<String>> {
    let tokens = reference.split_ascii_whitespace().collect::<Vec<_>>();
    if tokens.is_empty() || tokens.len() > MAX_CONDITIONAL_FORMAT_RANGES {
        return Err(invalid(
            part_name,
            format!("contains invalid conditional-format sqref '{reference}'"),
        ));
    }
    tokens
        .into_iter()
        .map(|range| {
            CellRange::parse(range)
                .map(|range| range.a1())
                .map_err(|error| {
                    invalid(
                        part_name,
                        format!("contains invalid conditional-format range '{range}': {error}"),
                    )
                })
        })
        .collect()
}

fn next_priority(rules: &[ExistingRule<'_>]) -> UseResult<u32> {
    rules
        .iter()
        .map(|rule| rule.priority)
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| {
            editor_error(
                "use.office.spreadsheet_conditional_format_priority_exhausted",
                "Spreadsheet conditional-format priorities are exhausted.",
            )
        })
}

fn container_has_only_rule(
    part: &LosslessXmlPart,
    container: &IndexedXmlElement,
    rule: &IndexedXmlElement,
) -> bool {
    container.children.len() == 1
        && container.children[0].full_range == rule.full_range
        && content_around_rule_is_whitespace(part, container, rule)
}

fn container_is_removable(
    part: &LosslessXmlPart,
    container: &IndexedXmlElement,
    rule: &IndexedXmlElement,
) -> bool {
    container
        .qualified_attributes
        .keys()
        .all(|name| name == "sqref")
        && container_has_only_rule(part, container, rule)
}

fn content_around_rule_is_whitespace(
    part: &LosslessXmlPart,
    container: &IndexedXmlElement,
    rule: &IndexedXmlElement,
) -> bool {
    let bytes = part.parse_bytes();
    bytes
        .get(container.content_range.start..rule.full_range.start)
        .into_iter()
        .flatten()
        .chain(
            bytes
                .get(rule.full_range.end..container.content_range.end)
                .into_iter()
                .flatten(),
        )
        .all(u8::is_ascii_whitespace)
}

fn rule_not_found(sheet: &str, index: usize) -> a3s_use_core::UseError {
    let path = format!("{sheet}/cf[{index}]");
    editor_error(
        "use.office.node_not_found",
        format!("Office semantic path '{path}' does not exist."),
    )
    .with_detail("path", path)
}

fn invalid(part_name: &str, reason: impl Into<String>) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_conditional_format_invalid",
        format!(
            "Spreadsheet worksheet conditional formatting {}.",
            reason.into()
        ),
    )
    .with_detail("part", part_name)
}
