use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{editor_error, prefix, qualified, validate_mutation_path};
use crate::editor::{
    NativeSpreadsheetDataValidation, NativeSpreadsheetDataValidationErrorStyle,
    NativeSpreadsheetDataValidationOperator, NativeSpreadsheetDataValidationType,
};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::{
    first_intersecting_ranges, CellRange, MAX_DATA_VALIDATIONS, MAX_DATA_VALIDATION_RANGES,
};
use crate::xml_edit::{
    apply_patches, escape_attribute, escape_text, index_xml, insert_ordered_child,
    patch_start_tag_attributes, IndexedXmlElement, XmlPatch,
};
use crate::{DocumentKind, LosslessXmlPart, NativeOfficePackage};

mod normalize;

use normalize::{
    normalize_comparison_formula, normalize_list_formula, reject_operator, strip_formula_equals,
    validate_optional_text, validate_xml_text, workbook_uses_1904_date_system,
};

const WORKSHEET_CHILDREN_AFTER_VALIDATIONS: &[&str] = &[
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

const KNOWN_RULE_ATTRIBUTES: &[&str] = &[
    "type",
    "errorStyle",
    "operator",
    "allowBlank",
    "showDropDown",
    "showInputMessage",
    "showErrorMessage",
    "errorTitle",
    "error",
    "promptTitle",
    "prompt",
    "sqref",
];

struct ResolvedSheet {
    path: String,
    part: String,
}

struct ExistingRule<'a> {
    element: &'a IndexedXmlElement,
    ranges: Vec<CellRange>,
}

pub(super) fn is_path(path: &str) -> bool {
    path.rsplit_once('/')
        .is_some_and(|(_, segment)| parse_rule_segment(segment).is_some())
}

pub(super) fn add(
    package: &mut NativeOfficePackage,
    sheet: &str,
    validation: &NativeSpreadsheetDataValidation,
) -> UseResult<String> {
    let resolved = resolve_sheet(package, sheet)?;
    let validation = normalize_validation(package, validation)?;
    let part = package.xml_part(&resolved.part)?;
    let index = index_xml(&part)?;
    let container = validation_container(&index, &resolved.part)?;
    let rules = existing_rules(&index, container, &resolved.part)?;
    if rules.len() >= MAX_DATA_VALIDATIONS {
        return Err(editor_error(
            "use.office.spreadsheet_validation_limit",
            format!(
                "Worksheet '{}' already has the maximum {MAX_DATA_VALIDATIONS} data validations.",
                resolved.path
            ),
        ));
    }
    reject_overlap(&validation.ranges, &rules, None)?;

    let rule = rule_fragment(prefix(&index.qualified_name), None, &validation)?;
    let next_index = rules.len() + 1;
    let edited = if let Some(container) = container {
        validate_container_for_add(&resolved.part, container, index.namespace.as_deref())?;
        let inserted = insert_ordered_child(&part, container, rule, &["extLst"])?;
        update_count(&resolved.part, inserted, next_index)?
    } else {
        let name = qualified(prefix(&index.qualified_name), "dataValidations");
        let fragment = format!("<{name} count=\"1\">{rule}</{name}>");
        insert_ordered_child(
            &part,
            &index,
            fragment,
            WORKSHEET_CHILDREN_AFTER_VALIDATIONS,
        )?
    };
    package.set_part(&resolved.part, edited)?;
    Ok(format!("{}/dataValidation[{next_index}]", resolved.path))
}

pub(super) fn set(
    package: &mut NativeOfficePackage,
    path: &str,
    validation: &NativeSpreadsheetDataValidation,
) -> UseResult<String> {
    let (resolved, requested_index) = resolve_rule(package, path)?;
    let validation = normalize_validation(package, validation)?;
    let part = package.xml_part(&resolved.part)?;
    let index = index_xml(&part)?;
    let container = validation_container(&index, &resolved.part)?
        .ok_or_else(|| rule_not_found(&resolved.path, requested_index))?;
    let rules = existing_rules(&index, Some(container), &resolved.part)?;
    let existing = rules
        .get(requested_index - 1)
        .ok_or_else(|| rule_not_found(&resolved.path, requested_index))?;
    reject_overlap(&validation.ranges, &rules, Some(requested_index - 1))?;
    reject_unknown_rule_children(&resolved.part, &part, existing.element)?;

    let replacement = rule_fragment(
        prefix(&existing.element.qualified_name),
        Some(existing.element),
        &validation,
    )?;
    let edited = apply_patches(
        &part,
        vec![XmlPatch::new(
            existing.element.full_range.clone(),
            replacement,
        )],
    )?;
    package.set_part(&resolved.part, edited)?;
    Ok(format!(
        "{}/dataValidation[{requested_index}]",
        resolved.path
    ))
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let (resolved, requested_index) = resolve_rule(package, path)?;
    let part = package.xml_part(&resolved.part)?;
    let index = index_xml(&part)?;
    let container = validation_container(&index, &resolved.part)?
        .ok_or_else(|| rule_not_found(&resolved.path, requested_index))?;
    let rules = existing_rules(&index, Some(container), &resolved.part)?;
    let existing = rules
        .get(requested_index - 1)
        .ok_or_else(|| rule_not_found(&resolved.path, requested_index))?;
    let remaining = rules.len().saturating_sub(1);
    let edited = if remaining == 0 {
        if !container_is_removable(&part, container, existing.element) {
            return Err(editor_error(
                "use.office.spreadsheet_validation_unknown_content",
                "The final data validation cannot be removed without discarding unknown dataValidations collection content.",
            )
            .with_suggestion(
                "Inspect the worksheet with native raw XML and preserve or relocate the unknown data first.",
            )
            .with_detail("part", resolved.part)
            .with_detail("path", path));
        }
        apply_patches(
            &part,
            vec![XmlPatch::new(container.full_range.clone(), Vec::new())],
        )?
    } else {
        let removed = apply_patches(
            &part,
            vec![XmlPatch::new(
                existing.element.full_range.clone(),
                Vec::new(),
            )],
        )?;
        update_count(&resolved.part, removed, remaining)?
    };
    package.set_part(&resolved.part, edited)
}

fn resolve_sheet(package: &NativeOfficePackage, requested: &str) -> UseResult<ResolvedSheet> {
    if package.kind() != DocumentKind::Spreadsheet {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Data-validation operations are available only for Spreadsheet documents.",
        ));
    }
    validate_mutation_path(requested)?;
    if requested.trim_start_matches('/').contains('/') {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Adding a data validation requires a worksheet path such as /Sheet1.",
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
        .ok_or_else(|| {
            editor_error(
                "use.office.node_not_found",
                format!("Office semantic path '{requested}' does not exist."),
            )
            .with_detail("path", requested)
        })?;
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
            "Data-validation updates require a path such as /Sheet1/dataValidation[1].",
        )
    })?;
    let index = parse_rule_segment(segment).ok_or_else(|| {
        editor_error(
            "use.office.mutation_path_unsupported",
            "Data-validation updates require a path such as /Sheet1/dataValidation[1].",
        )
    })?;
    Ok((resolve_sheet(package, sheet)?, index))
}

fn parse_rule_segment(segment: &str) -> Option<usize> {
    let (name, index) = segment.split_once('[')?;
    if !matches!(
        name.to_ascii_lowercase().as_str(),
        "datavalidation" | "validation"
    ) {
        return None;
    }
    index
        .strip_suffix(']')?
        .parse::<usize>()
        .ok()
        .filter(|index| *index > 0)
}

fn normalize_validation(
    package: &NativeOfficePackage,
    validation: &NativeSpreadsheetDataValidation,
) -> UseResult<NativeSpreadsheetDataValidation> {
    let mut normalized = validation.clone();
    if normalized.ranges.is_empty() || normalized.ranges.len() > MAX_DATA_VALIDATION_RANGES {
        return Err(editor_error(
            "use.office.spreadsheet_validation_range_invalid",
            format!(
                "A data validation requires 1-{MAX_DATA_VALIDATION_RANGES} rectangular A1 ranges."
            ),
        )
        .with_detail("ranges", normalized.ranges.len()));
    }
    let ranges = normalized
        .ranges
        .iter()
        .map(|range| CellRange::parse(range))
        .collect::<UseResult<Vec<_>>>()?;
    if let Some((left, right)) = first_intersecting_ranges(&ranges) {
        return Err(overlap_error(ranges[right], ranges[left]));
    }
    normalized.ranges = ranges.iter().map(|range| range.a1()).collect();

    let formula1 = normalized.formula1.clone();
    if formula1.is_empty() {
        return Err(editor_error(
            "use.office.spreadsheet_validation_formula_required",
            "Data-validation formula1 must contain 1-255 characters.",
        ));
    }
    validate_xml_text(&formula1, "formula1", 255)?;
    if let Some(formula2) = normalized.formula2.as_deref() {
        validate_xml_text(formula2, "formula2", 255)?;
    }

    let date_1904 = if normalized.validation_type == NativeSpreadsheetDataValidationType::Date {
        workbook_uses_1904_date_system(package)?
    } else {
        false
    };
    match normalized.validation_type {
        NativeSpreadsheetDataValidationType::List => {
            reject_operator(&normalized)?;
            if normalized.formula2.is_some() {
                return Err(editor_error(
                    "use.office.spreadsheet_validation_formula2_unsupported",
                    "List data validation does not accept formula2.",
                ));
            }
            normalized.formula1 = normalize_list_formula(&formula1)?;
        }
        NativeSpreadsheetDataValidationType::Custom => {
            reject_operator(&normalized)?;
            if normalized.formula2.is_some() {
                return Err(editor_error(
                    "use.office.spreadsheet_validation_formula2_unsupported",
                    "Custom data validation does not accept formula2.",
                ));
            }
            normalized.formula1 = strip_formula_equals(&formula1)?;
        }
        NativeSpreadsheetDataValidationType::Whole
        | NativeSpreadsheetDataValidationType::Decimal
        | NativeSpreadsheetDataValidationType::Date
        | NativeSpreadsheetDataValidationType::Time
        | NativeSpreadsheetDataValidationType::TextLength => {
            let operator = normalized.operator.ok_or_else(|| {
                editor_error(
                    "use.office.spreadsheet_validation_operator_required",
                    "Whole, decimal, date, time, and text-length validation requires an operator.",
                )
            })?;
            let requires_second = matches!(
                operator,
                NativeSpreadsheetDataValidationOperator::Between
                    | NativeSpreadsheetDataValidationOperator::NotBetween
            );
            if requires_second && normalized.formula2.is_none() {
                return Err(editor_error(
                    "use.office.spreadsheet_validation_formula2_required",
                    "Between and not-between data validation requires formula2.",
                ));
            }
            if !requires_second && normalized.formula2.is_some() {
                return Err(editor_error(
                    "use.office.spreadsheet_validation_formula2_unsupported",
                    "Only between and not-between data validation accepts formula2.",
                ));
            }
            normalized.formula1 =
                normalize_comparison_formula(normalized.validation_type, &formula1, date_1904)?;
            normalized.formula2 = normalized
                .formula2
                .as_deref()
                .map(|formula| {
                    normalize_comparison_formula(normalized.validation_type, formula, date_1904)
                })
                .transpose()?;
        }
    }

    if normalized.validation_type != NativeSpreadsheetDataValidationType::List
        && !normalized.in_cell_dropdown
    {
        return Err(editor_error(
            "use.office.spreadsheet_validation_dropdown_unsupported",
            "inCellDropdown=false is supported only for list data validation.",
        ));
    }
    validate_optional_text(normalized.prompt_title.as_deref(), "promptTitle", 32)?;
    validate_optional_text(normalized.prompt.as_deref(), "prompt", 255)?;
    validate_optional_text(normalized.error_title.as_deref(), "errorTitle", 32)?;
    validate_optional_text(normalized.error.as_deref(), "error", 225)?;
    validate_xml_text(&normalized.formula1, "formula1", 255)?;
    if let Some(formula2) = normalized.formula2.as_deref() {
        validate_xml_text(formula2, "formula2", 255)?;
    }
    Ok(normalized)
}

fn validation_container<'a>(
    worksheet: &'a IndexedXmlElement,
    part_name: &str,
) -> UseResult<Option<&'a IndexedXmlElement>> {
    let candidates = worksheet
        .children
        .iter()
        .filter(|child| child.local_name == "dataValidations")
        .collect::<Vec<_>>();
    if candidates.len() > 1 {
        return Err(invalid_collection(
            part_name,
            "contains multiple dataValidations collections",
        ));
    }
    let Some(container) = candidates.first().copied() else {
        return Ok(None);
    };
    if container.namespace != worksheet.namespace {
        return Err(invalid_collection(
            part_name,
            "uses dataValidations in an unexpected namespace",
        ));
    }
    Ok(Some(container))
}

fn validate_container_for_add(
    part_name: &str,
    container: &IndexedXmlElement,
    worksheet_namespace: Option<&str>,
) -> UseResult<()> {
    let mut saw_extension_list = false;
    for child in &container.children {
        let is_worksheet_child = child.namespace.as_deref() == worksheet_namespace;
        if is_worksheet_child && child.local_name == "dataValidation" && !saw_extension_list {
            continue;
        }
        if is_worksheet_child && child.local_name == "extLst" && !saw_extension_list {
            saw_extension_list = true;
            continue;
        }
        return Err(editor_error(
            "use.office.spreadsheet_validation_unknown_content",
            "A data validation cannot be added without risking invalid dataValidations child order.",
        )
        .with_detail("part", part_name)
        .with_detail("child", child.qualified_name.clone()));
    }
    Ok(())
}

fn existing_rules<'a>(
    worksheet: &'a IndexedXmlElement,
    container: Option<&'a IndexedXmlElement>,
    part_name: &str,
) -> UseResult<Vec<ExistingRule<'a>>> {
    let Some(container) = container else {
        return Ok(Vec::new());
    };
    let rules = container
        .children
        .iter()
        .filter(|child| {
            child.local_name == "dataValidation" && child.namespace == worksheet.namespace
        })
        .map(|element| {
            let reference = element.attributes.get("sqref").ok_or_else(|| {
                invalid_collection(part_name, "contains a dataValidation without sqref")
            })?;
            let ranges = parse_sqref(reference, part_name)?;
            Ok(ExistingRule { element, ranges })
        })
        .collect::<UseResult<Vec<_>>>()?;
    if rules.len() > MAX_DATA_VALIDATIONS {
        return Err(editor_error(
            "use.office.spreadsheet_validation_limit",
            format!(
                "Worksheet part '{part_name}' contains {} data validations; the limit is {MAX_DATA_VALIDATIONS}.",
                rules.len()
            ),
        ));
    }
    Ok(rules)
}

fn parse_sqref(reference: &str, part_name: &str) -> UseResult<Vec<CellRange>> {
    let tokens = reference.split_ascii_whitespace().collect::<Vec<_>>();
    if tokens.is_empty() || tokens.len() > MAX_DATA_VALIDATION_RANGES {
        return Err(invalid_collection(
            part_name,
            format!("contains invalid validation sqref '{reference}'"),
        ));
    }
    tokens
        .into_iter()
        .map(|range| {
            CellRange::parse(range).map_err(|error| {
                invalid_collection(
                    part_name,
                    format!("contains invalid validation range '{range}': {error}"),
                )
            })
        })
        .collect()
}

fn reject_overlap(
    requested: &[String],
    existing: &[ExistingRule<'_>],
    excluded_rule: Option<usize>,
) -> UseResult<()> {
    let requested = requested
        .iter()
        .map(|range| CellRange::parse(range))
        .collect::<UseResult<Vec<_>>>()?;
    for (rule_index, rule) in existing.iter().enumerate() {
        if excluded_rule == Some(rule_index) {
            continue;
        }
        for requested in &requested {
            if let Some(existing) = rule
                .ranges
                .iter()
                .find(|range| range.intersects(*requested))
            {
                return Err(overlap_error(*requested, *existing));
            }
        }
    }
    Ok(())
}

fn overlap_error(requested: CellRange, existing: CellRange) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_validation_overlap",
        format!(
            "Data-validation range '{}' overlaps existing validation range '{}'.",
            requested.a1(),
            existing.a1()
        ),
    )
    .with_suggestion("Remove or move the existing rule, or choose disjoint ranges.")
    .with_detail("requested", requested.a1())
    .with_detail("existing", existing.a1())
}

fn rule_fragment(
    namespace_prefix: Option<&str>,
    existing: Option<&IndexedXmlElement>,
    validation: &NativeSpreadsheetDataValidation,
) -> UseResult<String> {
    let mut attributes = existing
        .map(|element| element.qualified_attributes.clone())
        .unwrap_or_default();
    for attribute in KNOWN_RULE_ATTRIBUTES {
        attributes.remove(*attribute);
    }
    attributes.insert(
        "type".into(),
        validation_type_token(validation.validation_type).into(),
    );
    attributes.insert("sqref".into(), validation.ranges.join(" "));
    attributes.insert(
        "allowBlank".into(),
        bool_token(validation.allow_blank).into(),
    );
    attributes.insert(
        "showInputMessage".into(),
        bool_token(validation.show_input).into(),
    );
    attributes.insert(
        "showErrorMessage".into(),
        bool_token(validation.show_error).into(),
    );
    if !validation.in_cell_dropdown {
        attributes.insert("showDropDown".into(), "1".into());
    }
    if let Some(operator) = validation.operator {
        attributes.insert("operator".into(), operator_token(operator).into());
    }
    if validation.error_style != NativeSpreadsheetDataValidationErrorStyle::Stop {
        attributes.insert(
            "errorStyle".into(),
            error_style_token(validation.error_style).into(),
        );
    }
    for (name, value) in [
        ("promptTitle", validation.prompt_title.as_deref()),
        ("prompt", validation.prompt.as_deref()),
        ("errorTitle", validation.error_title.as_deref()),
        ("error", validation.error.as_deref()),
    ] {
        if let Some(value) = value {
            attributes.insert(name.into(), value.into());
        }
    }
    let attributes = attributes
        .into_iter()
        .map(|(name, value)| format!(" {name}=\"{}\"", escape_attribute(&value)))
        .collect::<String>();
    let rule_name = existing.map_or_else(
        || qualified(namespace_prefix, "dataValidation"),
        |element| element.qualified_name.clone(),
    );
    let formula_prefix = prefix(&rule_name).or(namespace_prefix);
    let formula1_name = qualified(formula_prefix, "formula1");
    let mut children = format!(
        "<{formula1_name}>{}</{formula1_name}>",
        escape_text(&validation.formula1)
    );
    if let Some(formula2) = validation.formula2.as_deref() {
        let formula2_name = qualified(formula_prefix, "formula2");
        children.push_str(&format!(
            "<{formula2_name}>{}</{formula2_name}>",
            escape_text(formula2)
        ));
    }
    Ok(format!("<{rule_name}{attributes}>{children}</{rule_name}>"))
}

fn reject_unknown_rule_children(
    part_name: &str,
    part: &LosslessXmlPart,
    rule: &IndexedXmlElement,
) -> UseResult<()> {
    let mut counts = BTreeMap::<&str, usize>::new();
    let bytes = part.parse_bytes();
    let mut cursor = rule.content_range.start;
    for child in &rule.children {
        let gap = bytes
            .get(cursor..child.full_range.start)
            .ok_or_else(|| invalid_collection(part_name, "has invalid rule child ranges"))?;
        if !gap.iter().all(u8::is_ascii_whitespace) {
            return Err(unknown_rule_content(part_name, "non-element child content"));
        }
        if child.namespace != rule.namespace
            || !matches!(child.local_name.as_str(), "formula1" | "formula2")
        {
            return Err(unknown_rule_content(part_name, &child.qualified_name));
        }
        let content = bytes
            .get(child.content_range.clone())
            .ok_or_else(|| invalid_collection(part_name, "has invalid formula child ranges"))?;
        if !child.qualified_attributes.is_empty()
            || !child.children.is_empty()
            || content.contains(&b'<')
        {
            return Err(unknown_rule_content(part_name, &child.qualified_name));
        }
        *counts.entry(child.local_name.as_str()).or_default() += 1;
        cursor = child.full_range.end;
    }
    let trailing = bytes
        .get(cursor..rule.content_range.end)
        .ok_or_else(|| invalid_collection(part_name, "has invalid trailing rule content"))?;
    if !trailing.iter().all(u8::is_ascii_whitespace) {
        return Err(unknown_rule_content(part_name, "non-element child content"));
    }
    if counts.values().any(|count| *count > 1) {
        return Err(invalid_collection(
            part_name,
            "contains duplicate data-validation formula children",
        ));
    }
    Ok(())
}

fn unknown_rule_content(part_name: &str, child: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_validation_unknown_content",
        "The data validation cannot be replaced without discarding unknown child content.",
    )
    .with_detail("part", part_name)
    .with_detail("child", child)
}

fn update_count(part_name: &str, bytes: Vec<u8>, count: usize) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(part_name.to_string(), bytes)?;
    let index = index_xml(&part)?;
    let container = validation_container(&index, part_name)?.ok_or_else(|| {
        invalid_collection(
            part_name,
            "lost its dataValidations collection during mutation",
        )
    })?;
    patch_start_tag_attributes(
        &part,
        container,
        &BTreeMap::from([("count".to_string(), Some(count.to_string()))]),
    )
}

fn container_is_removable(
    part: &LosslessXmlPart,
    container: &IndexedXmlElement,
    rule: &IndexedXmlElement,
) -> bool {
    if container
        .qualified_attributes
        .keys()
        .any(|name| name != "count")
        || container.children.len() != 1
        || container.children[0].full_range != rule.full_range
    {
        return false;
    }
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

const fn bool_token(value: bool) -> &'static str {
    if value {
        "1"
    } else {
        "0"
    }
}

const fn validation_type_token(value: NativeSpreadsheetDataValidationType) -> &'static str {
    match value {
        NativeSpreadsheetDataValidationType::List => "list",
        NativeSpreadsheetDataValidationType::Whole => "whole",
        NativeSpreadsheetDataValidationType::Decimal => "decimal",
        NativeSpreadsheetDataValidationType::Date => "date",
        NativeSpreadsheetDataValidationType::Time => "time",
        NativeSpreadsheetDataValidationType::TextLength => "textLength",
        NativeSpreadsheetDataValidationType::Custom => "custom",
    }
}

const fn operator_token(value: NativeSpreadsheetDataValidationOperator) -> &'static str {
    match value {
        NativeSpreadsheetDataValidationOperator::Between => "between",
        NativeSpreadsheetDataValidationOperator::NotBetween => "notBetween",
        NativeSpreadsheetDataValidationOperator::Equal => "equal",
        NativeSpreadsheetDataValidationOperator::NotEqual => "notEqual",
        NativeSpreadsheetDataValidationOperator::GreaterThan => "greaterThan",
        NativeSpreadsheetDataValidationOperator::GreaterThanOrEqual => "greaterThanOrEqual",
        NativeSpreadsheetDataValidationOperator::LessThan => "lessThan",
        NativeSpreadsheetDataValidationOperator::LessThanOrEqual => "lessThanOrEqual",
    }
}

const fn error_style_token(value: NativeSpreadsheetDataValidationErrorStyle) -> &'static str {
    match value {
        NativeSpreadsheetDataValidationErrorStyle::Stop => "stop",
        NativeSpreadsheetDataValidationErrorStyle::Warning => "warning",
        NativeSpreadsheetDataValidationErrorStyle::Information => "information",
    }
}

fn rule_not_found(sheet: &str, index: usize) -> a3s_use_core::UseError {
    let path = format!("{sheet}/dataValidation[{index}]");
    editor_error(
        "use.office.node_not_found",
        format!("Office semantic path '{path}' does not exist."),
    )
    .with_detail("path", path)
}

fn invalid_collection(part_name: &str, reason: impl Into<String>) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_validation_invalid",
        format!(
            "Spreadsheet worksheet validation collection {}.",
            reason.into()
        ),
    )
    .with_detail("part", part_name)
}
