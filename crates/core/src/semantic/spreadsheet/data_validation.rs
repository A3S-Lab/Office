use a3s_use_core::UseResult;

use super::{direct_text, semantic_error, DocumentNode, OfficeNodeType, XmlElement};
use crate::spreadsheet_reference::{
    first_intersecting_ranges, CellRange, MAX_DATA_VALIDATIONS, MAX_DATA_VALIDATION_RANGES,
};

pub(super) fn read(
    worksheet: &XmlElement,
    part_name: &str,
    sheet_path: &str,
) -> UseResult<Vec<DocumentNode>> {
    let containers = worksheet
        .child_elements()
        .filter(|element| {
            element.local_name == "dataValidations" && element.namespace == worksheet.namespace
        })
        .collect::<Vec<_>>();
    if containers.len() > 1 {
        return Err(invalid(
            part_name,
            "contains multiple dataValidations collections",
        ));
    }
    let Some(container) = containers.first() else {
        return Ok(Vec::new());
    };
    let rules = container
        .child_elements()
        .filter(|element| {
            element.local_name == "dataValidation" && element.namespace == worksheet.namespace
        })
        .collect::<Vec<_>>();
    if rules.len() > MAX_DATA_VALIDATIONS {
        return Err(semantic_error(
            "use.office.spreadsheet_validation_limit",
            format!(
                "Worksheet part '{part_name}' contains {} data validations; the limit is {MAX_DATA_VALIDATIONS}.",
                rules.len()
            ),
        ));
    }

    let mut nodes = Vec::with_capacity(rules.len());
    let mut all_ranges = Vec::new();
    let mut owners = Vec::new();
    for (offset, rule) in rules.into_iter().enumerate() {
        let index = offset + 1;
        let reference = unqualified_attribute(rule, "sqref")
            .ok_or_else(|| invalid(part_name, "contains a dataValidation without sqref"))?;
        let normalized = normalize_ranges(reference, part_name)?;
        for range in &normalized {
            all_ranges.push(*range);
            owners.push(index);
        }

        let validation_type = unqualified_attribute(rule, "type").ok_or_else(|| {
            invalid(
                part_name,
                format!("contains dataValidation[{index}] without type"),
            )
        })?;
        if !matches!(
            validation_type,
            "list" | "whole" | "decimal" | "date" | "time" | "textLength" | "custom"
        ) {
            return Err(invalid(
                part_name,
                format!(
                    "contains dataValidation[{index}] with unsupported type '{validation_type}'"
                ),
            ));
        }

        let mut node = DocumentNode::new(
            format!("{sheet_path}/dataValidation[{index}]"),
            "dataValidation",
            OfficeNodeType::DataValidation,
        );
        let reference = normalized
            .iter()
            .map(|range| range.a1())
            .collect::<Vec<_>>()
            .join(" ");
        node.text = reference.clone();
        node.preview = Some(format!("{validation_type} ({reference})"));
        node.format.insert("ref".into(), reference);
        node.format
            .insert("type".into(), validation_type.to_string());

        if let Some(operator) = unqualified_attribute(rule, "operator") {
            if !matches!(
                operator,
                "between"
                    | "notBetween"
                    | "equal"
                    | "notEqual"
                    | "greaterThan"
                    | "greaterThanOrEqual"
                    | "lessThan"
                    | "lessThanOrEqual"
            ) {
                return Err(invalid(
                    part_name,
                    format!(
                        "contains dataValidation[{index}] with unsupported operator '{operator}'"
                    ),
                ));
            }
            node.format.insert("operator".into(), operator.to_string());
        } else if matches!(
            validation_type,
            "whole" | "decimal" | "date" | "time" | "textLength"
        ) {
            node.format.insert("operator".into(), "between".into());
        }
        if let Some(error_style) = unqualified_attribute(rule, "errorStyle") {
            if !matches!(error_style, "stop" | "warning" | "information") {
                return Err(invalid(
                    part_name,
                    format!(
                        "contains dataValidation[{index}] with unsupported errorStyle '{error_style}'"
                    ),
                ));
            }
            node.format
                .insert("errorStyle".into(), error_style.to_string());
        } else {
            node.format.insert("errorStyle".into(), "stop".into());
        }

        let allow_blank = boolean_attribute(rule, "allowBlank", false, part_name, index)?;
        let show_input = boolean_attribute(rule, "showInputMessage", false, part_name, index)?;
        let show_error = boolean_attribute(rule, "showErrorMessage", false, part_name, index)?;
        let hide_dropdown = boolean_attribute(rule, "showDropDown", false, part_name, index)?;
        node.format
            .insert("allowBlank".into(), allow_blank.to_string());
        node.format
            .insert("showInput".into(), show_input.to_string());
        node.format
            .insert("showError".into(), show_error.to_string());
        node.format
            .insert("inCellDropdown".into(), (!hide_dropdown).to_string());

        for (attribute, key) in [
            ("promptTitle", "promptTitle"),
            ("prompt", "prompt"),
            ("errorTitle", "errorTitle"),
            ("error", "error"),
        ] {
            if let Some(value) = unqualified_attribute(rule, attribute) {
                node.format.insert(key.into(), value.into());
            }
        }
        for (child_name, key) in [("formula1", "formula1"), ("formula2", "formula2")] {
            let children = rule
                .child_elements()
                .filter(|child| {
                    child.local_name == child_name && child.namespace == worksheet.namespace
                })
                .collect::<Vec<_>>();
            if children.len() > 1 {
                return Err(invalid(
                    part_name,
                    format!("contains dataValidation[{index}] with multiple {child_name} elements"),
                ));
            }
            if let Some(child) = children.first() {
                let value = direct_text(child);
                if value.chars().count() > 255 {
                    return Err(invalid(
                        part_name,
                        format!(
                            "contains dataValidation[{index}] {child_name} longer than 255 characters"
                        ),
                    ));
                }
                node.format.insert(key.into(), value);
            }
        }
        nodes.push(node);
    }

    if let Some((left, right)) = first_intersecting_ranges(&all_ranges) {
        return Err(semantic_error(
            "use.office.spreadsheet_validation_overlap",
            format!(
                "Worksheet part '{part_name}' contains overlapping validation ranges '{}' and '{}'.",
                all_ranges[left].a1(),
                all_ranges[right].a1()
            ),
        )
        .with_detail("leftRule", owners[left])
        .with_detail("rightRule", owners[right])
        .with_detail("left", all_ranges[left].a1())
        .with_detail("right", all_ranges[right].a1()));
    }
    Ok(nodes)
}

fn normalize_ranges(reference: &str, part_name: &str) -> UseResult<Vec<CellRange>> {
    let tokens = reference.split_ascii_whitespace().collect::<Vec<_>>();
    if tokens.is_empty() || tokens.len() > MAX_DATA_VALIDATION_RANGES {
        return Err(invalid(
            part_name,
            format!(
                "contains a dataValidation with {} ranges; expected 1-{MAX_DATA_VALIDATION_RANGES}",
                tokens.len()
            ),
        ));
    }
    tokens
        .into_iter()
        .map(|token| {
            CellRange::parse(token).map_err(|error| {
                invalid(
                    part_name,
                    format!("contains invalid data-validation range '{token}': {error}"),
                )
            })
        })
        .collect()
}

fn boolean_attribute(
    element: &XmlElement,
    name: &str,
    default: bool,
    part_name: &str,
    index: usize,
) -> UseResult<bool> {
    match unqualified_attribute(element, name) {
        None => Ok(default),
        Some("1" | "true") => Ok(true),
        Some("0" | "false") => Ok(false),
        Some(value) => Err(invalid(
            part_name,
            format!("contains dataValidation[{index}] with invalid {name}='{value}'"),
        )),
    }
}

fn unqualified_attribute<'a>(element: &'a XmlElement, local_name: &str) -> Option<&'a str> {
    element
        .attributes
        .iter()
        .find(|attribute| attribute.namespace.is_none() && attribute.local_name == local_name)
        .map(|attribute| attribute.value.as_str())
}

fn invalid(part_name: &str, reason: impl Into<String>) -> a3s_use_core::UseError {
    semantic_error(
        "use.office.spreadsheet_validation_invalid",
        format!(
            "Spreadsheet worksheet part '{part_name}' {}.",
            reason.into()
        ),
    )
    .with_detail("part", part_name)
}
