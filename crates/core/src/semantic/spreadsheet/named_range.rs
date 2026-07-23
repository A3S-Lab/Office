use std::collections::BTreeSet;

use a3s_use_core::UseResult;

use super::{direct_text, semantic_error, DocumentNode, OfficeNodeType, XmlElement, XmlNode};
use crate::spreadsheet_named_range::{
    canonical_named_range_path, is_protected_named_range, named_range_scope_label,
    parse_named_range_path, validate_named_range_comment, validate_named_range_name,
    validate_named_range_reference, NamedRangePathSelector, MAX_SPREADSHEET_NAMED_RANGES,
};

pub(super) fn read(
    workbook: &XmlElement,
    sheet_names: &[String],
) -> UseResult<Option<DocumentNode>> {
    let containers = workbook
        .child_elements()
        .filter(|element| element.local_name == "definedNames")
        .collect::<Vec<_>>();
    if containers.len() > 1 {
        return Err(invalid("contains multiple definedNames collections"));
    }
    let Some(container) = containers.first().copied() else {
        return Ok(None);
    };
    if container.namespace != workbook.namespace {
        return Err(invalid("uses definedNames in an unexpected namespace"));
    }
    let names = container
        .child_elements()
        .filter(|element| {
            element.local_name == "definedName" && element.namespace == workbook.namespace
        })
        .collect::<Vec<_>>();
    if names.len() > MAX_SPREADSHEET_NAMED_RANGES {
        return Err(semantic_error(
            "use.office.spreadsheet_named_range_limit",
            format!(
                "Spreadsheet workbook contains {} defined names; the limit is {MAX_SPREADSHEET_NAMED_RANGES}.",
                names.len()
            ),
        ));
    }

    let mut identities = BTreeSet::new();
    let mut collection = DocumentNode::new(
        "/namedrange",
        "namedranges",
        OfficeNodeType::NamedRangeCollection,
    );
    for (offset, element) in names.into_iter().enumerate() {
        if element
            .children
            .iter()
            .any(|child| matches!(child, XmlNode::Element(_)))
        {
            return Err(invalid(format!(
                "contains definedName[{}] with nested element content",
                offset + 1
            )));
        }
        let name = unqualified_attribute(element, "name")
            .ok_or_else(|| invalid(format!("contains definedName[{}] without name", offset + 1)))?;
        validate_named_range_name(name)
            .map_err(|error| invalid(format!("contains invalid defined name '{name}': {error}")))?;
        let reference = direct_text(element);
        validate_named_range_reference(&reference).map_err(|error| {
            invalid(format!(
                "contains invalid ref for defined name '{name}': {error}"
            ))
        })?;
        let (scope, local_sheet_id) = match unqualified_attribute(element, "localSheetId") {
            Some(value) => {
                let index = value
                    .parse::<usize>()
                    .ok()
                    .filter(|index| *index < sheet_names.len())
                    .ok_or_else(|| {
                        invalid(format!(
                            "contains defined name '{name}' with invalid localSheetId='{value}'"
                        ))
                    })?;
                (sheet_names[index].clone(), Some(index))
            }
            None => ("workbook".to_string(), None),
        };
        let scope = named_range_scope_label(&scope, local_sheet_id.is_some());
        let volatile = boolean_attribute(element, "function", false, name)?;
        let comment = unqualified_attribute(element, "comment");
        validate_named_range_comment(comment)
            .map_err(|error| invalid(format!("contains invalid comment for '{name}': {error}")))?;
        if !identities.insert((name.to_ascii_lowercase(), local_sheet_id)) {
            return Err(semantic_error(
                "use.office.spreadsheet_named_range_duplicate",
                format!("Spreadsheet defined name '{name}' is duplicated in scope '{scope}'."),
            )
            .with_detail("name", name)
            .with_detail("scope", scope));
        }

        let path = canonical_named_range_path(name, &scope);
        let mut node = DocumentNode::new(path, "namedrange", OfficeNodeType::NamedRange);
        node.text = reference.clone();
        node.preview = Some(format!("{name} → {reference}"));
        node.format.insert("name".into(), name.to_string());
        node.format.insert("ref".into(), reference);
        node.format.insert("scope".into(), scope);
        node.format.insert("volatile".into(), volatile.to_string());
        if let Some(comment) = comment {
            node.format.insert("comment".into(), comment.to_string());
        }
        if is_protected_named_range(name) {
            node.format.insert("protected".into(), "true".into());
        }
        collection.children.push(node);
    }
    Ok((!collection.children.is_empty()).then_some(collection))
}

pub(super) fn virtual_get(
    root: &DocumentNode,
    path: &str,
    depth: usize,
) -> UseResult<Option<DocumentNode>> {
    let Some(selector) = parse_named_range_path(path)? else {
        return Ok(None);
    };
    let collection = root
        .children
        .iter()
        .find(|node| node.node_type == OfficeNodeType::NamedRangeCollection);
    match selector {
        NamedRangePathSelector::Collection => Ok(Some(collection.map_or_else(
            || {
                DocumentNode::new(
                    "/namedrange",
                    "namedranges",
                    OfficeNodeType::NamedRangeCollection,
                )
            },
            |node| node.clone_to_depth(depth),
        ))),
        NamedRangePathSelector::Position(position) => Ok(collection
            .and_then(|collection| collection.children.get(position - 1))
            .map(|node| node.clone_to_depth(depth))),
        NamedRangePathSelector::Name { name, scope } => {
            let matches = collection
                .into_iter()
                .flat_map(|collection| &collection.children)
                .filter(|node| {
                    node.format
                        .get("name")
                        .is_some_and(|candidate| candidate.eq_ignore_ascii_case(&name))
                        && scope.as_deref().is_none_or(|scope| {
                            node.format
                                .get("scope")
                                .is_some_and(|candidate| candidate.eq_ignore_ascii_case(scope))
                        })
                })
                .collect::<Vec<_>>();
            if matches.len() > 1 {
                return Err(semantic_error(
                    "use.office.spreadsheet_named_range_ambiguous",
                    format!("Spreadsheet defined name '{name}' exists in multiple scopes."),
                )
                .with_suggestion(
                    "Add [@scope=workbook] or [@scope=SheetName] to select one stable identity.",
                )
                .with_detail("name", name)
                .with_detail("matches", matches.len()));
            }
            Ok(matches
                .into_iter()
                .next()
                .map(|node| node.clone_to_depth(depth)))
        }
    }
}

fn boolean_attribute(
    element: &XmlElement,
    attribute: &str,
    default: bool,
    name: &str,
) -> UseResult<bool> {
    match unqualified_attribute(element, attribute) {
        None => Ok(default),
        Some("1" | "true") => Ok(true),
        Some("0" | "false") => Ok(false),
        Some(value) => Err(invalid(format!(
            "contains defined name '{name}' with invalid {attribute}='{value}'"
        ))),
    }
}

fn unqualified_attribute<'a>(element: &'a XmlElement, local_name: &str) -> Option<&'a str> {
    element
        .attributes
        .iter()
        .find(|attribute| attribute.namespace.is_none() && attribute.local_name == local_name)
        .map(|attribute| attribute.value.as_str())
}

fn invalid(reason: impl Into<String>) -> a3s_use_core::UseError {
    semantic_error(
        "use.office.spreadsheet_named_range_invalid",
        format!("Spreadsheet workbook.xml {}.", reason.into()),
    )
    .with_detail("part", "xl/workbook.xml")
}
