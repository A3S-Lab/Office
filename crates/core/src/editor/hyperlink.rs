use a3s_use_core::UseResult;

use super::part::{dialect, ensure_namespace as ensure_part_namespace, relationship_part};
use super::NativeOfficeHyperlink;
use crate::xml_edit::{index_xml, IndexedXmlElement};
use crate::{DocumentKind, LosslessXmlPart, NativeOfficePackage, RelationshipSource};

mod presentation;
mod spreadsheet;
mod word;

pub(super) fn set(
    package: &mut NativeOfficePackage,
    path: &str,
    hyperlink: &NativeOfficeHyperlink,
) -> UseResult<String> {
    match package.kind() {
        DocumentKind::Word => word::set(package, path, hyperlink),
        DocumentKind::Spreadsheet => spreadsheet::set(package, path, hyperlink),
        DocumentKind::Presentation => presentation::set(package, path, hyperlink),
    }
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    match package.kind() {
        DocumentKind::Word => word::remove(package, path),
        DocumentKind::Spreadsheet => spreadsheet::remove(package, path),
        DocumentKind::Presentation => presentation::remove(package, path),
    }
}

pub(super) fn remove_spreadsheet_range_links(
    package: &mut NativeOfficePackage,
    path: &str,
) -> UseResult<()> {
    spreadsheet::remove_for_range(package, path)
}

pub(super) fn add_external_relationship(
    package: &mut NativeOfficePackage,
    owner: &str,
    uri: &str,
) -> UseResult<String> {
    let dialect = dialect(package)?;
    crate::opc_edit::add_external_relationship(
        package,
        &relationship_part(owner),
        &dialect.relationship_type("hyperlink"),
        uri,
    )
}

pub(super) fn existing_external_relationship(
    package: &NativeOfficePackage,
    owner: &str,
    id: Option<&str>,
    uri: &str,
) -> UseResult<Option<String>> {
    let Some(id) = id else {
        return Ok(None);
    };
    let source = RelationshipSource::Part {
        part_name: owner.to_string(),
    };
    let matched = package
        .opc_model()?
        .relationships()
        .relationship(&source, id)
        .is_some_and(|relationship| {
            relationship.relationship_type.ends_with("/hyperlink")
                && matches!(
                    &relationship.target,
                    crate::RelationshipTarget::External { uri: existing } if existing == uri
                )
        });
    Ok(matched.then(|| id.to_string()))
}

pub(super) fn remove_relationship_if_unused(
    package: &mut NativeOfficePackage,
    owner: &str,
    id: Option<&str>,
) -> UseResult<()> {
    let Some(id) = id else {
        return Ok(());
    };
    let source = RelationshipSource::Part {
        part_name: owner.to_string(),
    };
    let is_owned_action = package
        .opc_model()?
        .relationships()
        .relationship(&source, id)
        .is_some_and(|relationship| {
            relationship.relationship_type.ends_with("/hyperlink")
                || relationship.relationship_type.ends_with("/slide")
        });
    if !is_owned_action || xml_references_id(package, owner, id)? {
        return Ok(());
    }
    crate::opc_edit::remove_relationship(package, &relationship_part(owner), id)
}

fn xml_references_id(package: &NativeOfficePackage, owner: &str, id: &str) -> UseResult<bool> {
    let part = package.xml_part(owner)?;
    let root = index_xml(&part)?;
    Ok(element_references_id(&root, id))
}

fn element_references_id(element: &IndexedXmlElement, id: &str) -> bool {
    element.qualified_attributes.iter().any(|(name, value)| {
        name.rsplit_once(':')
            .is_some_and(|(_, local)| matches!(local, "id" | "embed" | "link") && value == id)
    }) || element
        .children
        .iter()
        .any(|child| element_references_id(child, id))
}

pub(super) fn ensure_namespace(
    part: &LosslessXmlPart,
    preferred: &str,
    namespace: &str,
) -> UseResult<(Vec<u8>, String)> {
    ensure_part_namespace(
        part,
        preferred,
        namespace,
        "use.office.hyperlink_namespace_exhausted",
    )
}

pub(super) fn old_relationship_id(element: &IndexedXmlElement) -> Option<String> {
    element
        .qualified_attributes
        .iter()
        .find_map(|(name, value)| {
            name.rsplit_once(':')
                .filter(|(_, local)| *local == "id")
                .map(|_| value.clone())
        })
}

pub(super) fn owned_hyperlink_relationship_ids(element: &IndexedXmlElement) -> Vec<String> {
    let mut output = Vec::new();
    collect_owned_hyperlink_relationship_ids(element, &mut output);
    output.sort();
    output.dedup();
    output
}

fn collect_owned_hyperlink_relationship_ids(element: &IndexedXmlElement, output: &mut Vec<String>) {
    if matches!(element.local_name.as_str(), "hyperlink" | "hlinkClick") {
        if let Some(id) = old_relationship_id(element) {
            output.push(id);
        }
    }
    for child in &element.children {
        collect_owned_hyperlink_relationship_ids(child, output);
    }
}

pub(super) fn remove_relationships_if_unused(
    package: &mut NativeOfficePackage,
    owner: &str,
    ids: &[String],
) -> UseResult<()> {
    for id in ids {
        remove_relationship_if_unused(package, owner, Some(id))?;
    }
    Ok(())
}
