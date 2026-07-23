use std::collections::BTreeSet;

use a3s_use_core::{UseError, UseResult};

use crate::discovery::office_error;
use crate::xml_edit::{index_xml, insert_child};
use crate::{LosslessXmlPart, NativeOfficePackage};

const EMPTY_RELATIONSHIPS: &[u8] =
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;

pub(crate) fn add_content_type_override(
    package: &mut NativeOfficePackage,
    part_name: &str,
    content_type: &str,
) -> UseResult<()> {
    let part_name = format!("/{}", part_name.trim_start_matches('/'));
    let part = package.xml_part("[Content_Types].xml")?;
    let index = index_xml(&part)?;
    for child in index
        .children
        .iter()
        .filter(|child| child.local_name == "Override")
    {
        if child
            .attributes
            .get("PartName")
            .is_some_and(|existing| existing.eq_ignore_ascii_case(&part_name))
        {
            return if child.attributes.get("ContentType").map(String::as_str) == Some(content_type)
            {
                Ok(())
            } else {
                Err(opc_edit_error(
                    "use.office.content_type_conflict",
                    format!("OOXML part '{part_name}' already has a different content type."),
                ))
            };
        }
    }
    let tag = qualified(prefix(&index.qualified_name), "Override");
    let fragment = format!(
        "<{tag} PartName=\"{}\" ContentType=\"{}\"/>",
        escape_attribute(&part_name),
        escape_attribute(content_type)
    );
    let edited = insert_child(&part, &index, fragment)?;
    package.set_part("[Content_Types].xml", edited)
}

pub(crate) fn add_relationship(
    package: &mut NativeOfficePackage,
    relationship_part: &str,
    relationship_type: &str,
    target: &str,
) -> UseResult<String> {
    add_relationship_with_mode(package, relationship_part, relationship_type, target, None)
}

pub(crate) fn add_external_relationship(
    package: &mut NativeOfficePackage,
    relationship_part: &str,
    relationship_type: &str,
    target: &str,
) -> UseResult<String> {
    add_relationship_with_mode(
        package,
        relationship_part,
        relationship_type,
        target,
        Some("External"),
    )
}

fn add_relationship_with_mode(
    package: &mut NativeOfficePackage,
    relationship_part: &str,
    relationship_type: &str,
    target: &str,
    target_mode: Option<&str>,
) -> UseResult<String> {
    let part = if package.contains_part(relationship_part) {
        package.xml_part(relationship_part)?
    } else {
        LosslessXmlPart::parse(relationship_part.to_string(), EMPTY_RELATIONSHIPS)?
    };
    let index = index_xml(&part)?;
    if index.local_name != "Relationships" {
        return Err(opc_edit_error(
            "use.office.relationships_invalid",
            format!("OOXML relationship part '{relationship_part}' has an invalid root."),
        ));
    }
    let ids = index
        .children
        .iter()
        .filter(|child| child.local_name == "Relationship")
        .filter_map(|child| child.attributes.get("Id").cloned())
        .collect::<BTreeSet<_>>();
    let id = (1..=package.limits().max_entries.saturating_add(1))
        .map(|number| format!("rId{number}"))
        .find(|candidate| !ids.contains(candidate))
        .ok_or_else(|| {
            opc_edit_error(
                "use.office.relationship_id_exhausted",
                format!("OOXML relationship part '{relationship_part}' has no free ID."),
            )
        })?;
    let tag = qualified(prefix(&index.qualified_name), "Relationship");
    let target_mode = target_mode
        .map(|mode| format!(" TargetMode=\"{}\"", escape_attribute(mode)))
        .unwrap_or_default();
    let fragment = format!(
        "<{tag} Id=\"{}\" Type=\"{}\" Target=\"{}\"{target_mode}/>",
        escape_attribute(&id),
        escape_attribute(relationship_type),
        escape_attribute(target)
    );
    let edited = insert_child(&part, &index, fragment)?;
    package.set_part(relationship_part, edited)?;
    Ok(id)
}

pub(crate) fn remove_content_type_override(
    package: &mut NativeOfficePackage,
    part_name: &str,
) -> UseResult<()> {
    let part_name = format!("/{}", part_name.trim_start_matches('/'));
    let part = package.xml_part("[Content_Types].xml")?;
    let index = index_xml(&part)?;
    let target = index
        .children
        .iter()
        .filter(|child| child.local_name == "Override")
        .find(|child| {
            child
                .attributes
                .get("PartName")
                .is_some_and(|existing| existing.eq_ignore_ascii_case(&part_name))
        })
        .ok_or_else(|| {
            opc_edit_error(
                "use.office.content_type_missing",
                format!("OOXML part '{part_name}' has no content type override."),
            )
        })?;
    let edited = crate::xml_edit::apply_patches(
        &part,
        vec![crate::xml_edit::XmlPatch::new(
            target.full_range.clone(),
            Vec::new(),
        )],
    )?;
    package.set_part("[Content_Types].xml", edited)
}

pub(crate) fn remove_relationship(
    package: &mut NativeOfficePackage,
    relationship_part: &str,
    id: &str,
) -> UseResult<()> {
    let part = package.xml_part(relationship_part)?;
    let index = index_xml(&part)?;
    let target = index
        .children
        .iter()
        .filter(|child| child.local_name == "Relationship")
        .find(|child| child.qualified_attributes.get("Id").map(String::as_str) == Some(id))
        .ok_or_else(|| {
            opc_edit_error(
                "use.office.relationship_missing",
                format!(
                    "OOXML relationship part '{relationship_part}' has no relationship '{id}'."
                ),
            )
        })?;
    let edited = crate::xml_edit::apply_patches(
        &part,
        vec![crate::xml_edit::XmlPatch::new(
            target.full_range.clone(),
            Vec::new(),
        )],
    )?;
    package.set_part(relationship_part, edited)
}

fn prefix(qualified_name: &str) -> Option<&str> {
    qualified_name.rsplit_once(':').map(|(prefix, _)| prefix)
}

fn qualified(prefix: Option<&str>, local_name: &str) -> String {
    prefix.map_or_else(
        || local_name.to_string(),
        |prefix| format!("{prefix}:{local_name}"),
    )
}

fn escape_attribute(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

fn opc_edit_error(code: &str, message: impl Into<String>) -> UseError {
    office_error(code, message)
}
