use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{
    add_external_relationship, ensure_namespace, existing_external_relationship,
    old_relationship_id, remove_relationship_if_unused,
};
use crate::editor::part::{dialect, relationship_part, relative_target};
use crate::editor::presentation::locate_path;
use crate::editor::{editor_error, node_not_found, qualified};
use crate::xml_edit::{
    apply_patches, index_xml, insert_ordered_child, patch_start_tag_attributes, IndexedXmlElement,
    XmlPatch,
};
use crate::{
    LosslessXmlPart, NativeOfficeDocument, NativeOfficeHyperlink, NativeOfficeHyperlinkTarget,
    NativeOfficePackage, OfficeNodeType, RelationshipSource, RelationshipTarget,
};

const SLIDE_JUMP_ACTION: &str = "ppaction://hlinksldjump";

#[derive(Debug)]
enum ResolvedTarget {
    External { uri: String },
    Slide { part: String },
}

pub(super) fn set(
    package: &mut NativeOfficePackage,
    path: &str,
    hyperlink: &NativeOfficeHyperlink,
) -> UseResult<String> {
    if hyperlink.display.is_some() {
        return Err(editor_error(
            "use.office.hyperlink_display_unsupported",
            "Presentation shape hyperlinks use the shape's existing text and do not accept separate display text.",
        ));
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let target = match &hyperlink.target {
        NativeOfficeHyperlinkTarget::External { uri } => {
            ResolvedTarget::External { uri: uri.clone() }
        }
        NativeOfficeHyperlinkTarget::Internal { location } => {
            let (_, part) = resolve_slide_target(&snapshot, location)?;
            ResolvedTarget::Slide { part }
        }
    };
    let requested = snapshot.get(path, 0)?;
    let owner_path = match requested.node_type {
        OfficeNodeType::Shape | OfficeNodeType::Placeholder => requested.path,
        OfficeNodeType::Hyperlink => requested
            .path
            .strip_suffix("/hyperlink")
            .map(str::to_string)
            .ok_or_else(|| node_not_found(path))?,
        _ => {
            return Err(editor_error(
                "use.office.hyperlink_owner_unsupported",
                "Native Presentation hyperlinks require a shape or shape hyperlink path.",
            ))
        }
    };
    let slide_path = owner_path
        .split('/')
        .find(|segment| segment.starts_with("slide["))
        .map(|segment| format!("/{segment}"))
        .ok_or_else(|| node_not_found(path))?;
    let slide = snapshot.get(&slide_path, 0)?;
    let owner = slide.format.get("part").ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            "Presentation semantic slide has no source part.",
        )
    })?;

    let original = package.xml_part(owner)?;
    let office_dialect = dialect(package)?;
    let (bytes, relationship_prefix) =
        ensure_namespace(&original, "r", office_dialect.relationship_namespace())?;
    let namespaced = LosslessXmlPart::parse(owner.clone(), bytes)?;
    let (bytes, drawing_prefix) =
        ensure_namespace(&namespaced, "a", office_dialect.drawing_namespace())?;
    let part = LosslessXmlPart::parse(owner.clone(), bytes)?;
    let index = index_xml(&part)?;
    let shape = locate_path(&index, &owner_path)?;
    let properties = shape.descendant("cNvPr").ok_or_else(|| {
        editor_error(
            "use.office.presentation_shape_invalid",
            format!("Presentation shape '{owner_path}' has no non-visual properties."),
        )
    })?;
    let existing = properties.child("hlinkClick", 1);
    let old_id = existing.and_then(old_relationship_id);
    let relationship_id = match &target {
        ResolvedTarget::External { uri } => {
            existing_external_relationship(package, owner, old_id.as_deref(), uri)?
                .map(Ok)
                .unwrap_or_else(|| add_external_relationship(package, owner, uri))?
        }
        ResolvedTarget::Slide { part, .. } => {
            existing_slide_relationship(package, owner, old_id.as_deref(), part)?
                .map(Ok)
                .unwrap_or_else(|| add_slide_relationship(package, owner, part))?
        }
    };
    let action = matches!(target, ResolvedTarget::Slide { .. }).then_some(SLIDE_JUMP_ACTION);

    let edited = if let Some(existing) = existing {
        update_existing(
            &part,
            existing,
            &relationship_id,
            &relationship_prefix,
            hyperlink.tooltip.as_deref(),
            action,
        )?
    } else {
        let tooltip = hyperlink
            .tooltip
            .as_ref()
            .map_or_else(String::new, |value| {
                format!(" tooltip=\"{}\"", crate::xml_edit::escape_attribute(value))
            });
        let tag = qualified(Some(&drawing_prefix), "hlinkClick");
        let id = qualified(Some(&relationship_prefix), "id");
        let action = action.map_or_else(String::new, |value| {
            format!(" action=\"{}\"", crate::xml_edit::escape_attribute(value))
        });
        insert_ordered_child(
            &part,
            properties,
            format!(
                "<{tag} {id}=\"{}\"{tooltip}{action}/>",
                crate::xml_edit::escape_attribute(&relationship_id)
            ),
            &["hlinkHover", "extLst"],
        )?
    };
    package.set_part(owner, edited)?;
    remove_relationship_if_unused(package, owner, old_id.as_deref())?;
    Ok(format!("{owner_path}/hyperlink"))
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let owner_path = path.strip_suffix("/hyperlink").ok_or_else(|| {
        editor_error(
            "use.office.mutation_type_unsupported",
            "Native Presentation hyperlink removal requires a shape hyperlink path.",
        )
    })?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let slide_path = owner_path
        .split('/')
        .find(|segment| segment.starts_with("slide["))
        .map(|segment| format!("/{segment}"))
        .ok_or_else(|| node_not_found(path))?;
    let slide = snapshot.get(&slide_path, 0)?;
    let owner = slide.format.get("part").ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            "Presentation semantic slide has no source part.",
        )
    })?;
    let part = package.xml_part(owner)?;
    let index = index_xml(&part)?;
    let shape = locate_path(&index, owner_path)?;
    let hyperlink = shape
        .descendant("cNvPr")
        .and_then(|properties| properties.child("hlinkClick", 1))
        .ok_or_else(|| node_not_found(path))?;
    let old_id = old_relationship_id(hyperlink);
    package.set_part(
        owner,
        apply_patches(
            &part,
            vec![XmlPatch::new(hyperlink.full_range.clone(), Vec::new())],
        )?,
    )?;
    remove_relationship_if_unused(package, owner, old_id.as_deref())
}

fn update_existing(
    part: &LosslessXmlPart,
    existing: &IndexedXmlElement,
    relationship_id: &str,
    relationship_prefix: &str,
    tooltip: Option<&str>,
    action: Option<&str>,
) -> UseResult<Vec<u8>> {
    let mut updates = removal_updates(existing, &["id", "tooltip", "action"]);
    updates.insert(
        qualified(Some(relationship_prefix), "id"),
        Some(relationship_id.to_string()),
    );
    if let Some(tooltip) = tooltip {
        updates.insert("tooltip".into(), Some(tooltip.to_string()));
    }
    if let Some(action) = action {
        updates.insert("action".into(), Some(action.to_string()));
    }
    patch_start_tag_attributes(part, existing, &updates)
}

fn resolve_slide_target(
    snapshot: &NativeOfficeDocument,
    location: &str,
) -> UseResult<(String, String)> {
    let path = if location.starts_with('/') {
        location.to_string()
    } else {
        format!("/{location}")
    };
    let position = path
        .strip_prefix("/slide[")
        .and_then(|value| value.strip_suffix(']'))
        .filter(|value| !value.contains('/'))
        .and_then(|value| value.parse::<usize>().ok())
        .filter(|value| *value > 0)
        .ok_or_else(|| invalid_slide_location(location))?;
    let path = format!("/slide[{position}]");
    let slide = snapshot
        .get(&path, 0)
        .ok()
        .filter(|node| node.node_type == OfficeNodeType::Slide)
        .ok_or_else(|| invalid_slide_location(location))?;
    let part = slide.format.get("part").cloned().ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            format!("Presentation hyperlink target '{path}' has no source part."),
        )
    })?;
    Ok((path, part.trim_start_matches('/').to_string()))
}

fn invalid_slide_location(location: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.hyperlink_location_invalid",
        "Native Presentation internal hyperlinks require an existing slide[N] or /slide[N] target.",
    )
    .with_detail("location", location)
}

fn existing_slide_relationship(
    package: &NativeOfficePackage,
    owner: &str,
    id: Option<&str>,
    target_part: &str,
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
            relationship.relationship_type.ends_with("/slide")
                && matches!(
                    &relationship.target,
                    RelationshipTarget::Internal { part_name, fragment: None }
                        if part_name == target_part
                )
        });
    Ok(matched.then(|| id.to_string()))
}

fn add_slide_relationship(
    package: &mut NativeOfficePackage,
    owner: &str,
    target_part: &str,
) -> UseResult<String> {
    let office_dialect = dialect(package)?;
    crate::opc_edit::add_relationship(
        package,
        &relationship_part(owner),
        &office_dialect.relationship_type("slide"),
        &relative_target(owner, target_part),
    )
}

fn removal_updates(
    element: &IndexedXmlElement,
    local_names: &[&str],
) -> BTreeMap<String, Option<String>> {
    element
        .qualified_attributes
        .keys()
        .filter(|name| {
            let local = name
                .rsplit_once(':')
                .map_or(name.as_str(), |(_, local)| local);
            local_names.contains(&local)
        })
        .map(|name| (name.clone(), None))
        .collect()
}
