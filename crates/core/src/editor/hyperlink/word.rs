use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{
    add_external_relationship, ensure_namespace, existing_external_relationship,
    old_relationship_id, remove_relationship_if_unused,
};
use crate::editor::part::dialect;
use crate::editor::word::locate_word_path;
use crate::editor::{editor_error, node_not_found, parse_segments, prefix, qualified};
use crate::xml_edit::{
    apply_patches, index_xml, insert_child, patch_start_tag_attributes, replace_text_descendants,
    IndexedXmlElement, XmlPatch,
};
use crate::{
    LosslessXmlPart, NativeOfficeDocument, NativeOfficeHyperlink, NativeOfficeHyperlinkTarget,
    NativeOfficePackage, OfficeNodeType,
};

const DOCUMENT_PART: &str = "word/document.xml";

pub(super) fn set(
    package: &mut NativeOfficePackage,
    path: &str,
    hyperlink: &NativeOfficeHyperlink,
) -> UseResult<String> {
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let requested = snapshot.get(path, 0)?;
    let (paragraph_path, existing_path) = match requested.node_type {
        OfficeNodeType::Paragraph => (requested.path, None),
        OfficeNodeType::Hyperlink => (
            requested
                .path
                .rsplit_once('/')
                .map(|(parent, _)| parent.to_string())
                .ok_or_else(|| node_not_found(path))?,
            Some(requested.path),
        ),
        _ => {
            return Err(editor_error(
                "use.office.hyperlink_owner_unsupported",
                "Native Word hyperlinks require a paragraph or hyperlink path.",
            ))
        }
    };
    if let NativeOfficeHyperlinkTarget::Internal { location } = &hyperlink.target {
        validate_bookmark_name(location)?;
    }

    let owner = owner_part(&snapshot, &paragraph_path)?;
    let original = package.xml_part(&owner)?;
    let relationship_namespace = dialect(package)?.relationship_namespace();
    let (bytes, relationship_prefix) = match &hyperlink.target {
        NativeOfficeHyperlinkTarget::External { .. } => {
            ensure_namespace(&original, "r", relationship_namespace)?
        }
        NativeOfficeHyperlinkTarget::Internal { .. } => (original.raw().to_vec(), String::new()),
    };
    let part = LosslessXmlPart::parse(owner.clone(), bytes)?;
    let index = index_xml(&part)?;
    let paragraph = locate_part_path(&index, &paragraph_path)?;
    if paragraph.local_name != "p" {
        return Err(editor_error(
            "use.office.hyperlink_owner_unsupported",
            "Native Word hyperlinks require a paragraph owner.",
        ));
    }

    let existing = existing_path
        .as_deref()
        .map(|existing| locate_part_path(&index, existing))
        .transpose()?;
    let old_id = existing.and_then(old_relationship_id);
    let relationship_id = match &hyperlink.target {
        NativeOfficeHyperlinkTarget::External { uri } => {
            existing_external_relationship(package, &owner, old_id.as_deref(), uri)?
                .map(Ok)
                .unwrap_or_else(|| add_external_relationship(package, &owner, uri))?
                .into()
        }
        NativeOfficeHyperlinkTarget::Internal { .. } => None,
    };

    let link_path = if let Some(existing_path) = existing_path {
        let edited = update_existing(
            &part,
            &existing_path,
            hyperlink,
            relationship_id.as_deref(),
            &relationship_prefix,
        )?;
        package.set_part(&owner, edited)?;
        existing_path
    } else {
        let position = paragraph
            .children
            .iter()
            .filter(|child| child.local_name == "hyperlink")
            .count()
            + 1;
        let fragment = new_hyperlink_fragment(
            paragraph,
            hyperlink,
            relationship_id.as_deref(),
            &relationship_prefix,
        );
        package.set_part(&owner, insert_child(&part, paragraph, fragment)?)?;
        format!("{paragraph_path}/hyperlink[{position}]")
    };
    remove_relationship_if_unused(package, &owner, old_id.as_deref())?;
    Ok(link_path)
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let requested = snapshot.get(path, 0)?;
    if requested.node_type != OfficeNodeType::Hyperlink {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Native Word hyperlink removal requires a hyperlink path.",
        ));
    }
    let owner = owner_part(&snapshot, path)?;
    let part = package.xml_part(&owner)?;
    let index = index_xml(&part)?;
    let hyperlink = locate_part_path(&index, path)?;
    if hyperlink.local_name != "hyperlink" {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Native Word hyperlink removal requires a hyperlink path.",
        ));
    }
    let old_id = old_relationship_id(hyperlink);
    let edited = apply_patches(
        &part,
        vec![XmlPatch::new(hyperlink.full_range.clone(), Vec::new())],
    )?;
    package.set_part(&owner, edited)?;
    remove_relationship_if_unused(package, &owner, old_id.as_deref())
}

fn update_existing(
    part: &LosslessXmlPart,
    path: &str,
    hyperlink: &NativeOfficeHyperlink,
    relationship_id: Option<&str>,
    relationship_prefix: &str,
) -> UseResult<Vec<u8>> {
    let index = index_xml(part)?;
    let existing = locate_part_path(&index, path)?;
    let word_prefix = prefix(&existing.qualified_name);
    let mut updates = removal_updates(existing, &["id", "anchor", "tooltip"]);
    match &hyperlink.target {
        NativeOfficeHyperlinkTarget::External { .. } => {
            let id = relationship_id.ok_or_else(|| {
                editor_error(
                    "use.office.hyperlink_relationship_missing",
                    "Native Word external hyperlink has no relationship ID.",
                )
            })?;
            updates.insert(
                qualified(Some(relationship_prefix), "id"),
                Some(id.to_string()),
            );
        }
        NativeOfficeHyperlinkTarget::Internal { location } => {
            updates.insert(qualified(word_prefix, "anchor"), Some(location.to_string()));
        }
    }
    if let Some(tooltip) = &hyperlink.tooltip {
        updates.insert(qualified(word_prefix, "tooltip"), Some(tooltip.to_string()));
    }
    let bytes = patch_start_tag_attributes(part, existing, &updates)?;
    let Some(display) = hyperlink.display.as_deref() else {
        return Ok(bytes);
    };
    let edited = LosslessXmlPart::parse(part.name().to_string(), bytes)?;
    let edited_index = index_xml(&edited)?;
    let edited_link = locate_part_path(&edited_index, path)?;
    replace_text_descendants(&edited, edited_link, "t", display, None)
}

fn owner_part(snapshot: &NativeOfficeDocument, path: &str) -> UseResult<String> {
    if path.starts_with("/body/") {
        return Ok(DOCUMENT_PART.to_string());
    }
    let first = path
        .trim_start_matches('/')
        .split('/')
        .next()
        .ok_or_else(|| node_not_found(path))?;
    if !(first.starts_with("header[") || first.starts_with("footer[")) {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Native Word hyperlinks require body, header, or footer paragraph paths.",
        ));
    }
    snapshot
        .get(&format!("/{first}"), 0)?
        .format
        .get("part")
        .map(|part| part.trim_start_matches('/').to_string())
        .ok_or_else(|| {
            editor_error(
                "use.office.hyperlink_owner_unsupported",
                format!("Native Word hyperlink owner '{path}' has no source part."),
            )
        })
}

fn locate_part_path<'a>(
    root: &'a IndexedXmlElement,
    path: &str,
) -> UseResult<&'a IndexedXmlElement> {
    if root.local_name == "document" {
        return locate_word_path(root, path);
    }
    let expected_root = match root.local_name.as_str() {
        "hdr" => "header",
        "ftr" => "footer",
        _ => {
            return Err(editor_error(
                "use.office.word_xml_invalid",
                "Word hyperlink owner part must have a document, header, or footer root.",
            ))
        }
    };
    let segments = parse_segments(path)?;
    if segments.first().map(|segment| segment.name.as_str()) != Some(expected_root) {
        return Err(node_not_found(path));
    }
    let mut current = root;
    for segment in segments.into_iter().skip(1) {
        let local_name = match segment.name.as_str() {
            "p" | "paragraph" => "p",
            "tbl" | "table" => "tbl",
            "tr" => "tr",
            "tc" | "cell" => "tc",
            "hyperlink" => "hyperlink",
            name => {
                return Err(editor_error(
                    "use.office.mutation_path_unsupported",
                    format!("Word hyperlink path element '{name}' is not supported."),
                ))
            }
        };
        current = current
            .child(local_name, segment.position.unwrap_or(1))
            .ok_or_else(|| node_not_found(path))?;
    }
    Ok(current)
}

fn new_hyperlink_fragment(
    paragraph: &crate::xml_edit::IndexedXmlElement,
    hyperlink: &NativeOfficeHyperlink,
    relationship_id: Option<&str>,
    relationship_prefix: &str,
) -> String {
    let word_prefix = prefix(&paragraph.qualified_name);
    let hyperlink_tag = qualified(word_prefix, "hyperlink");
    let run_tag = qualified(word_prefix, "r");
    let properties_tag = qualified(word_prefix, "rPr");
    let color_tag = qualified(word_prefix, "color");
    let underline_tag = qualified(word_prefix, "u");
    let text_tag = qualified(word_prefix, "t");
    let value_attribute = qualified(word_prefix, "val");
    let target = match &hyperlink.target {
        NativeOfficeHyperlinkTarget::External { .. } => format!(
            " {}=\"{}\"",
            qualified(Some(relationship_prefix), "id"),
            crate::xml_edit::escape_attribute(relationship_id.unwrap_or_default())
        ),
        NativeOfficeHyperlinkTarget::Internal { location } => format!(
            " {}=\"{}\"",
            qualified(word_prefix, "anchor"),
            crate::xml_edit::escape_attribute(location)
        ),
    };
    let tooltip = hyperlink
        .tooltip
        .as_ref()
        .map_or_else(String::new, |value| {
            format!(
                " {}=\"{}\"",
                qualified(word_prefix, "tooltip"),
                crate::xml_edit::escape_attribute(value)
            )
        });
    let display = hyperlink
        .display
        .as_deref()
        .unwrap_or_else(|| hyperlink.default_display());
    let space =
        if display.starts_with(char::is_whitespace) || display.ends_with(char::is_whitespace) {
            " xml:space=\"preserve\""
        } else {
            ""
        };
    format!(
        "<{hyperlink_tag}{target}{tooltip}><{run_tag}><{properties_tag}><{color_tag} {value_attribute}=\"0563C1\" {}=\"hyperlink\"/><{underline_tag} {value_attribute}=\"single\"/></{properties_tag}><{text_tag}{space}>{}</{text_tag}></{run_tag}></{hyperlink_tag}>",
        qualified(word_prefix, "themeColor"),
        crate::xml_edit::escape_text(display)
    )
}

fn removal_updates(
    element: &crate::xml_edit::IndexedXmlElement,
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

fn validate_bookmark_name(location: &str) -> UseResult<()> {
    let valid = location.len() <= 40
        && location
            .bytes()
            .next()
            .is_some_and(|byte| byte.is_ascii_alphabetic() || byte == b'_')
        && location
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || byte == b'_');
    if !valid {
        return Err(editor_error(
            "use.office.hyperlink_location_invalid",
            "Native Word internal hyperlinks require a 1-40 character ASCII bookmark name beginning with a letter or underscore.",
        ));
    }
    Ok(())
}
