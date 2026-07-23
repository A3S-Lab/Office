use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{
    comment_error, derive_initials, ensure_related_part, related_part, remove_related_part,
    CreateRelatedPart, PRESENTATION_AUTHORS_CONTENT_TYPE, PRESENTATION_COMMENTS_CONTENT_TYPE,
};
use crate::editor::part::dialect;
use crate::editor::{
    node_not_found, prefix, qualified, NativeOfficeComment, NativeOfficeCommentUpdate,
};
use crate::xml_edit::{
    apply_patches, index_xml, insert_child, patch_start_tag_attributes,
    replace_namespaced_text_descendants, IndexedXmlElement, XmlPatch,
};
use crate::{NativeOfficeDocument, NativeOfficePackage, OfficeNodeType};

const PRESENTATION_PART: &str = "ppt/presentation.xml";

pub(super) fn add(
    package: &mut NativeOfficePackage,
    parent: &str,
    comment: &NativeOfficeComment,
) -> UseResult<String> {
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let slide = snapshot.get(parent, 0)?;
    if slide.node_type != OfficeNodeType::Slide {
        return Err(comment_error(
            "use.office.comment_parent_unsupported",
            "Native Presentation comments require a slide parent such as /slide[1].",
        ));
    }
    let slide_part = slide.format.get("part").cloned().ok_or_else(|| {
        comment_error(
            "use.office.presentation_slide_invalid",
            "Presentation semantic slide has no source part.",
        )
    })?;
    let office_dialect = dialect(package)?;
    let authors = ensure_related_part(
        package,
        CreateRelatedPart {
            owner: PRESENTATION_PART,
            directory: "ppt",
            stem: "commentAuthors",
            extension: "xml",
            content_type: PRESENTATION_AUTHORS_CONTENT_TYPE,
            relationship_name: "commentAuthors",
            xml: format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:cmAuthorLst xmlns:p=\"{}\"/>",
                office_dialect.presentation_namespace()
            ),
        },
    )?;
    let initials = comment
        .initials
        .clone()
        .unwrap_or_else(|| derive_initials(&comment.author));
    let author_id = ensure_author(package, &authors.part_name, &comment.author, &initials)?;
    let index = issue_author_index(package, &authors.part_name, author_id)?;
    let comments = ensure_related_part(
        package,
        CreateRelatedPart {
            owner: &slide_part,
            directory: "ppt/comments",
            stem: "comment",
            extension: "xml",
            content_type: PRESENTATION_COMMENTS_CONTENT_TYPE,
            relationship_name: "comments",
            xml: format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:cmLst xmlns:p=\"{}\"/>",
                office_dialect.presentation_namespace()
            ),
        },
    )?;
    let part = package.xml_part(&comments.part_name)?;
    let root = index_xml(&part)?;
    require_comment_list(&root, &comments.part_name)?;
    let position = root
        .children
        .iter()
        .filter(|child| child.local_name == "cm")
        .count()
        + 1;
    let coordinates = comment
        .position
        .unwrap_or(crate::NativeOfficeCommentPosition::new(0, 0));
    let presentation_prefix = prefix(&root.qualified_name);
    let comment_tag = qualified(presentation_prefix, "cm");
    let position_tag = qualified(presentation_prefix, "pos");
    let text_tag = qualified(presentation_prefix, "text");
    let fragment = format!(
        "<{comment_tag} authorId=\"{author_id}\" idx=\"{index}\"><{position_tag} x=\"{}\" y=\"{}\"/><{text_tag}>{}</{text_tag}></{comment_tag}>",
        coordinates.x_emu,
        coordinates.y_emu,
        crate::xml_edit::escape_text(&comment.text)
    );
    package.set_part(&comments.part_name, insert_child(&part, &root, fragment)?)?;
    Ok(format!("{}/comment[{position}]", slide.path))
}

pub(super) fn set(
    package: &mut NativeOfficePackage,
    path: &str,
    update: &NativeOfficeCommentUpdate,
) -> UseResult<String> {
    let node = comment_node(package, path)?;
    let part_name = required_format(&node, "part", "Presentation comment source part")?;
    let ordinal = comment_ordinal(&node.path)?;
    if update.author.is_some() || update.initials.is_some() {
        let authors =
            related_part(package, PRESENTATION_PART, "commentAuthors")?.ok_or_else(|| {
                comment_error(
                    "use.office.comment_author_part_missing",
                    "Presentation has no legacy comment author part.",
                )
            })?;
        let current_author = node.format.get("author").cloned().unwrap_or_default();
        let current_initials = node
            .format
            .get("initials")
            .cloned()
            .unwrap_or_else(|| derive_initials(&current_author));
        let author = update.author.as_deref().unwrap_or(&current_author);
        let initials = update.initials.as_deref().unwrap_or(&current_initials);
        let author_id = ensure_author(package, &authors.part_name, author, initials)?;
        let current_author_id = node
            .format
            .get("authorId")
            .and_then(|value| value.parse::<u32>().ok());
        if current_author_id != Some(author_id) {
            let index = issue_author_index(package, &authors.part_name, author_id)?;
            let part = package.xml_part(&part_name)?;
            let root = index_xml(&part)?;
            let target = comment_at(&root, ordinal).ok_or_else(|| node_not_found(path))?;
            let updates = BTreeMap::from([
                ("authorId".to_string(), Some(author_id.to_string())),
                ("idx".to_string(), Some(index.to_string())),
            ]);
            package.set_part(
                &part_name,
                patch_start_tag_attributes(&part, target, &updates)?,
            )?;
        }
    }
    if let Some(position) = update.position {
        let part = package.xml_part(&part_name)?;
        let root = index_xml(&part)?;
        let target = comment_at(&root, ordinal).ok_or_else(|| node_not_found(path))?;
        let position_element = target.child("pos", 1).ok_or_else(|| {
            comment_error(
                "use.office.comment_position_missing",
                "Presentation comment has no position element.",
            )
        })?;
        let updates = BTreeMap::from([
            ("x".to_string(), Some(position.x_emu.to_string())),
            ("y".to_string(), Some(position.y_emu.to_string())),
        ]);
        package.set_part(
            &part_name,
            patch_start_tag_attributes(&part, position_element, &updates)?,
        )?;
    }
    if let Some(text) = &update.text {
        let part = package.xml_part(&part_name)?;
        let root = index_xml(&part)?;
        let target = comment_at(&root, ordinal).ok_or_else(|| node_not_found(path))?;
        let text_tag = qualified(prefix(&root.qualified_name), "text");
        package.set_part(
            &part_name,
            replace_namespaced_text_descendants(
                &part,
                target,
                "text",
                root.namespace.as_deref(),
                text,
                Some(format!(
                    "<{text_tag}>{}</{text_tag}>",
                    crate::xml_edit::escape_text(text)
                )),
            )?,
        )?;
    }
    Ok(node.path)
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let node = comment_node(package, path)?;
    let owner = required_format(&node, "ownerPart", "Presentation comment owner part")?;
    let related = related_part(package, &owner, "comments")?.ok_or_else(|| {
        comment_error(
            "use.office.comment_part_missing",
            "Presentation slide has no legacy comments part.",
        )
    })?;
    let ordinal = comment_ordinal(&node.path)?;
    let part = package.xml_part(&related.part_name)?;
    let root = index_xml(&part)?;
    let target = comment_at(&root, ordinal).ok_or_else(|| node_not_found(path))?;
    let count = root
        .children
        .iter()
        .filter(|child| child.local_name == "cm")
        .count();
    if count == 1 {
        remove_related_part(package, &owner, &related)
    } else {
        package.set_part(
            &related.part_name,
            apply_patches(
                &part,
                vec![XmlPatch::new(target.full_range.clone(), Vec::new())],
            )?,
        )
    }
}

pub(super) fn remove_for_slide(
    package: &mut NativeOfficePackage,
    slide_part: &str,
) -> UseResult<()> {
    if let Some(related) = related_part(package, slide_part, "comments")? {
        remove_related_part(package, slide_part, &related)?;
    }
    Ok(())
}

fn comment_node(package: &NativeOfficePackage, path: &str) -> UseResult<crate::DocumentNode> {
    let node = NativeOfficeDocument::from_package(package.clone())?.get(path, 0)?;
    if node.node_type != OfficeNodeType::Comment || !node.path.contains("/comment[") {
        return Err(comment_error(
            "use.office.comment_path_unsupported",
            "Native Presentation comment paths use /slide[N]/comment[M].",
        ));
    }
    Ok(node)
}

fn required_format(node: &crate::DocumentNode, key: &str, label: &str) -> UseResult<String> {
    node.format.get(key).cloned().ok_or_else(|| {
        comment_error(
            "use.office.comment_part_invalid",
            format!("{label} is missing."),
        )
    })
}

fn comment_ordinal(path: &str) -> UseResult<usize> {
    path.rsplit_once("/comment[")
        .and_then(|(_, ordinal)| ordinal.strip_suffix(']'))
        .and_then(|ordinal| ordinal.parse::<usize>().ok())
        .filter(|ordinal| *ordinal > 0)
        .ok_or_else(|| node_not_found(path))
}

fn ensure_author(
    package: &mut NativeOfficePackage,
    part_name: &str,
    name: &str,
    initials: &str,
) -> UseResult<u32> {
    let part = package.xml_part(part_name)?;
    let root = index_xml(&part)?;
    require_author_list(&root, part_name)?;
    for author in root
        .children
        .iter()
        .filter(|child| child.local_name == "cmAuthor")
    {
        if author.attributes.get("name").map(String::as_str) == Some(name)
            && author.attributes.get("initials").map(String::as_str) == Some(initials)
        {
            return author
                .attributes
                .get("id")
                .and_then(|id| id.parse::<u32>().ok())
                .ok_or_else(|| {
                    comment_error(
                        "use.office.comment_author_invalid",
                        "Presentation comment author has an invalid ID.",
                    )
                });
        }
    }
    let next_id = root
        .children
        .iter()
        .filter(|child| child.local_name == "cmAuthor")
        .filter_map(|author| author.attributes.get("id"))
        .filter_map(|id| id.parse::<u32>().ok())
        .max()
        .map_or(Some(0), |id| id.checked_add(1))
        .ok_or_else(|| {
            comment_error(
                "use.office.comment_author_id_exhausted",
                "Presentation comment author IDs are exhausted.",
            )
        })?;
    let tag = qualified(prefix(&root.qualified_name), "cmAuthor");
    let fragment = format!(
        "<{tag} id=\"{next_id}\" name=\"{}\" initials=\"{}\" lastIdx=\"0\" clrIdx=\"{}\"/>",
        crate::xml_edit::escape_attribute(name),
        crate::xml_edit::escape_attribute(initials),
        next_id % 8
    );
    package.set_part(part_name, insert_child(&part, &root, fragment)?)?;
    Ok(next_id)
}

fn issue_author_index(
    package: &mut NativeOfficePackage,
    authors_part: &str,
    author_id: u32,
) -> UseResult<u32> {
    let mut maximum = 0_u32;
    for part_name in package
        .part_names()
        .filter(|part| part.starts_with("ppt/comments/") && part.ends_with(".xml"))
        .map(str::to_string)
        .collect::<Vec<_>>()
    {
        let part = package.xml_part(&part_name)?;
        let root = index_xml(&part)?;
        for comment in root
            .children
            .iter()
            .filter(|child| child.local_name == "cm")
        {
            if comment
                .attributes
                .get("authorId")
                .and_then(|id| id.parse::<u32>().ok())
                == Some(author_id)
            {
                maximum = maximum.max(
                    comment
                        .attributes
                        .get("idx")
                        .and_then(|index| index.parse::<u32>().ok())
                        .unwrap_or(0),
                );
            }
        }
    }
    let part = package.xml_part(authors_part)?;
    let root = index_xml(&part)?;
    let author = find_author(&root, author_id).ok_or_else(|| {
        comment_error(
            "use.office.comment_author_invalid",
            format!("Presentation comment author {author_id} does not exist."),
        )
    })?;
    maximum = maximum.max(
        author
            .attributes
            .get("lastIdx")
            .and_then(|value| value.parse::<u32>().ok())
            .unwrap_or(0),
    );
    let next = maximum.checked_add(1).ok_or_else(|| {
        comment_error(
            "use.office.comment_index_exhausted",
            "Presentation per-author comment indexes are exhausted.",
        )
    })?;
    let updates = BTreeMap::from([("lastIdx".to_string(), Some(next.to_string()))]);
    package.set_part(
        authors_part,
        patch_start_tag_attributes(&part, author, &updates)?,
    )?;
    Ok(next)
}

fn find_author(root: &IndexedXmlElement, id: u32) -> Option<&IndexedXmlElement> {
    root.children
        .iter()
        .filter(|child| child.local_name == "cmAuthor")
        .find(|author| {
            author
                .attributes
                .get("id")
                .and_then(|value| value.parse::<u32>().ok())
                == Some(id)
        })
}

fn comment_at(root: &IndexedXmlElement, ordinal: usize) -> Option<&IndexedXmlElement> {
    if ordinal == 0 {
        return None;
    }
    root.children
        .iter()
        .filter(|child| child.local_name == "cm")
        .nth(ordinal - 1)
}

fn require_author_list(root: &IndexedXmlElement, part_name: &str) -> UseResult<()> {
    if root.local_name == "cmAuthorLst" {
        Ok(())
    } else {
        Err(comment_error(
            "use.office.comment_author_part_invalid",
            format!("Presentation comment author part '{part_name}' has an invalid root."),
        ))
    }
}

fn require_comment_list(root: &IndexedXmlElement, part_name: &str) -> UseResult<()> {
    if root.local_name == "cmLst" {
        Ok(())
    } else {
        Err(comment_error(
            "use.office.comment_part_invalid",
            format!("Presentation comment part '{part_name}' has an invalid root."),
        ))
    }
}
