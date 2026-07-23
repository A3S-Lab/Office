use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::UseResult;

use super::{
    comment_error, ensure_related_part, related_part, remove_related_part, CreateRelatedPart,
    WORD_COMMENTS_CONTENT_TYPE,
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

const DOCUMENT_PART: &str = "word/document.xml";

pub(super) fn add(
    package: &mut NativeOfficePackage,
    parent: &str,
    comment: &NativeOfficeComment,
) -> UseResult<String> {
    if comment.position.is_some() {
        return Err(comment_error(
            "use.office.comment_position_unsupported",
            "Native Word comments use paragraph or run anchors and do not accept slide coordinates.",
        ));
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let parent_node = snapshot.get(parent, 0)?;
    if !matches!(
        parent_node.node_type,
        OfficeNodeType::Paragraph | OfficeNodeType::Run
    ) || !parent_node.path.starts_with("/body/")
    {
        return Err(comment_error(
            "use.office.comment_parent_unsupported",
            "Native Word comments require a main-document paragraph or run parent.",
        ));
    }

    let office_dialect = dialect(package)?;
    let comments = ensure_related_part(
        package,
        CreateRelatedPart {
            owner: DOCUMENT_PART,
            directory: "word",
            stem: "comments",
            extension: "xml",
            content_type: WORD_COMMENTS_CONTENT_TYPE,
            relationship_name: "comments",
            xml: format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:comments xmlns:w=\"{}\"/>",
                office_dialect.word_namespace()
            ),
        },
    )?;
    let comments_part = package.xml_part(&comments.part_name)?;
    let comments_index = index_xml(&comments_part)?;
    require_comments_root(&comments_index, &comments.part_name)?;
    let id = next_comment_id(&comments_index)?;
    let position = comments_index
        .children
        .iter()
        .filter(|child| child.local_name == "comment")
        .count()
        + 1;
    let word_prefix = prefix(&comments_index.qualified_name);
    let fragment = comment_fragment(word_prefix, id, comment);
    package.set_part(
        &comments.part_name,
        insert_child(&comments_part, &comments_index, fragment)?,
    )?;
    add_anchor(package, &parent_node.path, id)?;
    Ok(format!("/comments/comment[{position}]"))
}

pub(super) fn set(
    package: &mut NativeOfficePackage,
    path: &str,
    update: &NativeOfficeCommentUpdate,
) -> UseResult<String> {
    if update.position.is_some() {
        return Err(comment_error(
            "use.office.comment_position_unsupported",
            "Native Word comment updates do not accept slide coordinates.",
        ));
    }
    let node = comment_node(package, path)?;
    let part_name = node.format.get("part").ok_or_else(|| {
        comment_error(
            "use.office.comment_part_missing",
            "Native Word comment has no source OOXML part.",
        )
    })?;
    let id = node
        .format
        .get("id")
        .ok_or_else(|| comment_error("use.office.comment_id_missing", "Word comment has no ID."))?;

    let mut part = package.xml_part(part_name)?;
    let mut root = index_xml(&part)?;
    let mut target = find_comment(&root, id).ok_or_else(|| node_not_found(path))?;
    if update.author.is_some() || update.initials.is_some() {
        let mut attributes = BTreeMap::new();
        if let Some(author) = &update.author {
            attributes.insert(
                qualified(prefix(&target.qualified_name), "author"),
                Some(author.clone()),
            );
        }
        if let Some(initials) = &update.initials {
            attributes.insert(
                qualified(prefix(&target.qualified_name), "initials"),
                Some(initials.clone()),
            );
        }
        package.set_part(
            part_name,
            patch_start_tag_attributes(&part, target, &attributes)?,
        )?;
        part = package.xml_part(part_name)?;
        root = index_xml(&part)?;
        target = find_comment(&root, id).ok_or_else(|| node_not_found(path))?;
    }
    if let Some(text) = &update.text {
        let word_prefix = prefix(&target.qualified_name);
        let paragraph = qualified(word_prefix, "p");
        let run = qualified(word_prefix, "r");
        let text_tag = qualified(word_prefix, "t");
        let insertion = format!(
            "<{paragraph}><{run}><{text_tag}>{}</{text_tag}></{run}></{paragraph}>",
            crate::xml_edit::escape_text(text)
        );
        package.set_part(
            part_name,
            replace_namespaced_text_descendants(
                &part,
                target,
                "t",
                target.namespace.as_deref(),
                text,
                Some(insertion),
            )?,
        )?;
    }
    Ok(node.path)
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let node = comment_node(package, path)?;
    let id =
        node.format.get("id").cloned().ok_or_else(|| {
            comment_error("use.office.comment_id_missing", "Word comment has no ID.")
        })?;
    let related = related_part(package, DOCUMENT_PART, "comments")?.ok_or_else(|| {
        comment_error(
            "use.office.comment_part_missing",
            "Word document has no comments relationship.",
        )
    })?;
    remove_anchors(package, &id)?;
    let part = package.xml_part(&related.part_name)?;
    let root = index_xml(&part)?;
    let target = find_comment(&root, &id).ok_or_else(|| node_not_found(path))?;
    let count = root
        .children
        .iter()
        .filter(|child| child.local_name == "comment")
        .count();
    if count == 1 {
        remove_related_part(package, DOCUMENT_PART, &related)
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

pub(super) fn remove_owned(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let part = package.xml_part(DOCUMENT_PART)?;
    let root = index_xml(&part)?;
    let target = crate::editor::word::locate_word_path(&root, path)?;
    let mut ids = Vec::new();
    collect_owned_comment_ids(target, &mut ids);
    collect_adjacent_comment_ids(&root, target, &mut ids);
    ids.sort();
    ids.dedup();
    if ids.is_empty() {
        return Ok(());
    }
    let owned_ids = ids.into_iter().collect::<BTreeSet<_>>();
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let paths = snapshot
        .query("comment")?
        .into_iter()
        .filter(|comment| {
            comment
                .format
                .get("id")
                .is_some_and(|id| owned_ids.contains(id))
        })
        .map(|comment| comment.path)
        .collect::<Vec<_>>();
    for path in paths.into_iter().rev() {
        remove(package, &path)?;
    }
    Ok(())
}

fn comment_node(package: &NativeOfficePackage, path: &str) -> UseResult<crate::DocumentNode> {
    let node = NativeOfficeDocument::from_package(package.clone())?.get(path, 0)?;
    if node.node_type != OfficeNodeType::Comment || !node.path.starts_with("/comments/comment[") {
        return Err(comment_error(
            "use.office.comment_path_unsupported",
            "Native Word comment paths use /comments/comment[N].",
        ));
    }
    Ok(node)
}

fn add_anchor(package: &mut NativeOfficePackage, parent: &str, id: u32) -> UseResult<()> {
    let part = package.xml_part(DOCUMENT_PART)?;
    let index = index_xml(&part)?;
    let target = crate::editor::word::locate_word_path(&index, parent)?;
    let word_prefix = prefix(&target.qualified_name);
    let range_start = qualified(word_prefix, "commentRangeStart");
    let range_end = qualified(word_prefix, "commentRangeEnd");
    let run = qualified(word_prefix, "r");
    let reference = qualified(word_prefix, "commentReference");
    let id_attribute = qualified(word_prefix, "id");
    let start = format!("<{range_start} {id_attribute}=\"{id}\"/>");
    let end = format!(
        "<{range_end} {id_attribute}=\"{id}\"/><{run}><{reference} {id_attribute}=\"{id}\"/></{run}>"
    );
    let (start_position, end_position) = if target.local_name == "p" {
        let start = target
            .child("pPr", 1)
            .map_or(target.content_range.start, |properties| {
                properties.full_range.end
            });
        (start, target.content_range.end)
    } else if target.local_name == "r" {
        (target.full_range.start, target.full_range.end)
    } else {
        return Err(comment_error(
            "use.office.comment_parent_unsupported",
            "Native Word comments require a paragraph or run parent.",
        ));
    };
    package.set_part(
        DOCUMENT_PART,
        apply_patches(
            &part,
            vec![
                XmlPatch::new(start_position..start_position, start),
                XmlPatch::new(end_position..end_position, end),
            ],
        )?,
    )
}

fn remove_anchors(package: &mut NativeOfficePackage, id: &str) -> UseResult<()> {
    let part = package.xml_part(DOCUMENT_PART)?;
    let root = index_xml(&part)?;
    let mut patches = Vec::new();
    collect_anchor_patches(&root, id, &mut patches);
    if patches.is_empty() {
        return Ok(());
    }
    package.set_part(DOCUMENT_PART, apply_patches(&part, patches)?)
}

fn collect_anchor_patches(element: &IndexedXmlElement, id: &str, output: &mut Vec<XmlPatch>) {
    for child in &element.children {
        if child.local_name == "r" && reference_only_run(child, id) {
            output.push(XmlPatch::new(child.full_range.clone(), Vec::new()));
            continue;
        }
        if matches!(
            child.local_name.as_str(),
            "commentRangeStart" | "commentRangeEnd" | "commentReference"
        ) && child.attributes.get("id").map(String::as_str) == Some(id)
        {
            output.push(XmlPatch::new(child.full_range.clone(), Vec::new()));
            continue;
        }
        collect_anchor_patches(child, id, output);
    }
}

fn collect_owned_comment_ids(element: &IndexedXmlElement, output: &mut Vec<String>) {
    if matches!(
        element.local_name.as_str(),
        "commentRangeStart" | "commentRangeEnd" | "commentReference"
    ) {
        if let Some(id) = element.attributes.get("id") {
            output.push(id.clone());
        }
    }
    for child in &element.children {
        collect_owned_comment_ids(child, output);
    }
}

fn collect_adjacent_comment_ids(
    root: &IndexedXmlElement,
    target: &IndexedXmlElement,
    output: &mut Vec<String>,
) {
    let Some(parent) = find_parent(root, target) else {
        return;
    };
    let Some(target_index) = parent
        .children
        .iter()
        .position(|child| std::ptr::eq(child, target))
    else {
        return;
    };

    let starts = parent.children[..target_index]
        .iter()
        .rev()
        .take_while(|sibling| sibling.local_name == "commentRangeStart")
        .filter_map(|marker| marker.attributes.get("id").cloned())
        .collect::<BTreeSet<_>>();
    let mut ends = BTreeSet::new();
    for sibling in &parent.children[target_index + 1..] {
        if sibling.local_name == "commentRangeEnd" {
            if let Some(id) = sibling.attributes.get("id") {
                ends.insert(id.clone());
            }
        } else if sibling.local_name != "r" || comment_reference_id(sibling).is_none() {
            break;
        }
    }
    for id in starts.intersection(&ends) {
        output.push(id.clone());
    }
}

fn find_parent<'a>(
    element: &'a IndexedXmlElement,
    target: &IndexedXmlElement,
) -> Option<&'a IndexedXmlElement> {
    if element
        .children
        .iter()
        .any(|child| std::ptr::eq(child, target))
    {
        return Some(element);
    }
    element
        .children
        .iter()
        .find_map(|child| find_parent(child, target))
}

fn reference_only_run(run: &IndexedXmlElement, id: &str) -> bool {
    comment_reference_id(run) == Some(id)
}

fn comment_reference_id(run: &IndexedXmlElement) -> Option<&str> {
    let meaningful = run
        .children
        .iter()
        .filter(|child| child.local_name != "rPr")
        .collect::<Vec<_>>();
    (meaningful.len() == 1 && meaningful[0].local_name == "commentReference")
        .then(|| meaningful[0].attributes.get("id").map(String::as_str))
        .flatten()
}

fn require_comments_root(root: &IndexedXmlElement, part_name: &str) -> UseResult<()> {
    if root.local_name == "comments" {
        Ok(())
    } else {
        Err(comment_error(
            "use.office.comment_part_invalid",
            format!("Word comment part '{part_name}' has an invalid root."),
        ))
    }
}

fn next_comment_id(root: &IndexedXmlElement) -> UseResult<u32> {
    let maximum = root
        .children
        .iter()
        .filter(|child| child.local_name == "comment")
        .filter_map(|comment| comment.attributes.get("id"))
        .filter_map(|id| id.parse::<u32>().ok())
        .max();
    maximum.map_or(Ok(0), |id| {
        id.checked_add(1).ok_or_else(|| {
            comment_error(
                "use.office.comment_id_exhausted",
                "Word comment IDs are exhausted.",
            )
        })
    })
}

fn find_comment<'a>(root: &'a IndexedXmlElement, id: &str) -> Option<&'a IndexedXmlElement> {
    root.children
        .iter()
        .filter(|child| child.local_name == "comment")
        .find(|comment| comment.attributes.get("id").map(String::as_str) == Some(id))
}

fn comment_fragment(prefix: Option<&str>, id: u32, comment: &NativeOfficeComment) -> String {
    let comment_tag = qualified(prefix, "comment");
    let paragraph = qualified(prefix, "p");
    let run = qualified(prefix, "r");
    let text = qualified(prefix, "t");
    let id_attribute = qualified(prefix, "id");
    let author_attribute = qualified(prefix, "author");
    let initials = comment.initials.as_ref().map_or_else(String::new, |value| {
        format!(
            " {}=\"{}\"",
            qualified(prefix, "initials"),
            crate::xml_edit::escape_attribute(value)
        )
    });
    format!(
        "<{comment_tag} {id_attribute}=\"{id}\" {author_attribute}=\"{}\"{initials}><{paragraph}><{run}><{text}>{}</{text}></{run}></{paragraph}></{comment_tag}>",
        crate::xml_edit::escape_attribute(&comment.author),
        crate::xml_edit::escape_text(&comment.text)
    )
}
