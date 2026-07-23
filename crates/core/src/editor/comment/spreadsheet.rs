use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{
    comment_error, ensure_namespace, ensure_related_part, related_part, remove_related_part,
    CreateRelatedPart, SPREADSHEET_COMMENTS_CONTENT_TYPE, VML_CONTENT_TYPE,
};
use crate::editor::part::dialect;
use crate::editor::{
    node_not_found, prefix, qualified, NativeOfficeComment, NativeOfficeCommentUpdate,
};
use crate::spreadsheet_reference::{CellRange, CellReference, MAX_COLUMNS, MAX_ROWS};
use crate::xml_edit::{
    apply_patches, decoded_element_text, index_xml, insert_child, insert_ordered_child,
    patch_start_tag_attributes, replace_namespaced_text_descendants, IndexedXmlElement, XmlPatch,
};
use crate::{LosslessXmlPart, NativeOfficeDocument, NativeOfficePackage, OfficeNodeType};

const VML_NAMESPACE: &str = "urn:schemas-microsoft-com:vml";
const OFFICE_NAMESPACE: &str = "urn:schemas-microsoft-com:office:office";
const EXCEL_NAMESPACE: &str = "urn:schemas-microsoft-com:office:excel";
const WORKSHEET_COMMENT_LATER_SIBLINGS: &[&str] = &[
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
];

pub(super) fn add(
    package: &mut NativeOfficePackage,
    parent: &str,
    comment: &NativeOfficeComment,
) -> UseResult<String> {
    reject_unsupported_fields(comment.initials.as_ref(), comment.position.as_ref())?;
    let (sheet_path, reference, owner) = resolve_cell_parent(package, parent)?;
    let office_dialect = dialect(package)?;
    let related = ensure_related_part(
        package,
        CreateRelatedPart {
            owner: &owner,
            directory: "xl",
            stem: "comments",
            extension: "xml",
            content_type: SPREADSHEET_COMMENTS_CONTENT_TYPE,
            relationship_name: "comments",
            xml: format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><comments xmlns=\"{}\"><authors/><commentList/></comments>",
                office_dialect.spreadsheet_namespace()
            ),
        },
    )?;
    let author_id = ensure_author(package, &related.part_name, &comment.author)?;
    let part = package.xml_part(&related.part_name)?;
    let root = index_xml(&part)?;
    require_comments_root(&root, &related.part_name)?;
    let list = root.child("commentList", 1).ok_or_else(|| {
        comment_error(
            "use.office.comment_part_invalid",
            format!(
                "Spreadsheet comment part '{}' has no commentList.",
                related.part_name
            ),
        )
    })?;
    if find_comment(list, &reference).is_some() {
        return Err(comment_error(
            "use.office.comment_exists",
            format!("Spreadsheet cell '{sheet_path}/{reference}' already has a classic comment."),
        ));
    }
    let spreadsheet_prefix = prefix(&root.qualified_name);
    let comment_tag = qualified(spreadsheet_prefix, "comment");
    let text_tag = qualified(spreadsheet_prefix, "text");
    let value_tag = qualified(spreadsheet_prefix, "t");
    let fragment = format!(
        "<{comment_tag} ref=\"{reference}\" authorId=\"{author_id}\"><{text_tag}><{value_tag}>{}</{value_tag}></{text_tag}></{comment_tag}>",
        crate::xml_edit::escape_text(&comment.text)
    );
    package.set_part(&related.part_name, insert_child(&part, list, fragment)?)?;
    ensure_vml_note(package, &owner, CellReference::parse(&reference)?)?;
    Ok(format!("{sheet_path}/{reference}/comment"))
}

pub(super) fn set(
    package: &mut NativeOfficePackage,
    path: &str,
    update: &NativeOfficeCommentUpdate,
) -> UseResult<String> {
    reject_unsupported_fields(update.initials.as_ref(), update.position.as_ref())?;
    let node = comment_node(package, path)?;
    let part_name = required_format(&node, "part", "Spreadsheet comment source part")?;
    let reference = required_format(&node, "ref", "Spreadsheet comment cell reference")?;
    if let Some(author) = &update.author {
        let author_id = ensure_author(package, &part_name, author)?;
        let part = package.xml_part(&part_name)?;
        let root = index_xml(&part)?;
        let target = root
            .child("commentList", 1)
            .and_then(|list| find_comment(list, &reference))
            .ok_or_else(|| node_not_found(path))?;
        let updates = BTreeMap::from([("authorId".to_string(), Some(author_id.to_string()))]);
        package.set_part(
            &part_name,
            patch_start_tag_attributes(&part, target, &updates)?,
        )?;
    }
    if let Some(text) = &update.text {
        let part = package.xml_part(&part_name)?;
        let root = index_xml(&part)?;
        let target = root
            .child("commentList", 1)
            .and_then(|list| find_comment(list, &reference))
            .ok_or_else(|| node_not_found(path))?;
        let spreadsheet_prefix = prefix(&root.qualified_name);
        let text_tag = qualified(spreadsheet_prefix, "text");
        let value_tag = qualified(spreadsheet_prefix, "t");
        let insertion = format!(
            "<{text_tag}><{value_tag}>{}</{value_tag}></{text_tag}>",
            crate::xml_edit::escape_text(text)
        );
        package.set_part(
            &part_name,
            replace_namespaced_text_descendants(
                &part,
                target,
                "t",
                root.namespace.as_deref(),
                text,
                Some(insertion),
            )?,
        )?;
    }
    Ok(node.path)
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let node = comment_node(package, path)?;
    let owner = required_format(&node, "ownerPart", "Spreadsheet comment owner part")?;
    let reference = required_format(&node, "ref", "Spreadsheet comment cell reference")?;
    let related = related_part(package, &owner, "comments")?.ok_or_else(|| {
        comment_error(
            "use.office.comment_part_missing",
            "Spreadsheet worksheet has no comments relationship.",
        )
    })?;
    let part = package.xml_part(&related.part_name)?;
    let root = index_xml(&part)?;
    let list = root
        .child("commentList", 1)
        .ok_or_else(|| node_not_found(path))?;
    let target = find_comment(list, &reference).ok_or_else(|| node_not_found(path))?;
    let count = list
        .children
        .iter()
        .filter(|child| child.local_name == "comment")
        .count();
    if count == 1 {
        remove_related_part(package, &owner, &related)?;
    } else {
        package.set_part(
            &related.part_name,
            apply_patches(
                &part,
                vec![XmlPatch::new(target.full_range.clone(), Vec::new())],
            )?,
        )?;
    }
    remove_vml_note(package, &owner, CellReference::parse(&reference)?)
}

pub(super) fn remove_for_range(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let (sheet_path, requested) = path.rsplit_once('/').ok_or_else(|| node_not_found(path))?;
    let range = CellRange::parse(requested)?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let paths = snapshot
        .query("comment")?
        .into_iter()
        .filter(|comment| comment.path.starts_with(&format!("{sheet_path}/")))
        .filter(|comment| {
            comment
                .format
                .get("ref")
                .and_then(|reference| CellReference::parse(reference).ok())
                .is_some_and(|reference| range.contains(reference))
        })
        .map(|comment| comment.path)
        .collect::<Vec<_>>();
    for comment_path in paths {
        remove(package, &comment_path)?;
    }
    Ok(())
}

fn reject_unsupported_fields(
    initials: Option<&String>,
    position: Option<&crate::NativeOfficeCommentPosition>,
) -> UseResult<()> {
    if initials.is_some() {
        return Err(comment_error(
            "use.office.comment_initials_unsupported",
            "Classic Spreadsheet comments store an author name but not separate initials.",
        ));
    }
    if position.is_some() {
        return Err(comment_error(
            "use.office.comment_position_unsupported",
            "Classic Spreadsheet comment positions are derived from their cell and do not accept slide coordinates.",
        ));
    }
    Ok(())
}

fn resolve_cell_parent(
    package: &NativeOfficePackage,
    parent: &str,
) -> UseResult<(String, String, String)> {
    let (requested_sheet, requested_reference) = parent
        .strip_prefix('/')
        .and_then(|path| path.split_once('/'))
        .ok_or_else(|| {
            comment_error(
                "use.office.comment_parent_unsupported",
                "Native Spreadsheet comments require a single cell parent such as /Sheet1/B2.",
            )
        })?;
    if requested_sheet.is_empty()
        || requested_reference.is_empty()
        || requested_reference.contains('/')
    {
        return Err(comment_error(
            "use.office.comment_parent_unsupported",
            "Native Spreadsheet comments require a single cell parent such as /Sheet1/B2.",
        ));
    }
    let reference = CellReference::parse(requested_reference)?.a1();
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let sheet = snapshot
        .root()
        .children
        .iter()
        .find(|node| {
            node.node_type == OfficeNodeType::Worksheet
                && node.path[1..].eq_ignore_ascii_case(requested_sheet)
        })
        .ok_or_else(|| node_not_found(parent))?;
    let owner = sheet.format.get("part").cloned().ok_or_else(|| {
        comment_error(
            "use.office.spreadsheet_sheet_invalid",
            "Spreadsheet semantic sheet has no source part.",
        )
    })?;
    Ok((sheet.path.clone(), reference, owner))
}

fn comment_node(package: &NativeOfficePackage, path: &str) -> UseResult<crate::DocumentNode> {
    let node = NativeOfficeDocument::from_package(package.clone())?.get(path, 0)?;
    if node.node_type != OfficeNodeType::Comment {
        return Err(comment_error(
            "use.office.comment_path_unsupported",
            "Native Spreadsheet comment paths use /SheetName/CellRef/comment.",
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

fn ensure_author(
    package: &mut NativeOfficePackage,
    part_name: &str,
    requested: &str,
) -> UseResult<usize> {
    let part = package.xml_part(part_name)?;
    let root = index_xml(&part)?;
    require_comments_root(&root, part_name)?;
    let authors = root.child("authors", 1).ok_or_else(|| {
        comment_error(
            "use.office.comment_part_invalid",
            format!("Spreadsheet comment part '{part_name}' has no authors collection."),
        )
    })?;
    for (index, author) in authors
        .children
        .iter()
        .filter(|child| child.local_name == "author")
        .enumerate()
    {
        if decoded_element_text(&part, author)? == requested {
            return Ok(index);
        }
    }
    let author_id = authors
        .children
        .iter()
        .filter(|child| child.local_name == "author")
        .count();
    let tag = qualified(prefix(&root.qualified_name), "author");
    package.set_part(
        part_name,
        insert_child(
            &part,
            authors,
            format!("<{tag}>{}</{tag}>", crate::xml_edit::escape_text(requested)),
        )?,
    )?;
    Ok(author_id)
}

fn ensure_vml_note(
    package: &mut NativeOfficePackage,
    owner: &str,
    reference: CellReference,
) -> UseResult<()> {
    let root_xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><xml xmlns:v=\"{VML_NAMESPACE}\" xmlns:o=\"{OFFICE_NAMESPACE}\" xmlns:x=\"{EXCEL_NAMESPACE}\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype></xml>"
    );
    let related = ensure_related_part(
        package,
        CreateRelatedPart {
            owner,
            directory: "xl/drawings",
            stem: "vmlDrawing",
            extension: "vml",
            content_type: VML_CONTENT_TYPE,
            relationship_name: "vmlDrawing",
            xml: root_xml,
        },
    )?;
    ensure_legacy_drawing_reference(package, owner, &related)?;

    let original = package.xml_part(&related.part_name)?;
    let (bytes, vml_prefix) = ensure_namespace(&original, "v", VML_NAMESPACE)?;
    let namespaced = LosslessXmlPart::parse(related.part_name.clone(), bytes)?;
    let (bytes, office_prefix) = ensure_namespace(&namespaced, "o", OFFICE_NAMESPACE)?;
    let namespaced = LosslessXmlPart::parse(related.part_name.clone(), bytes)?;
    let (bytes, excel_prefix) = ensure_namespace(&namespaced, "x", EXCEL_NAMESPACE)?;
    let part = LosslessXmlPart::parse(related.part_name.clone(), bytes)?;
    let root = index_xml(&part)?;
    let shape_id = next_shape_id(&root)?;
    let column = reference.column - 1;
    let row = reference.row - 1;
    let end_column = reference.column.saturating_add(3).min(MAX_COLUMNS) - 1;
    let end_row = reference.row.saturating_add(4).min(MAX_ROWS) - 1;
    let v = |name| qualified(Some(&vml_prefix), name);
    let o = |name| qualified(Some(&office_prefix), name);
    let x = |name| qualified(Some(&excel_prefix), name);
    let fragment = format!(
        "<{} id=\"_x0000_s{shape_id}\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden\" fillcolor=\"#ffffe1\" {}=\"auto\"><{} color=\"black\" opacity=\"0.5\"/><{} style=\"mso-direction-alt:auto\"/><{} ObjectType=\"Note\"><{}/><{}/><{}>{column}, 15, {row}, 2, {end_column}, 15, {end_row}, 2</{}><{}>False</{}><{}>{row}</{}><{}>{column}</{}></{}></{}>",
        v("shape"),
        o("insetmode"),
        v("shadow"),
        v("textbox"),
        x("ClientData"),
        x("MoveWithCells"),
        x("SizeWithCells"),
        x("Anchor"),
        x("Anchor"),
        x("AutoFill"),
        x("AutoFill"),
        x("Row"),
        x("Row"),
        x("Column"),
        x("Column"),
        x("ClientData"),
        v("shape")
    );
    package.set_part(&related.part_name, insert_child(&part, &root, fragment)?)
}

fn ensure_legacy_drawing_reference(
    package: &mut NativeOfficePackage,
    owner: &str,
    related: &super::RelatedPart,
) -> UseResult<()> {
    let original = package.xml_part(owner)?;
    let relationship_namespace = dialect(package)?.relationship_namespace();
    let (bytes, relationship_prefix) = ensure_namespace(&original, "r", relationship_namespace)?;
    let part = LosslessXmlPart::parse(owner.to_string(), bytes)?;
    let root = index_xml(&part)?;
    if let Some(existing) = root.child("legacyDrawing", 1) {
        let existing_id = crate::editor::hyperlink::old_relationship_id(existing);
        if existing_id.as_deref() == Some(&related.relationship_id) {
            package.set_part(owner, part.raw().to_vec())?;
            return Ok(());
        }
        return Err(comment_error(
            "use.office.spreadsheet_vml_conflict",
            "Worksheet already references a different legacy VML drawing.",
        ));
    }
    let tag = qualified(prefix(&root.qualified_name), "legacyDrawing");
    let id = qualified(Some(&relationship_prefix), "id");
    package.set_part(
        owner,
        insert_ordered_child(
            &part,
            &root,
            format!(
                "<{tag} {id}=\"{}\"/>",
                crate::xml_edit::escape_attribute(&related.relationship_id)
            ),
            WORKSHEET_COMMENT_LATER_SIBLINGS,
        )?,
    )
}

fn remove_vml_note(
    package: &mut NativeOfficePackage,
    owner: &str,
    reference: CellReference,
) -> UseResult<()> {
    let Some(related) = related_part(package, owner, "vmlDrawing")? else {
        return Ok(());
    };
    let part = package.xml_part(&related.part_name)?;
    let root = index_xml(&part)?;
    let mut shapes = Vec::new();
    root.descendants_named("shape", &mut shapes);
    let target = shapes
        .iter()
        .copied()
        .find(|shape| vml_shape_reference(&part, shape).ok().flatten() == Some(reference));
    let Some(target) = target else {
        return Ok(());
    };
    let edited = apply_patches(
        &part,
        vec![XmlPatch::new(target.full_range.clone(), Vec::new())],
    )?;
    package.set_part(&related.part_name, edited)?;
    let remaining = package.xml_part(&related.part_name)?;
    let remaining_root = index_xml(&remaining)?;
    let mut remaining_shapes = Vec::new();
    remaining_root.descendants_named("shape", &mut remaining_shapes);
    if !remaining_shapes.is_empty() {
        return Ok(());
    }
    remove_legacy_drawing_reference(package, owner, &related.relationship_id)?;
    remove_related_part(package, owner, &related)
}

fn remove_legacy_drawing_reference(
    package: &mut NativeOfficePackage,
    owner: &str,
    relationship_id: &str,
) -> UseResult<()> {
    let part = package.xml_part(owner)?;
    let root = index_xml(&part)?;
    let Some(legacy) = root.child("legacyDrawing", 1) else {
        return Ok(());
    };
    if crate::editor::hyperlink::old_relationship_id(legacy).as_deref() != Some(relationship_id) {
        return Ok(());
    }
    package.set_part(
        owner,
        apply_patches(
            &part,
            vec![XmlPatch::new(legacy.full_range.clone(), Vec::new())],
        )?,
    )
}

fn vml_shape_reference(
    part: &LosslessXmlPart,
    shape: &IndexedXmlElement,
) -> UseResult<Option<CellReference>> {
    let Some(client_data) = shape.descendant("ClientData") else {
        return Ok(None);
    };
    if client_data.attributes.get("ObjectType").map(String::as_str) != Some("Note") {
        return Ok(None);
    }
    let Some(row) = client_data.child("Row", 1) else {
        return Ok(None);
    };
    let Some(column) = client_data.child("Column", 1) else {
        return Ok(None);
    };
    let row = decoded_element_text(part, row)?
        .parse::<u32>()
        .ok()
        .and_then(|value| value.checked_add(1));
    let column = decoded_element_text(part, column)?
        .parse::<u32>()
        .ok()
        .and_then(|value| value.checked_add(1));
    Ok(match (column, row) {
        (Some(column), Some(row)) if column <= MAX_COLUMNS && row <= MAX_ROWS => {
            Some(CellReference { column, row })
        }
        _ => None,
    })
}

fn next_shape_id(root: &IndexedXmlElement) -> UseResult<u32> {
    let mut shapes = Vec::new();
    root.descendants_named("shape", &mut shapes);
    let maximum = shapes
        .into_iter()
        .filter_map(|shape| shape.attributes.get("id"))
        .filter_map(|id| id.strip_prefix("_x0000_s"))
        .filter_map(|id| id.parse::<u32>().ok())
        .max()
        .unwrap_or(1024);
    maximum.checked_add(1).ok_or_else(|| {
        comment_error(
            "use.office.spreadsheet_vml_id_exhausted",
            "Spreadsheet VML shape IDs are exhausted.",
        )
    })
}

fn require_comments_root(root: &IndexedXmlElement, part_name: &str) -> UseResult<()> {
    if root.local_name == "comments" {
        Ok(())
    } else {
        Err(comment_error(
            "use.office.comment_part_invalid",
            format!("Spreadsheet comment part '{part_name}' has an invalid root."),
        ))
    }
}

fn find_comment<'a>(list: &'a IndexedXmlElement, reference: &str) -> Option<&'a IndexedXmlElement> {
    list.children
        .iter()
        .filter(|child| child.local_name == "comment")
        .find(|comment| {
            comment
                .attributes
                .get("ref")
                .and_then(|value| CellReference::parse(value).ok())
                .is_some_and(|value| value.a1() == reference)
        })
}
