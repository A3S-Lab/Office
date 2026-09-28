use std::collections::HashMap;
use std::sync::Arc;

use a3s_use_core::UseResult;
use yrs::types::xml::XmlIn;
use yrs::{
    GetString, ReadTxn, Transact, Xml, XmlElementPrelim, XmlElementRef, XmlFragment, XmlOut,
    XmlTextPrelim,
};

use super::super::super::{
    collaboration_error, NativeOfficeCollaborationManifest,
    NativeOfficeCollaborationParagraphPosition,
};
use super::super::utf16_len;
use super::identity::{
    ancestor_table_rows, document_identity_attribute, table_row_text_id_rotations,
    validate_paragraph_id_input, PARAGRAPH_ID_ATTRIBUTE, ROW_TEXT_ID_ATTRIBUTE, TEXT_ID_ATTRIBUTE,
};

const MAX_DOCUMENT_PARAGRAPH_TEXT_UTF16: u32 = 1_048_576;
const MAX_DOCUMENT_ANCESTOR_DEPTH: usize = 64;

pub(super) fn validate_insert_list_item(
    anchor_paragraph_id: &str,
    expected_text_id: &str,
    paragraph_id: &str,
    text_id: &str,
    text: &str,
) -> UseResult<()> {
    validate_paragraph_id_input(anchor_paragraph_id, "anchorParagraphId")?;
    validate_paragraph_id_input(expected_text_id, "expectedTextId")?;
    validate_paragraph_id_input(paragraph_id, "paragraphId")?;
    validate_paragraph_id_input(text_id, "textId")?;
    if anchor_paragraph_id == paragraph_id {
        return Err(invalid_list_item(
            "A new Document list item paragraph ID must differ from its anchor paragraph ID.",
        ));
    }
    if utf16_len(text)? > MAX_DOCUMENT_PARAGRAPH_TEXT_UTF16 {
        return Err(collaboration_error(
            "office.collaboration.mutation_too_large",
            "The new Document list item text exceeds the supported UTF-16 length.",
        ));
    }
    Ok(())
}

pub(super) fn validate_delete_list_item(
    paragraph_id: &str,
    expected_text_id: &str,
    expected_text: &str,
) -> UseResult<()> {
    validate_paragraph_id_input(paragraph_id, "paragraphId")?;
    validate_paragraph_id_input(expected_text_id, "expectedTextId")?;
    if utf16_len(expected_text)? > MAX_DOCUMENT_PARAGRAPH_TEXT_UTF16 {
        return Err(collaboration_error(
            "office.collaboration.mutation_too_large",
            "The observed Document list item text exceeds the supported UTF-16 length.",
        ));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
pub(super) fn insert_list_item(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    anchor_paragraph_id: &str,
    expected_text_id: &str,
    position: NativeOfficeCollaborationParagraphPosition,
    paragraph_id: &str,
    text_id: &str,
    text: &str,
) -> UseResult<()> {
    let root = format!("{}.document.content", manifest.namespace);
    let fragment = doc.get_or_insert_xml_fragment(root);
    let transaction = doc.transact();
    let anchor = paragraph(&fragment, &transaction, anchor_paragraph_id)?;
    let current_text_id = required_text_id(&anchor, &transaction)?;
    if current_text_id != expected_text_id {
        return Err(match_conflict(
            "The target Document list item identity changed after it was inspected.",
        ));
    }
    let item = leading_list_item(&anchor, &transaction)?;
    let list = parent_list(&item)?;
    assert_block_reaches_section(&fragment, &transaction, &list)?;
    ensure_paragraph_available(&fragment, &transaction, paragraph_id)?;
    let rotations = table_row_text_id_rotations(
        &ancestor_table_rows(std::slice::from_ref(&item))?,
        &transaction,
    )?;
    let item_index = child_index(&list, &transaction, &item)?;
    let insert_index = match position {
        NativeOfficeCollaborationParagraphPosition::Before => item_index,
        NativeOfficeCollaborationParagraphPosition::After => {
            item_index.checked_add(1).ok_or_else(list_too_large)?
        }
    };
    drop(transaction);

    let created = new_list_item(paragraph_id, text_id, text);
    let mut transaction = doc.transact_mut();
    list.insert(&mut transaction, insert_index, created);
    for (ancestor, next_text_id) in rotations {
        ancestor.insert_attribute(&mut transaction, ROW_TEXT_ID_ATTRIBUTE, next_text_id);
    }
    Ok(())
}

pub(super) fn delete_list_item(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    paragraph_id: &str,
    expected_text_id: &str,
    expected_text: &str,
) -> UseResult<()> {
    let root = format!("{}.document.content", manifest.namespace);
    let fragment = doc.get_or_insert_xml_fragment(root);
    let transaction = doc.transact();
    let anchor = paragraph(&fragment, &transaction, paragraph_id)?;
    let current_text_id = required_text_id(&anchor, &transaction)?;
    if current_text_id != expected_text_id {
        return Err(match_conflict(
            "The target Document list item identity changed after it was inspected.",
        ));
    }
    let item = leading_list_item(&anchor, &transaction)?;
    let list = parent_list(&item)?;
    assert_block_reaches_section(&fragment, &transaction, &list)?;
    if list_item_count(&list, &transaction)? <= 1 {
        return Err(structure_conflict(
            "A Document list must retain at least one item.",
        ));
    }
    prove_single_paragraph_item(&item, &transaction, paragraph_id, expected_text)?;
    let rotations = table_row_text_id_rotations(
        &ancestor_table_rows(std::slice::from_ref(&item))?,
        &transaction,
    )?;
    let index = child_index(&list, &transaction, &item)?;
    drop(transaction);

    let mut transaction = doc.transact_mut();
    list.remove(&mut transaction, index);
    for (ancestor, next_text_id) in rotations {
        ancestor.insert_attribute(&mut transaction, ROW_TEXT_ID_ATTRIBUTE, next_text_id);
    }
    Ok(())
}

fn new_list_item(paragraph_id: &str, text_id: &str, text: &str) -> XmlElementPrelim {
    let mut attributes = HashMap::new();
    attributes.insert(Arc::from(PARAGRAPH_ID_ATTRIBUTE), paragraph_id.to_owned());
    attributes.insert(Arc::from(TEXT_ID_ATTRIBUTE), text_id.to_owned());
    let children = if text.is_empty() {
        Vec::new()
    } else {
        vec![XmlIn::from(XmlTextPrelim::new(text))]
    };
    let paragraph = XmlElementPrelim {
        tag: Arc::from("paragraph"),
        attributes,
        children,
    };
    XmlElementPrelim {
        tag: Arc::from("listItem"),
        attributes: HashMap::new(),
        children: vec![XmlIn::from(paragraph)],
    }
}

fn paragraph<T: ReadTxn>(
    fragment: &impl XmlFragment,
    transaction: &T,
    paragraph_id: &str,
) -> UseResult<XmlElementRef> {
    let mut matches = Vec::new();
    for node in fragment.successors(transaction) {
        let XmlOut::Element(element) = node else {
            continue;
        };
        if document_identity_attribute(&element, transaction, PARAGRAPH_ID_ATTRIBUTE)?.as_deref()
            == Some(paragraph_id)
        {
            matches.push(element);
        }
    }
    match matches.len() {
        0 => Err(collaboration_error(
            "office.collaboration.mutation_target_missing",
            format!("Document paragraph '{paragraph_id}' does not exist."),
        )),
        1 => Ok(matches.pop().expect("length checked")),
        count => Err(collaboration_error(
            "office.collaboration.mutation_identity_conflict",
            format!("Document paragraph ID '{paragraph_id}' is assigned to {count} live nodes."),
        )),
    }
}

fn required_text_id<T: ReadTxn>(element: &XmlElementRef, transaction: &T) -> UseResult<String> {
    document_identity_attribute(element, transaction, TEXT_ID_ATTRIBUTE)?.ok_or_else(|| {
        collaboration_error(
            "office.collaboration.content_invalid",
            "The target Document list paragraph does not have a valid Word text identity.",
        )
    })
}

fn leading_list_item<T: ReadTxn>(
    anchor: &XmlElementRef,
    transaction: &T,
) -> UseResult<XmlElementRef> {
    if anchor.tag().as_ref() != "paragraph" {
        return Err(structure_conflict(
            "A Document list item must begin with a plain paragraph.",
        ));
    }
    let item = match anchor.parent() {
        Some(XmlOut::Element(parent)) if parent.tag().as_ref() == "listItem" => parent,
        _ => {
            return Err(structure_conflict(
                "The anchor paragraph is not inside a Document list item.",
            ))
        }
    };
    let first = item.children(transaction).next();
    let Some(XmlOut::Element(first)) = first else {
        return Err(structure_conflict(
            "A Document list item must begin with a plain paragraph.",
        ));
    };
    if first != *anchor {
        return Err(structure_conflict(
            "The anchor paragraph is not the first block of its Document list item.",
        ));
    }
    Ok(item)
}

fn parent_list(item: &XmlElementRef) -> UseResult<XmlElementRef> {
    match item.parent() {
        Some(XmlOut::Element(parent))
            if matches!(parent.tag().as_ref(), "bulletList" | "orderedList") =>
        {
            Ok(parent)
        }
        _ => Err(structure_conflict(
            "A Document list item must be a direct child of a bullet or ordered list.",
        )),
    }
}

fn prove_single_paragraph_item<T: ReadTxn>(
    item: &XmlElementRef,
    transaction: &T,
    paragraph_id: &str,
    expected_text: &str,
) -> UseResult<()> {
    let mut blocks = Vec::new();
    for child in item.children(transaction) {
        let XmlOut::Element(element) = child else {
            return Err(structure_conflict(
                "A Document list item contains a child that is not a block.",
            ));
        };
        blocks.push(element);
    }
    if blocks.len() != 1 {
        return Err(structure_conflict(
            "A Document list item can be deleted only when it holds one plain paragraph.",
        ));
    }
    let paragraph = &blocks[0];
    if paragraph.tag().as_ref() != "paragraph" {
        return Err(structure_conflict(
            "A Document list item can be deleted only when it holds one plain paragraph.",
        ));
    }
    let current_id = document_identity_attribute(paragraph, transaction, PARAGRAPH_ID_ATTRIBUTE)?
        .ok_or_else(|| {
        structure_conflict(
            "The Document list item paragraph does not have a valid Word paragraph identity.",
        )
    })?;
    if current_id != paragraph_id {
        return Err(match_conflict(
            "The target Document list item paragraph changed after it was inspected.",
        ));
    }
    let mut text = String::new();
    for child in paragraph.children(transaction) {
        let XmlOut::Text(child) = child else {
            return Err(structure_conflict(
                "A Document list item paragraph containing inline atoms cannot be deleted by this frame.",
            ));
        };
        text.push_str(&child.get_string(transaction));
    }
    if text != expected_text {
        return Err(match_conflict(
            "The target Document list item text changed after it was inspected.",
        ));
    }
    Ok(())
}

fn list_item_count<T: ReadTxn>(list: &XmlElementRef, transaction: &T) -> UseResult<u32> {
    let mut count = 0_u32;
    for child in list.children(transaction) {
        let XmlOut::Element(element) = child else {
            return Err(structure_conflict(
                "A Document list contains a child that is not an item.",
            ));
        };
        if element.tag().as_ref() != "listItem" {
            return Err(structure_conflict(
                "A Document list contains a child that is not an item.",
            ));
        }
        count = count.checked_add(1).ok_or_else(list_too_large)?;
    }
    Ok(count)
}

fn ensure_paragraph_available<T: ReadTxn>(
    fragment: &impl XmlFragment,
    transaction: &T,
    paragraph_id: &str,
) -> UseResult<()> {
    for node in fragment.successors(transaction) {
        let XmlOut::Element(element) = node else {
            continue;
        };
        if document_identity_attribute(&element, transaction, PARAGRAPH_ID_ATTRIBUTE)?.as_deref()
            == Some(paragraph_id)
        {
            return Err(collaboration_error(
                "office.collaboration.mutation_identity_conflict",
                format!("Document paragraph ID '{paragraph_id}' is already in use."),
            ));
        }
    }
    Ok(())
}

fn assert_block_reaches_section<T: ReadTxn>(
    fragment: &impl XmlFragment,
    transaction: &T,
    block: &XmlElementRef,
) -> UseResult<()> {
    let mut current = match block.parent() {
        Some(XmlOut::Element(parent)) => parent,
        _ => {
            return Err(structure_conflict(
                "The target Document list is detached from a supported block container.",
            ))
        }
    };
    for _ in 0..MAX_DOCUMENT_ANCESTOR_DEPTH {
        match current.tag().as_ref() {
            "documentSection" => {
                if fragment
                    .children(transaction)
                    .any(|node| matches!(node, XmlOut::Element(candidate) if candidate == current))
                {
                    return Ok(());
                }
                return Err(structure_conflict(
                    "The target Document list is not contained by a top-level section.",
                ));
            }
            "listItem" => {
                let list = parent_element(&current, "A Document list item is detached.")?;
                if !matches!(list.tag().as_ref(), "bulletList" | "orderedList") {
                    return Err(structure_conflict(
                        "A Document list item must be a direct child of a bullet or ordered list.",
                    ));
                }
                current = parent_element(&list, "The target Document list is detached.")?;
            }
            "tableCell" | "tableHeader" => {
                let row = parent_element(&current, "A Document table cell is detached.")?;
                if row.tag().as_ref() != "tableRow" {
                    return Err(structure_conflict(
                        "A Document table cell must be a direct child of a table row.",
                    ));
                }
                let table = parent_element(&row, "A Document table row is detached.")?;
                if table.tag().as_ref() != "table" {
                    return Err(structure_conflict(
                        "A Document table row must be a direct child of a table.",
                    ));
                }
                current = parent_element(&table, "The target Document table is detached.")?;
            }
            "blockquote" => {
                current = parent_element(&current, "The target Document blockquote is detached.")?;
            }
            _ => {
                return Err(structure_conflict(
                    "The target Document list is not inside a supported section, list item, table cell, or blockquote.",
                ))
            }
        }
    }
    Err(collaboration_error(
        "office.collaboration.mutation_too_large",
        "The target Document list exceeds the supported structural depth.",
    ))
}

fn child_index<T: ReadTxn>(
    container: &XmlElementRef,
    transaction: &T,
    target: &XmlElementRef,
) -> UseResult<u32> {
    container
        .children(transaction)
        .enumerate()
        .find_map(|(index, node)| {
            matches!(node, XmlOut::Element(candidate) if candidate == *target)
                .then(|| u32::try_from(index).ok())
                .flatten()
        })
        .ok_or_else(|| {
            structure_conflict("The target Document list item cannot be located inside its list.")
        })
}

fn parent_element(element: &XmlElementRef, message: &str) -> UseResult<XmlElementRef> {
    match element.parent() {
        Some(XmlOut::Element(parent)) => Ok(parent),
        _ => Err(structure_conflict(message)),
    }
}

fn match_conflict(message: &str) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_match_conflict", message)
}

fn invalid_list_item(message: &str) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_invalid", message)
}

fn structure_conflict(message: &str) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_structure_conflict", message)
}

fn list_too_large() -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.mutation_too_large",
        "The target Document list contains too many items for a structural mutation.",
    )
}
