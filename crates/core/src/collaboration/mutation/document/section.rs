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
    document_identity_attribute, validate_paragraph_id_input, PARAGRAPH_ID_ATTRIBUTE,
    TEXT_ID_ATTRIBUTE,
};

const MAX_DOCUMENT_PARAGRAPH_TEXT_UTF16: u32 = 1_048_576;
const MAX_SECTION_ID_LEN: usize = 128;
const SECTION_ID_ATTRIBUTE: &str = "id";

pub(super) fn validate_insert_section(
    anchor_section_id: &str,
    expected_paragraph_id: &str,
    expected_text_id: &str,
    section_id: &str,
    paragraph_id: &str,
    text_id: &str,
    text: &str,
) -> UseResult<()> {
    validate_section_id(anchor_section_id, "anchorSectionId")?;
    validate_paragraph_id_input(expected_paragraph_id, "expectedParagraphId")?;
    validate_paragraph_id_input(expected_text_id, "expectedTextId")?;
    validate_section_id(section_id, "sectionId")?;
    validate_paragraph_id_input(paragraph_id, "paragraphId")?;
    validate_paragraph_id_input(text_id, "textId")?;
    if anchor_section_id == section_id {
        return Err(invalid_section(
            "A new Document section ID must differ from the observed section ID.",
        ));
    }
    if expected_paragraph_id == paragraph_id {
        return Err(invalid_section(
            "A new Document section paragraph ID must differ from the observed paragraph ID.",
        ));
    }
    if utf16_len(text)? > MAX_DOCUMENT_PARAGRAPH_TEXT_UTF16 {
        return Err(collaboration_error(
            "office.collaboration.mutation_too_large",
            "The new Document section text exceeds the supported UTF-16 length.",
        ));
    }
    Ok(())
}

pub(super) fn validate_delete_section(
    section_id: &str,
    paragraph_id: &str,
    expected_text_id: &str,
    expected_text: &str,
) -> UseResult<()> {
    validate_section_id(section_id, "sectionId")?;
    validate_paragraph_id_input(paragraph_id, "paragraphId")?;
    validate_paragraph_id_input(expected_text_id, "expectedTextId")?;
    if utf16_len(expected_text)? > MAX_DOCUMENT_PARAGRAPH_TEXT_UTF16 {
        return Err(collaboration_error(
            "office.collaboration.mutation_too_large",
            "The observed Document section text exceeds the supported UTF-16 length.",
        ));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
pub(super) fn insert_section(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    anchor_section_id: &str,
    expected_paragraph_id: &str,
    expected_text_id: &str,
    position: NativeOfficeCollaborationParagraphPosition,
    section_id: &str,
    paragraph_id: &str,
    text_id: &str,
    text: &str,
) -> UseResult<()> {
    let root = format!("{}.document.content", manifest.namespace);
    let fragment = doc.get_or_insert_xml_fragment(root);
    let transaction = doc.transact();
    let anchor = top_level_section(&fragment, &transaction, anchor_section_id)?;
    let observed = paragraph(&fragment, &transaction, expected_paragraph_id)?;
    let current_text_id = required_text_id(&observed, &transaction)?;
    if current_text_id != expected_text_id {
        return Err(match_conflict(
            "The observed Document section paragraph identity changed after it was inspected.",
        ));
    }
    let containing = section_of_paragraph(&observed)?;
    if containing != anchor {
        return Err(structure_conflict(
            "The observed paragraph is not inside the observed Document section.",
        ));
    }
    ensure_section_available(&fragment, &transaction, section_id)?;
    ensure_paragraph_available(&fragment, &transaction, paragraph_id)?;
    let anchor_index = child_index(&fragment, &transaction, &anchor)?;
    let insert_index = match position {
        NativeOfficeCollaborationParagraphPosition::Before => anchor_index,
        NativeOfficeCollaborationParagraphPosition::After => {
            anchor_index.checked_add(1).ok_or_else(section_too_large)?
        }
    };
    drop(transaction);

    let created = new_section(section_id, paragraph_id, text_id, text);
    let mut transaction = doc.transact_mut();
    fragment.insert(&mut transaction, insert_index, created);
    Ok(())
}

pub(super) fn delete_section(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    section_id: &str,
    paragraph_id: &str,
    expected_text_id: &str,
    expected_text: &str,
) -> UseResult<()> {
    let root = format!("{}.document.content", manifest.namespace);
    let fragment = doc.get_or_insert_xml_fragment(root);
    let transaction = doc.transact();
    let section = top_level_section(&fragment, &transaction, section_id)?;
    if section_count(&fragment, &transaction)? <= 1 {
        return Err(structure_conflict(
            "A Document must retain at least one section.",
        ));
    }
    prove_single_paragraph_section(
        &section,
        &transaction,
        paragraph_id,
        expected_text_id,
        expected_text,
    )?;
    let index = child_index(&fragment, &transaction, &section)?;
    drop(transaction);

    let mut transaction = doc.transact_mut();
    fragment.remove(&mut transaction, index);
    Ok(())
}

fn new_section(
    section_id: &str,
    paragraph_id: &str,
    text_id: &str,
    text: &str,
) -> XmlElementPrelim {
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
    let mut section_attributes = HashMap::new();
    section_attributes.insert(Arc::from(SECTION_ID_ATTRIBUTE), section_id.to_owned());
    XmlElementPrelim {
        tag: Arc::from("documentSection"),
        attributes: section_attributes,
        children: vec![XmlIn::from(paragraph)],
    }
}

fn top_level_section<T: ReadTxn>(
    fragment: &impl XmlFragment,
    transaction: &T,
    section_id: &str,
) -> UseResult<XmlElementRef> {
    let mut matches = Vec::new();
    for node in fragment.children(transaction) {
        let XmlOut::Element(element) = node else {
            return Err(structure_conflict(
                "The Document contains a top-level child that is not a section.",
            ));
        };
        if element.tag().as_ref() != "documentSection" {
            return Err(structure_conflict(
                "The Document contains a top-level child that is not a section.",
            ));
        }
        if section_id_attribute(&element, transaction)?.as_deref() == Some(section_id) {
            matches.push(element);
        }
    }
    match matches.len() {
        0 => Err(collaboration_error(
            "office.collaboration.mutation_target_missing",
            format!("Document section '{section_id}' does not exist."),
        )),
        1 => Ok(matches.pop().expect("length checked")),
        count => Err(collaboration_error(
            "office.collaboration.mutation_identity_conflict",
            format!("Document section ID '{section_id}' is assigned to {count} live sections."),
        )),
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
            "The observed Document section paragraph does not have a valid Word text identity.",
        )
    })
}

fn section_of_paragraph(paragraph: &XmlElementRef) -> UseResult<XmlElementRef> {
    let mut current = paragraph.parent();
    for _ in 0..64 {
        let Some(XmlOut::Element(parent)) = current else {
            break;
        };
        if parent.tag().as_ref() == "documentSection" {
            return Ok(parent);
        }
        current = parent.parent();
    }
    Err(structure_conflict(
        "The observed paragraph is not inside a Document section.",
    ))
}

fn prove_single_paragraph_section<T: ReadTxn>(
    section: &XmlElementRef,
    transaction: &T,
    paragraph_id: &str,
    expected_text_id: &str,
    expected_text: &str,
) -> UseResult<()> {
    let mut blocks = Vec::new();
    for child in section.children(transaction) {
        let XmlOut::Element(element) = child else {
            return Err(structure_conflict(
                "A Document section contains a child that is not a block.",
            ));
        };
        blocks.push(element);
    }
    if blocks.len() != 1 || blocks[0].tag().as_ref() != "paragraph" {
        return Err(structure_conflict(
            "A Document section can be deleted only when it holds one plain paragraph.",
        ));
    }
    let paragraph = &blocks[0];
    let current_id = document_identity_attribute(paragraph, transaction, PARAGRAPH_ID_ATTRIBUTE)?
        .ok_or_else(|| {
        structure_conflict(
            "The Document section paragraph does not have a valid Word paragraph identity.",
        )
    })?;
    if current_id != paragraph_id {
        return Err(match_conflict(
            "The target Document section paragraph changed after it was inspected.",
        ));
    }
    let current_text_id = required_text_id(paragraph, transaction)?;
    if current_text_id != expected_text_id {
        return Err(match_conflict(
            "The target Document section paragraph identity changed after it was inspected.",
        ));
    }
    let mut text = String::new();
    for child in paragraph.children(transaction) {
        let XmlOut::Text(child) = child else {
            return Err(structure_conflict(
                "A Document section paragraph containing inline atoms cannot be deleted by this frame.",
            ));
        };
        text.push_str(&child.get_string(transaction));
    }
    if text != expected_text {
        return Err(match_conflict(
            "The target Document section text changed after it was inspected.",
        ));
    }
    Ok(())
}

fn section_count<T: ReadTxn>(fragment: &impl XmlFragment, transaction: &T) -> UseResult<u32> {
    let mut count = 0_u32;
    for node in fragment.children(transaction) {
        let XmlOut::Element(element) = node else {
            return Err(structure_conflict(
                "The Document contains a top-level child that is not a section.",
            ));
        };
        if element.tag().as_ref() != "documentSection" {
            return Err(structure_conflict(
                "The Document contains a top-level child that is not a section.",
            ));
        }
        count = count.checked_add(1).ok_or_else(section_too_large)?;
    }
    Ok(count)
}

fn ensure_section_available<T: ReadTxn>(
    fragment: &impl XmlFragment,
    transaction: &T,
    section_id: &str,
) -> UseResult<()> {
    for node in fragment.successors(transaction) {
        let XmlOut::Element(element) = node else {
            continue;
        };
        if element.tag().as_ref() == "documentSection"
            && section_id_attribute(&element, transaction)?.as_deref() == Some(section_id)
        {
            return Err(collaboration_error(
                "office.collaboration.mutation_identity_conflict",
                format!("Document section ID '{section_id}' is already in use."),
            ));
        }
    }
    Ok(())
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

fn section_id_attribute<T: ReadTxn>(
    element: &XmlElementRef,
    transaction: &T,
) -> UseResult<Option<String>> {
    match element.get_attribute(transaction, SECTION_ID_ATTRIBUTE) {
        Some(yrs::Out::Any(yrs::Any::String(value))) => {
            let value = value.to_string();
            if valid_section_id(&value) {
                Ok(Some(value))
            } else {
                Err(collaboration_error(
                    "office.collaboration.content_invalid",
                    "The Document section identity is not a supported section ID.",
                ))
            }
        }
        Some(_) => Err(collaboration_error(
            "office.collaboration.content_invalid",
            "The Document section identity is not a supported section ID.",
        )),
        None => Ok(None),
    }
}

fn child_index<T: ReadTxn>(
    fragment: &impl XmlFragment,
    transaction: &T,
    target: &XmlElementRef,
) -> UseResult<u32> {
    fragment
        .children(transaction)
        .enumerate()
        .find_map(|(index, node)| {
            matches!(node, XmlOut::Element(candidate) if candidate == *target)
                .then(|| u32::try_from(index).ok())
                .flatten()
        })
        .ok_or_else(|| structure_conflict("The target Document section cannot be located."))
}

fn validate_section_id(value: &str, field: &str) -> UseResult<()> {
    if valid_section_id(value) {
        return Ok(());
    }
    Err(invalid_section(format!(
        "Document mutation field '{field}' must be a section ID of 1 to {MAX_SECTION_ID_LEN} letters, digits, hyphens, or underscores."
    )))
}

fn valid_section_id(value: &str) -> bool {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if value.len() > MAX_SECTION_ID_LEN || !first.is_ascii_alphanumeric() {
        return false;
    }
    chars.all(|character| character.is_ascii_alphanumeric() || character == '-' || character == '_')
}

fn match_conflict(message: &str) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_match_conflict", message)
}

fn invalid_section(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_invalid", message)
}

fn structure_conflict(message: &str) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_structure_conflict", message)
}

fn section_too_large() -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.mutation_too_large",
        "The Document contains too many sections for a structural mutation.",
    )
}
