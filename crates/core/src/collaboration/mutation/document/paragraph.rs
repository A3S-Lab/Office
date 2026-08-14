use std::collections::HashMap;
use std::sync::Arc;

use a3s_use_core::UseResult;
use yrs::types::xml::XmlIn;
use yrs::{
    Any, Out, ReadTxn, Text, Transact, Xml, XmlElementPrelim, XmlElementRef, XmlFragment, XmlOut,
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
use super::structure::{
    ensure_paragraph_id_available, structural_paragraph, validate_paragraph_deletion,
    validate_paragraph_insertion,
};

const MAX_DOCUMENT_PARAGRAPH_TEXT_UTF16: u32 = 1_048_576;

pub(super) fn validate_insert_paragraph(
    anchor_paragraph_id: &str,
    paragraph_id: &str,
    text_id: &str,
    text: &str,
) -> UseResult<()> {
    validate_paragraph_id_input(anchor_paragraph_id, "anchorParagraphId")?;
    validate_paragraph_id_input(paragraph_id, "paragraphId")?;
    validate_paragraph_id_input(text_id, "textId")?;
    if anchor_paragraph_id == paragraph_id {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "A new Document paragraph ID must differ from its anchor paragraph ID.",
        ));
    }
    validate_paragraph_text(text, "text")
}

pub(super) fn validate_delete_paragraph(
    paragraph_id: &str,
    expected_text_id: &str,
    expected_text: &str,
) -> UseResult<()> {
    validate_paragraph_id_input(paragraph_id, "paragraphId")?;
    validate_paragraph_id_input(expected_text_id, "expectedTextId")?;
    validate_paragraph_text(expected_text, "expectedText")
}

pub(super) fn insert_paragraph(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    anchor_paragraph_id: &str,
    position: NativeOfficeCollaborationParagraphPosition,
    paragraph_id: &str,
    text_id: &str,
    text: &str,
) -> UseResult<()> {
    let root = format!("{}.document.content", manifest.namespace);
    let fragment = doc.get_or_insert_xml_fragment(root);
    let transaction = doc.transact();
    let anchor = structural_paragraph(&fragment, &transaction, anchor_paragraph_id)?;
    ensure_paragraph_id_available(&fragment, &transaction, paragraph_id)?;
    validate_paragraph_insertion(&anchor, &transaction)?;
    let table_rows = ancestor_table_rows(std::slice::from_ref(&anchor.paragraph))?;
    let row_text_id_rotations = table_row_text_id_rotations(&table_rows, &transaction)?;
    let insert_index = match position {
        NativeOfficeCollaborationParagraphPosition::Before => anchor.index,
        NativeOfficeCollaborationParagraphPosition::After => anchor
            .index
            .checked_add(1)
            .ok_or_else(document_structure_too_large)?,
    };
    drop(transaction);

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
    let mut transaction = doc.transact_mut();
    anchor
        .container
        .insert(&mut transaction, insert_index, paragraph);
    for (row, next_text_id) in row_text_id_rotations {
        row.insert_attribute(&mut transaction, ROW_TEXT_ID_ATTRIBUTE, next_text_id);
    }
    Ok(())
}

pub(super) fn delete_paragraph(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    paragraph_id: &str,
    expected_text_id: &str,
    expected_text: &str,
) -> UseResult<()> {
    let root = format!("{}.document.content", manifest.namespace);
    let fragment = doc.get_or_insert_xml_fragment(root);
    let transaction = doc.transact();
    let target = structural_paragraph(&fragment, &transaction, paragraph_id)?;
    if target.paragraph.tag().as_ref() != "paragraph" {
        return Err(document_structure_conflict(
            "Only a plain paragraph can be deleted by this mutation.",
        ));
    }
    let current_text_id =
        document_identity_attribute(&target.paragraph, &transaction, TEXT_ID_ATTRIBUTE)?
            .ok_or_else(|| {
                document_structure_conflict(
                    "The target paragraph does not have a valid Word text identity.",
                )
            })?;
    if current_text_id != expected_text_id {
        return Err(collaboration_error(
            "office.collaboration.mutation_match_conflict",
            "The target Document paragraph text identity changed after it was inspected.",
        )
        .with_suggestion(
            "Read the latest paragraph identity and text before retrying the deletion.",
        )
        .with_detail("currentTextId", current_text_id)
        .with_detail("expectedTextId", expected_text_id));
    }
    let current_text = deletable_paragraph_text(&target.paragraph, &transaction)?;
    if current_text != expected_text {
        return Err(collaboration_error(
            "office.collaboration.mutation_match_conflict",
            "The target Document paragraph text changed after it was inspected.",
        )
        .with_suggestion(
            "Read the latest paragraph identity and complete text before retrying the deletion.",
        ));
    }
    validate_paragraph_deletion(&target, &transaction)?;
    let table_rows = ancestor_table_rows(std::slice::from_ref(&target.paragraph))?;
    let row_text_id_rotations = table_row_text_id_rotations(&table_rows, &transaction)?;
    drop(transaction);
    let mut transaction = doc.transact_mut();
    target.container.remove(&mut transaction, target.index);
    for (row, next_text_id) in row_text_id_rotations {
        row.insert_attribute(&mut transaction, ROW_TEXT_ID_ATTRIBUTE, next_text_id);
    }
    Ok(())
}

fn deletable_paragraph_text<T: ReadTxn>(
    paragraph: &XmlElementRef,
    transaction: &T,
) -> UseResult<String> {
    let mut text = String::new();
    for child in paragraph.children(transaction) {
        let XmlOut::Text(child) = child else {
            return Err(document_structure_conflict(
                "A paragraph containing inline atoms cannot be deleted by this mutation.",
            ));
        };
        for chunk in child.diff(transaction, |_| ()) {
            if chunk
                .attributes
                .as_ref()
                .is_some_and(|attributes| attributes.keys().any(|key| is_review_attribute(key)))
            {
                return Err(document_structure_conflict(
                    "A paragraph containing comment or tracked-change marks cannot be deleted by this mutation.",
                ));
            }
            match chunk.insert {
                Out::Any(Any::String(value)) => text.push_str(&value),
                _ => {
                    return Err(document_structure_conflict(
                        "A paragraph containing inline embeds cannot be deleted by this mutation.",
                    ))
                }
            }
        }
    }
    if utf16_len(&text)? > MAX_DOCUMENT_PARAGRAPH_TEXT_UTF16 {
        return Err(document_structure_conflict(
            "The target paragraph is too large for a bounded structural mutation.",
        ));
    }
    Ok(text)
}

fn validate_paragraph_text(value: &str, field: &str) -> UseResult<()> {
    let length = utf16_len(value)?;
    if length <= MAX_DOCUMENT_PARAGRAPH_TEXT_UTF16 {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.mutation_too_large",
        format!(
            "Document mutation field '{field}' is {length} UTF-16 code units; the limit is {MAX_DOCUMENT_PARAGRAPH_TEXT_UTF16}."
        ),
    )
    .with_detail("lengthUtf16", length as u64)
    .with_detail(
        "maxLengthUtf16",
        MAX_DOCUMENT_PARAGRAPH_TEXT_UTF16 as u64,
    ))
}

fn is_review_attribute(value: &str) -> bool {
    matches!(
        value.split_once("--").map_or(value, |(name, _)| name),
        "documentComment" | "documentChange"
    )
}

fn document_structure_conflict(message: &str) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_structure_conflict", message)
}

fn document_structure_too_large() -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.mutation_too_large",
        "The Document container contains too many blocks for a structural mutation.",
    )
}

#[cfg(test)]
mod tests;
