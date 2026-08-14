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
    document_identity_attribute, is_identity_paragraph_tag, validate_paragraph_id_input,
    PARAGRAPH_ID_ATTRIBUTE, TEXT_ID_ATTRIBUTE,
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
    let anchor = direct_section_paragraph(&fragment, &transaction, anchor_paragraph_id)?;
    ensure_paragraph_id_available(&fragment, &transaction, paragraph_id)?;
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
    anchor
        .section
        .insert(&mut doc.transact_mut(), insert_index, paragraph);
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
    let target = direct_section_paragraph(&fragment, &transaction, paragraph_id)?;
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
    if target.section.len(&transaction) <= 1 {
        return Err(document_structure_conflict(
            "A Document section must retain at least one block.",
        ));
    }
    drop(transaction);
    target.section.remove(&mut doc.transact_mut(), target.index);
    Ok(())
}

struct DirectSectionParagraph {
    section: XmlElementRef,
    paragraph: XmlElementRef,
    index: u32,
}

fn direct_section_paragraph<T: ReadTxn>(
    fragment: &impl XmlFragment,
    transaction: &T,
    paragraph_id: &str,
) -> UseResult<DirectSectionParagraph> {
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
    let paragraph = match matches.len() {
        0 => {
            return Err(collaboration_error(
                "office.collaboration.mutation_target_missing",
                format!("Document paragraph '{paragraph_id}' does not exist."),
            ))
        }
        1 => matches.pop().ok_or_else(|| {
            collaboration_error(
                "office.collaboration.mutation_target_missing",
                format!("Document paragraph '{paragraph_id}' does not exist."),
            )
        })?,
        count => {
            return Err(collaboration_error(
                "office.collaboration.mutation_identity_conflict",
                format!(
                    "Document paragraph ID '{paragraph_id}' is assigned to {count} live nodes."
                ),
            ))
        }
    };
    if !is_identity_paragraph_tag(paragraph.tag()) {
        return Err(document_structure_conflict(
            "The target Word paragraph identity belongs to an unsupported node type.",
        ));
    }
    let Some(XmlOut::Element(section)) = paragraph.parent() else {
        return Err(document_structure_conflict(
            "The target paragraph is not a direct child of a Document section.",
        ));
    };
    if section.tag().as_ref() != "documentSection"
        || !fragment
            .children(transaction)
            .any(|node| matches!(node, XmlOut::Element(candidate) if candidate == section))
    {
        return Err(document_structure_conflict(
            "The target paragraph is not a direct child of a top-level Document section.",
        ));
    }
    let index = section
        .children(transaction)
        .enumerate()
        .find_map(|(index, node)| {
            matches!(node, XmlOut::Element(candidate) if candidate == paragraph)
                .then(|| u32::try_from(index).ok())
                .flatten()
        })
        .ok_or_else(|| {
            document_structure_conflict(
                "The target paragraph cannot be located inside its Document section.",
            )
        })?;
    Ok(DirectSectionParagraph {
        section,
        paragraph,
        index,
    })
}

fn ensure_paragraph_id_available<T: ReadTxn>(
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
        "The Document section contains too many blocks for a structural mutation.",
    )
}

#[cfg(test)]
mod tests {
    use yrs::{Transact, Xml, XmlFragment};

    use super::*;
    use crate::collaboration::document::new_replica_document;
    use crate::collaboration::{
        NativeOfficeCollaborationActorKind, NativeOfficeCollaborationArtifactKind,
        NativeOfficeCollaborationMode,
    };

    #[test]
    fn paragraph_contract_requires_canonical_bounded_word_identities() {
        for value in ["", "00000000", "80000000", "0000000a", " 0000000A"] {
            let error =
                validate_insert_paragraph("00000001", value, "00000002", "value").unwrap_err();
            assert_eq!(error.code, "office.collaboration.mutation_invalid");
        }
        validate_insert_paragraph("00000001", "0000000A", "7FFFFFFF", "value").unwrap();
    }

    #[test]
    fn paragraph_insert_and_delete_preserve_section_schema_and_identity_guards() {
        let doc = document_with_paragraphs(&[
            ("00000001", "00000002", "First"),
            ("00000003", "00000004", "Last"),
        ]);
        let manifest = manifest();

        insert_paragraph(
            &doc,
            &manifest,
            "00000001",
            NativeOfficeCollaborationParagraphPosition::After,
            "0000000A",
            "0000000B",
            "Native",
        )
        .unwrap();

        assert_eq!(
            paragraphs(&doc),
            vec![
                (
                    "00000001".to_owned(),
                    "00000002".to_owned(),
                    "First".to_owned()
                ),
                (
                    "0000000A".to_owned(),
                    "0000000B".to_owned(),
                    "Native".to_owned()
                ),
                (
                    "00000003".to_owned(),
                    "00000004".to_owned(),
                    "Last".to_owned()
                ),
            ]
        );
        let collision = insert_paragraph(
            &doc,
            &manifest,
            "00000001",
            NativeOfficeCollaborationParagraphPosition::Before,
            "0000000A",
            "0000000C",
            "Collision",
        )
        .unwrap_err();
        assert_eq!(
            collision.code,
            "office.collaboration.mutation_identity_conflict"
        );

        let stale =
            delete_paragraph(&doc, &manifest, "0000000A", "0000000C", "Native").unwrap_err();
        assert_eq!(stale.code, "office.collaboration.mutation_match_conflict");
        delete_paragraph(&doc, &manifest, "0000000A", "0000000B", "Native").unwrap();
        assert_eq!(paragraphs(&doc).len(), 2);
    }

    #[test]
    fn paragraph_delete_keeps_the_required_last_section_block() {
        let doc = document_with_paragraphs(&[("00000001", "00000002", "Only")]);
        let error =
            delete_paragraph(&doc, &manifest(), "00000001", "00000002", "Only").unwrap_err();
        assert_eq!(
            error.code,
            "office.collaboration.mutation_structure_conflict"
        );
        assert_eq!(paragraphs(&doc).len(), 1);
    }

    fn document_with_paragraphs(values: &[(&str, &str, &str)]) -> yrs::Doc {
        let doc = new_replica_document(
            7,
            "a3s.office",
            NativeOfficeCollaborationArtifactKind::Document,
        );
        let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
        let mut transaction = doc.transact_mut();
        let section =
            fragment.push_back(&mut transaction, XmlElementPrelim::empty("documentSection"));
        section.insert_attribute(&mut transaction, "id", "document-section-1");
        for (paragraph_id, text_id, text) in values {
            let paragraph =
                section.push_back(&mut transaction, XmlElementPrelim::empty("paragraph"));
            paragraph.insert_attribute(&mut transaction, PARAGRAPH_ID_ATTRIBUTE, *paragraph_id);
            paragraph.insert_attribute(&mut transaction, TEXT_ID_ATTRIBUTE, *text_id);
            if !text.is_empty() {
                paragraph.push_back(&mut transaction, XmlTextPrelim::new(*text));
            }
        }
        drop(transaction);
        doc
    }

    fn paragraphs(doc: &yrs::Doc) -> Vec<(String, String, String)> {
        let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
        let transaction = doc.transact();
        fragment
            .successors(&transaction)
            .filter_map(|node| match node {
                XmlOut::Element(element) if element.tag().as_ref() == "paragraph" => {
                    let paragraph_id =
                        document_identity_attribute(&element, &transaction, PARAGRAPH_ID_ATTRIBUTE)
                            .unwrap()?;
                    let text_id =
                        document_identity_attribute(&element, &transaction, TEXT_ID_ATTRIBUTE)
                            .unwrap()?;
                    Some((
                        paragraph_id,
                        text_id,
                        deletable_paragraph_text(&element, &transaction).unwrap(),
                    ))
                }
                _ => None,
            })
            .collect()
    }

    fn manifest() -> NativeOfficeCollaborationManifest {
        NativeOfficeCollaborationManifest {
            format: String::new(),
            schema_version: 1,
            protocol: String::new(),
            protocol_version: 1,
            namespace: "a3s.office".to_owned(),
            artifact_id: "document".to_owned(),
            kind: NativeOfficeCollaborationArtifactKind::Document,
            actor_id: "agent".to_owned(),
            actor_kind: NativeOfficeCollaborationActorKind::Agent,
            mode: NativeOfficeCollaborationMode::Edit,
            client_id: 7,
        }
    }
}
