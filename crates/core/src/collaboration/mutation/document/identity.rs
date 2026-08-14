use a3s_use_core::UseResult;
use yrs::{Any, Out, ReadTxn, Xml, XmlElementRef};

use super::super::super::collaboration_error;

pub(super) const PARAGRAPH_ID_ATTRIBUTE: &str = "paragraphId";
pub(super) const TEXT_ID_ATTRIBUTE: &str = "textId";

const MAX_WORD_PARAGRAPH_ID: u32 = 0x7fff_ffff;

pub(super) fn validate_paragraph_id_input(value: &str, field: &str) -> UseResult<()> {
    if canonical_document_paragraph_id(value).as_deref() == Some(value) {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.mutation_invalid",
        format!(
            "Document mutation field '{field}' must be an uppercase eight-digit positive 31-bit Word paragraph ID."
        ),
    ))
}

pub(super) fn document_identity_attribute<T: ReadTxn>(
    paragraph: &XmlElementRef,
    transaction: &T,
    name: &str,
) -> UseResult<Option<String>> {
    match paragraph.get_attribute(transaction, name) {
        Some(Out::Any(Any::String(value))) => canonical_document_paragraph_id(&value)
            .map(Some)
            .ok_or_else(|| invalid_shared_identity(name)),
        Some(_) => Err(invalid_shared_identity(name)),
        None => Ok(None),
    }
}

pub(super) fn paragraph_text_id_rotations<T: ReadTxn>(
    paragraphs: &[XmlElementRef],
    transaction: &T,
) -> UseResult<Vec<(XmlElementRef, String)>> {
    let mut rotations = Vec::new();
    for paragraph in paragraphs {
        if !is_identity_paragraph_tag(paragraph.tag()) {
            continue;
        }
        let paragraph_id =
            document_identity_attribute(paragraph, transaction, PARAGRAPH_ID_ATTRIBUTE)?;
        let text_id = document_identity_attribute(paragraph, transaction, TEXT_ID_ATTRIBUTE)?;
        match (paragraph_id, text_id) {
            (None, None) => {}
            (Some(_), Some(text_id)) => {
                rotations.push((paragraph.clone(), next_document_paragraph_id(&text_id)))
            }
            _ => {
                return Err(collaboration_error(
                    "office.collaboration.content_invalid",
                    "The edited Document paragraph has an incomplete Word paragraph identity.",
                ))
            }
        }
    }
    Ok(rotations)
}

pub(super) fn is_identity_paragraph_tag(tag: &str) -> bool {
    matches!(tag, "paragraph" | "heading" | "documentCaption")
}

fn canonical_document_paragraph_id(value: &str) -> Option<String> {
    let normalized = value.trim().to_ascii_uppercase();
    if normalized.len() != 8 || !normalized.bytes().all(|value| value.is_ascii_hexdigit()) {
        return None;
    }
    let number = u32::from_str_radix(&normalized, 16).ok()?;
    (1..=MAX_WORD_PARAGRAPH_ID)
        .contains(&number)
        .then_some(normalized)
}

fn next_document_paragraph_id(value: &str) -> String {
    let current = u32::from_str_radix(value, 16).unwrap_or(MAX_WORD_PARAGRAPH_ID);
    let next = if current >= MAX_WORD_PARAGRAPH_ID {
        1
    } else {
        current + 1
    };
    format!("{next:08X}")
}

fn invalid_shared_identity(name: &str) -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.content_invalid",
        format!("The shared Document paragraph attribute '{name}' is invalid."),
    )
}
