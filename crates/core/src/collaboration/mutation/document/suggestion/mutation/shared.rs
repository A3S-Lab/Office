use a3s_use_core::UseResult;
use yrs::types::Attrs;
use yrs::{Any, Out, ReadTxn, Text, Xml, XmlElementRef, XmlFragment, XmlOut, XmlTextRef};

use super::super::super::comment::valid_canonical_utc_timestamp;
use super::super::super::identity::{
    document_identity_attribute, is_identity_paragraph_tag, PARAGRAPH_ID_ATTRIBUTE,
    TEXT_ID_ATTRIBUTE,
};
use super::super::{parse_suggestion_identity, suggestion_identity, DocumentSuggestionIdentity};
use crate::collaboration::document::is_document_change_attribute;
use crate::collaboration::{collaboration_error, NativeOfficeCollaborationDocumentSuggestionKind};

use super::super::super::super::{is_utf16_boundary, utf16_len};

pub(super) const MAX_DOCUMENT_SUGGESTION_TEXT_UTF16: u32 = 1_048_576;
pub(super) const MAX_DOCUMENT_SUGGESTION_DECISIONS: usize = 4_096;

#[derive(Clone)]
pub(super) struct SuggestionTarget {
    pub paragraph: XmlElementRef,
    pub text: Option<XmlTextRef>,
    pub chunks: Vec<SuggestionTextChunk>,
}

#[derive(Clone)]
pub(super) struct SuggestionTextChunk {
    pub start_utf16: u32,
    pub end_utf16: u32,
    pub attributes: Option<Attrs>,
}

#[derive(Clone)]
pub(super) struct LiveSuggestionSegment {
    pub text: XmlTextRef,
    pub paragraph: Option<XmlElementRef>,
    pub attributes: Vec<(String, Any)>,
    pub identity: DocumentSuggestionIdentity,
    pub start_utf16: u32,
    pub end_utf16: u32,
}

#[allow(clippy::too_many_arguments)]
pub(super) fn resolve_suggestion_target<T: ReadTxn>(
    transaction: &T,
    fragment: &yrs::XmlFragmentRef,
    paragraph_id: &str,
    expected_text_id: &str,
    start_utf16: u32,
    end_utf16: u32,
    expected_text: &str,
) -> UseResult<SuggestionTarget> {
    let mut found = None;
    for node in fragment.successors(transaction) {
        let XmlOut::Element(paragraph) = node else {
            continue;
        };
        if !is_identity_paragraph_tag(paragraph.tag()) {
            continue;
        }
        if document_identity_attribute(&paragraph, transaction, PARAGRAPH_ID_ATTRIBUTE)?.as_deref()
            != Some(paragraph_id)
        {
            continue;
        }
        if found.is_some() {
            return Err(content_invalid(format!(
                "The shared Document repeats paragraph ID '{paragraph_id}'."
            )));
        }
        found = Some(paragraph);
    }
    let paragraph = found.ok_or_else(|| {
        match_conflict(format!(
            "Document paragraph ID '{paragraph_id}' does not exist for this suggestion."
        ))
    })?;
    if paragraph.tag().as_ref() != "paragraph" {
        return Err(structure_conflict(
            "A native Document suggestion can target only a plain paragraph.",
        ));
    }
    let current_text_id = document_identity_attribute(&paragraph, transaction, TEXT_ID_ATTRIBUTE)?;
    if current_text_id.as_deref() != Some(expected_text_id) {
        return Err(match_conflict(format!(
            "Document paragraph ID '{paragraph_id}' has text ID '{}', not expected text ID '{expected_text_id}'.",
            current_text_id.as_deref().unwrap_or("<missing>")
        )));
    }

    let mut text = None;
    for child in paragraph.children(transaction) {
        match child {
            XmlOut::Text(child) if text.is_none() => text = Some(child),
            _ => {
                return Err(structure_conflict(
                    "A native Document suggestion requires at most one text node and no inline objects.",
                ))
            }
        }
    }

    let mut current = String::new();
    let mut chunks = Vec::new();
    let mut cursor = 0_u32;
    if let Some(text) = &text {
        for chunk in text.diff(transaction, |_| ()) {
            let Out::Any(Any::String(value)) = chunk.insert else {
                return Err(structure_conflict(
                    "A native Document suggestion cannot cross an inline embed.",
                ));
            };
            let length = utf16_len(&value)?;
            let end = cursor.checked_add(length).ok_or_else(|| {
                collaboration_error(
                    "office.collaboration.mutation_too_large",
                    "The Document suggestion offset exceeds the supported UTF-16 range.",
                )
            })?;
            current.push_str(&value);
            chunks.push(SuggestionTextChunk {
                start_utf16: cursor,
                end_utf16: end,
                attributes: chunk.attributes.map(|attributes| *attributes),
            });
            cursor = end;
        }
    }
    if cursor > MAX_DOCUMENT_SUGGESTION_TEXT_UTF16 {
        return Err(structure_conflict(
            "The target paragraph is too large for a bounded Document suggestion mutation.",
        ));
    }
    if start_utf16 > end_utf16
        || end_utf16 > cursor
        || !is_utf16_boundary(&current, start_utf16)
        || !is_utf16_boundary(&current, end_utf16)
    {
        return Err(range_conflict(start_utf16, end_utf16, cursor));
    }
    let selected = utf16_slice(&current, start_utf16, end_utf16)
        .ok_or_else(|| range_conflict(start_utf16, end_utf16, cursor))?;
    if selected != expected_text {
        return Err(match_conflict(format!(
            "The current Document suggestion selection is '{selected}', not the expected text."
        )));
    }
    Ok(SuggestionTarget {
        paragraph,
        text,
        chunks,
    })
}

pub(super) fn insertion_attributes(target: &SuggestionTarget, index_utf16: u32) -> Attrs {
    let selected = target
        .chunks
        .iter()
        .rev()
        .find(|chunk| chunk.start_utf16 < index_utf16 && index_utf16 <= chunk.end_utf16)
        .or_else(|| {
            target
                .chunks
                .iter()
                .find(|chunk| chunk.start_utf16 <= index_utf16 && index_utf16 < chunk.end_utf16)
        });
    let mut attributes = selected
        .and_then(|chunk| chunk.attributes.clone())
        .unwrap_or_default();
    attributes.retain(|key, _| !is_document_change_attribute(key));
    attributes
}

pub(super) fn selection_overlaps_suggestion(
    target: &SuggestionTarget,
    start_utf16: u32,
    end_utf16: u32,
) -> bool {
    target.chunks.iter().any(|chunk| {
        let intersects = if start_utf16 == end_utf16 {
            chunk.start_utf16 < start_utf16 && start_utf16 < chunk.end_utf16
        } else {
            chunk.start_utf16 < end_utf16 && chunk.end_utf16 > start_utf16
        };
        intersects
            && chunk.attributes.as_ref().is_some_and(|attributes| {
                attributes
                    .keys()
                    .any(|key| is_document_change_attribute(key))
            })
    })
}

pub(super) fn collect_live_suggestion_segments<T: ReadTxn>(
    transaction: &T,
    fragment: &yrs::XmlFragmentRef,
) -> UseResult<Vec<LiveSuggestionSegment>> {
    let mut segments = Vec::new();
    for node in fragment.successors(transaction) {
        let XmlOut::Text(text) = node else {
            continue;
        };
        let paragraph = match text.parent() {
            Some(XmlOut::Element(value)) if is_identity_paragraph_tag(value.tag()) => Some(value),
            _ => None,
        };
        let mut cursor = 0_u32;
        for chunk in text.diff(transaction, |_| ()) {
            let length = match &chunk.insert {
                Out::Any(Any::String(value)) => utf16_len(value.as_ref())?,
                _ => 1,
            };
            let end = cursor.checked_add(length).ok_or_else(|| {
                content_invalid("A Document suggestion segment offset is too large.".to_owned())
            })?;
            let Some(identity) = suggestion_identity(chunk.attributes.as_deref())? else {
                cursor = end;
                continue;
            };
            if !matches!(&chunk.insert, Out::Any(Any::String(value)) if !value.is_empty()) {
                return Err(content_invalid(
                    "Document suggestions may mark non-empty text only.".to_owned(),
                ));
            }
            let attributes = chunk
                .attributes
                .as_ref()
                .into_iter()
                .flat_map(|attributes| attributes.iter())
                .filter(|(key, _)| is_document_change_attribute(key))
                .map(|(key, value)| {
                    let parsed = parse_suggestion_identity(value)?;
                    if parsed != identity {
                        return Err(content_invalid(
                            "Overlapping Document suggestion marks carry different identities."
                                .to_owned(),
                        ));
                    }
                    Ok((key.to_string(), value.clone()))
                })
                .collect::<UseResult<Vec<_>>>()?;
            segments.push(LiveSuggestionSegment {
                text: text.clone(),
                paragraph: paragraph.clone(),
                attributes,
                identity,
                start_utf16: cursor,
                end_utf16: end,
            });
            cursor = end;
        }
    }
    Ok(segments)
}

pub(super) fn validate_identifier(value: &str, field: &str) -> UseResult<()> {
    if !value.is_empty() && value == value.trim() && value.encode_utf16().count() <= 256 {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.mutation_invalid",
        format!("Document suggestion field '{field}' must contain 1 to 256 non-padded characters."),
    ))
}

pub(super) fn validate_change_id(value: &str, field: &str) -> UseResult<()> {
    validate_identifier(value, field)?;
    let decision_id = format!(
        "{}:{value}",
        NativeOfficeCollaborationDocumentSuggestionKind::Insertion.as_str()
    );
    if decision_id.encode_utf16().count() <= 256 {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.mutation_invalid",
        format!(
            "Document suggestion field '{field}' is too long to form a browser-compatible decision ID."
        ),
    ))
}

pub(super) fn validate_timestamp(value: &str, field: &str) -> UseResult<()> {
    if valid_canonical_utc_timestamp(value) && value.encode_utf16().count() <= 256 {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.mutation_invalid",
        format!(
            "Document suggestion field '{field}' must be a canonical UTC timestamp such as '2026-08-17T10:00:00.000Z'."
        ),
    ))
}

pub(super) fn validate_text(value: &str, field: &str) -> UseResult<u32> {
    let length = utf16_len(value)?;
    if length <= MAX_DOCUMENT_SUGGESTION_TEXT_UTF16 {
        return Ok(length);
    }
    Err(collaboration_error(
        "office.collaboration.mutation_too_large",
        format!(
            "Document suggestion field '{field}' exceeds {MAX_DOCUMENT_SUGGESTION_TEXT_UTF16} UTF-16 code units."
        ),
    ))
}

pub(super) fn match_conflict(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_match_conflict", message).with_suggestion(
        "Read the latest native collaboration projection before retrying the Document suggestion mutation.",
    )
}

pub(super) fn identity_conflict(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_identity_conflict", message)
}

pub(super) fn content_invalid(message: String) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.content_invalid", message)
}

fn structure_conflict(message: &str) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_structure_conflict", message)
}

fn range_conflict(start: u32, end: u32, length: u32) -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.mutation_range_invalid",
        "The Document suggestion range is outside the current paragraph or splits a UTF-16 surrogate pair.",
    )
    .with_detail("startUtf16", start as u64)
    .with_detail("endUtf16", end as u64)
    .with_detail("lengthUtf16", length as u64)
}

fn utf16_slice(value: &str, start_utf16: u32, end_utf16: u32) -> Option<&str> {
    let start = utf16_byte_index(value, start_utf16)?;
    let end = utf16_byte_index(value, end_utf16)?;
    value.get(start..end)
}

fn utf16_byte_index(value: &str, offset: u32) -> Option<usize> {
    if offset == 0 {
        return Some(0);
    }
    let mut cursor = 0_u32;
    for (byte_index, character) in value.char_indices() {
        cursor = cursor.checked_add(character.len_utf16() as u32)?;
        if cursor == offset {
            return Some(byte_index + character.len_utf8());
        }
        if cursor > offset {
            return None;
        }
    }
    (cursor == offset).then_some(value.len())
}
