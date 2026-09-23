use a3s_use_core::UseResult;
use yrs::{Any, GetString, Map, Out, ReadTxn, Text, Transact, Xml, XmlFragment, XmlOut};

use super::document::{canonical_state_vector, inspect_document};
use super::mutation::document::comment::project_document_comments;
use super::mutation::document::identity::{
    document_identity_attribute, is_identity_paragraph_tag, PARAGRAPH_ID_ATTRIBUTE,
    TEXT_ID_ATTRIBUTE,
};
use super::mutation::document::suggestion::{
    project_document_change_decisions, project_document_suggestions,
};
use super::mutation::{
    project_pdf_content, project_presentation_content, project_spreadsheet_content,
};
use super::{
    collaboration_error, sha256_hex, NativeOfficeCollaborationArtifactKind,
    NativeOfficeCollaborationDocumentParagraph, NativeOfficeCollaborationManifest,
    NativeOfficeCollaborationMarkdownSlice, NativeOfficeCollaborationProjectedContent,
    NativeOfficeCollaborationProjection, NATIVE_OFFICE_COLLABORATION_PROJECTION_SCHEMA,
    NATIVE_OFFICE_COLLABORATION_PROJECTION_VERSION,
};

const MAX_PROJECTED_DOCUMENT_PARAGRAPHS: usize = 1_048_576;
const MAX_PROJECTED_TEXT_BYTES: usize = 64 * 1024 * 1024;
const MAX_PROJECTED_ANCESTOR_DEPTH: usize = 64;

pub(super) fn project_collaboration_document(
    document: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    sequence: u64,
) -> UseResult<NativeOfficeCollaborationProjection> {
    let inspection = inspect_document(document, manifest)?;
    let initialized = inspection
        .metadata
        .as_ref()
        .is_some_and(|metadata| metadata.initialized);
    if !initialized {
        return Err(collaboration_error(
            "office.collaboration.not_initialized",
            "The Office collaboration document has not been initialized.",
        ));
    }
    if inspection.pending_updates {
        return Err(collaboration_error(
            "office.collaboration.projection_incomplete",
            "The Office collaboration document still has causally missing updates and cannot be projected safely.",
        )
        .with_suggestion("Deliver the missing Yjs updates before reading agent-visible content."));
    }

    let state_vector = {
        let transaction = document.transact();
        canonical_state_vector(&transaction.state_vector())
    };
    let content = match manifest.kind {
        NativeOfficeCollaborationArtifactKind::Markdown => {
            let transaction = document.transact();
            project_markdown(&transaction, manifest)?
        }
        NativeOfficeCollaborationArtifactKind::Document => {
            let transaction = document.transact();
            project_document_content(&transaction, manifest)?
        }
        NativeOfficeCollaborationArtifactKind::Spreadsheet => {
            project_spreadsheet_content(document, manifest)?
        }
        NativeOfficeCollaborationArtifactKind::Presentation => {
            project_presentation_content(document, manifest)?
        }
        NativeOfficeCollaborationArtifactKind::Pdf => project_pdf_content(document, manifest)?,
    };

    Ok(NativeOfficeCollaborationProjection {
        schema: NATIVE_OFFICE_COLLABORATION_PROJECTION_SCHEMA.to_owned(),
        version: NATIVE_OFFICE_COLLABORATION_PROJECTION_VERSION,
        artifact_id: manifest.artifact_id.clone(),
        artifact_kind: manifest.kind,
        sequence,
        state_vector_sha256: sha256_hex(&state_vector),
        state_vector,
        content,
    })
}

fn project_document_content<T: ReadTxn>(
    transaction: &T,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<NativeOfficeCollaborationProjectedContent> {
    let root = format!("{}.document.content", manifest.namespace);
    let fragment = match transaction.get(&root) {
        Some(Out::YXmlFragment(fragment)) => fragment,
        Some(_) => return Err(invalid_root(&root, "Y.XmlFragment")),
        None => return Err(missing_root(&root)),
    };
    let mut paragraphs = Vec::new();
    let mut plain_text = String::new();
    for node in fragment.successors(transaction) {
        let XmlOut::Element(element) = node else {
            continue;
        };
        if !is_identity_paragraph_tag(element.tag()) {
            continue;
        }
        if paragraphs.len() >= MAX_PROJECTED_DOCUMENT_PARAGRAPHS {
            return Err(collaboration_error(
                "office.collaboration.projection_too_large",
                "The collaborative Document contains too many paragraphs for a bounded native projection.",
            )
            .with_detail(
                "maximumParagraphs",
                MAX_PROJECTED_DOCUMENT_PARAGRAPHS as u64,
            ));
        }
        let projected = project_paragraph(&element, transaction, paragraphs.len() + 1)?;
        if !plain_text.is_empty() {
            plain_text.push('\n');
        }
        plain_text.push_str(&projected.text);
        ensure_projected_text_size(plain_text.len())?;
        paragraphs.push(projected);
    }
    if paragraphs.is_empty() {
        return Err(collaboration_error(
            "office.collaboration.content_invalid",
            "The initialized Document collaboration contains no supported text blocks.",
        ));
    }

    let options_root = format!("{}.document.options", manifest.namespace);
    let options = match transaction.get(&options_root) {
        Some(Out::YMap(options)) => options,
        Some(_) => return Err(invalid_root(&options_root, "Y.Map")),
        None => return Err(missing_root(&options_root)),
    };
    let page_color = optional_string(&options, transaction, "pageColor")?;
    let track_changes = optional_bool(&options, transaction, "trackChanges")?;
    let comments = project_document_comments(transaction, manifest, &fragment)?;
    let suggestions = project_document_suggestions(transaction, &fragment)?;
    let change_decisions = project_document_change_decisions(transaction, manifest)?;
    Ok(NativeOfficeCollaborationProjectedContent::Document {
        plain_text,
        paragraphs,
        comments,
        suggestions,
        change_decisions,
        page_color,
        track_changes,
    })
}

fn project_paragraph<T: ReadTxn>(
    paragraph: &yrs::XmlElementRef,
    transaction: &T,
    ordinal: usize,
) -> UseResult<NativeOfficeCollaborationDocumentParagraph> {
    let paragraph_id = document_identity_attribute(paragraph, transaction, PARAGRAPH_ID_ATTRIBUTE)?;
    let text_id = document_identity_attribute(paragraph, transaction, TEXT_ID_ATTRIBUTE)?;
    if paragraph_id.is_some() != text_id.is_some() {
        return Err(collaboration_error(
            "office.collaboration.content_invalid",
            "A shared Document paragraph has an incomplete Word paragraph identity.",
        ));
    }

    let mut text = String::new();
    let mut text_nodes = 0_usize;
    let mut has_inline_objects = false;
    let mut has_review_marks = false;
    for child in paragraph.children(transaction) {
        let XmlOut::Text(child) = child else {
            has_inline_objects = true;
            continue;
        };
        text_nodes += 1;
        for chunk in child.diff(transaction, |_| ()) {
            if chunk
                .attributes
                .as_ref()
                .is_some_and(|attributes| attributes.keys().any(|key| is_review_attribute(key)))
            {
                has_review_marks = true;
            }
            match chunk.insert {
                Out::Any(Any::String(value)) => text.push_str(&value),
                _ => has_inline_objects = true,
            }
        }
    }
    ensure_projected_text_size(text.len())?;
    let container_path = container_path(paragraph)?;
    let replaceable = paragraph.tag().as_ref() == "paragraph"
        && paragraph_id.is_some()
        && text_id.is_some()
        && text_nodes <= 1
        && !has_inline_objects
        && !has_review_marks;
    Ok(NativeOfficeCollaborationDocumentParagraph {
        ordinal: u32::try_from(ordinal).map_err(|_| {
            collaboration_error(
                "office.collaboration.projection_too_large",
                "The Document paragraph ordinal exceeds the supported range.",
            )
        })?,
        node_type: paragraph.tag().to_string(),
        paragraph_id,
        text_id,
        container_path,
        text: text.clone(),
        start_utf16: 0,
        end_utf16: utf16_len(&text)?,
        replaceable,
        has_inline_objects,
        has_review_marks,
    })
}

fn project_markdown<T: ReadTxn>(
    transaction: &T,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<NativeOfficeCollaborationProjectedContent> {
    let root = format!("{}.markdown.source", manifest.namespace);
    let source = match transaction.get(&root) {
        Some(Out::YText(text)) => text.get_string(transaction),
        Some(_) => return Err(invalid_root(&root, "Y.Text")),
        None => return Err(missing_root(&root)),
    };
    ensure_projected_text_size(source.len())?;
    let slices = markdown_slices(&source)?;
    Ok(NativeOfficeCollaborationProjectedContent::Markdown { source, slices })
}

fn markdown_slices(source: &str) -> UseResult<Vec<NativeOfficeCollaborationMarkdownSlice>> {
    let mut slices = Vec::new();
    let mut index = 0_u32;
    let mut cursor = 0_u32;
    let mut slice_start = 0_u32;
    let mut text = String::new();
    for character in source.chars() {
        let width = u32::try_from(character.len_utf16()).map_err(|_| utf16_overflow())?;
        if character == '\n' {
            push_markdown_slice(&mut slices, index, slice_start, cursor, &text)?;
            text.clear();
            index = index.checked_add(1).ok_or_else(utf16_overflow)?;
            cursor = cursor.checked_add(width).ok_or_else(utf16_overflow)?;
            slice_start = cursor;
        } else {
            text.push(character);
            cursor = cursor.checked_add(width).ok_or_else(utf16_overflow)?;
        }
        if slices.len() > MAX_PROJECTED_DOCUMENT_PARAGRAPHS {
            return Err(too_many_slices());
        }
    }
    push_markdown_slice(&mut slices, index, slice_start, cursor, &text)?;
    Ok(slices)
}

fn push_markdown_slice(
    slices: &mut Vec<NativeOfficeCollaborationMarkdownSlice>,
    index: u32,
    start_utf16: u32,
    end_utf16: u32,
    text: &str,
) -> UseResult<()> {
    if slices.len() >= MAX_PROJECTED_DOCUMENT_PARAGRAPHS {
        return Err(too_many_slices());
    }
    ensure_projected_text_size(text.len())?;
    slices.push(NativeOfficeCollaborationMarkdownSlice {
        index,
        start_utf16,
        end_utf16,
        text: text.to_owned(),
    });
    Ok(())
}

fn utf16_len(value: &str) -> UseResult<u32> {
    u32::try_from(value.encode_utf16().count()).map_err(|_| utf16_overflow())
}

fn utf16_overflow() -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.projection_too_large",
        "The native Office text projection exceeds the supported UTF-16 offset range.",
    )
}

fn too_many_slices() -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.projection_too_large",
        "The collaborative Markdown contains too many slices for a bounded native projection.",
    )
    .with_detail("maximumSlices", MAX_PROJECTED_DOCUMENT_PARAGRAPHS as u64)
}

fn container_path(paragraph: &yrs::XmlElementRef) -> UseResult<Vec<String>> {
    let mut path = Vec::new();
    let mut current = paragraph.parent();
    for _ in 0..MAX_PROJECTED_ANCESTOR_DEPTH {
        let Some(XmlOut::Element(parent)) = current else {
            current = None;
            break;
        };
        path.push(parent.tag().to_string());
        current = parent.parent();
    }
    if matches!(current, Some(XmlOut::Element(_))) {
        return Err(collaboration_error(
            "office.collaboration.content_invalid",
            "A shared Document paragraph exceeds the supported ancestor depth.",
        ));
    }
    path.reverse();
    Ok(path)
}

fn optional_string<T: ReadTxn>(
    map: &yrs::MapRef,
    transaction: &T,
    key: &str,
) -> UseResult<Option<String>> {
    match map.get(transaction, key) {
        Some(Out::Any(Any::String(value))) => Ok(Some(value.to_string())),
        Some(_) => Err(invalid_option(key)),
        None => Ok(None),
    }
}

fn optional_bool<T: ReadTxn>(
    map: &yrs::MapRef,
    transaction: &T,
    key: &str,
) -> UseResult<Option<bool>> {
    match map.get(transaction, key) {
        Some(Out::Any(Any::Bool(value))) => Ok(Some(value)),
        Some(_) => Err(invalid_option(key)),
        None => Ok(None),
    }
}

fn ensure_projected_text_size(bytes: usize) -> UseResult<()> {
    if bytes <= MAX_PROJECTED_TEXT_BYTES {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.projection_too_large",
        "The native Office text projection exceeds the bounded output size.",
    )
    .with_detail("bytes", bytes as u64)
    .with_detail("maximumBytes", MAX_PROJECTED_TEXT_BYTES as u64))
}

fn is_review_attribute(value: &str) -> bool {
    matches!(
        value.split_once("--").map_or(value, |(name, _)| name),
        "documentComment" | "documentChange"
    )
}

fn missing_root(root: &str) -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.content_invalid",
        format!("The initialized Office collaboration root '{root}' is missing."),
    )
}

fn invalid_root(root: &str, expected: &str) -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.root_invalid",
        format!("The Office collaboration root '{root}' must be a {expected}."),
    )
}

fn invalid_option(key: &str) -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.content_invalid",
        format!("The shared Document option '{key}' is invalid."),
    )
}
