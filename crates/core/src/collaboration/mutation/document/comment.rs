mod authorization;
mod json;
mod state;

use std::collections::HashMap;
use std::sync::Arc;

use a3s_use_core::UseResult;
use serde_json::{json, Value as JsonValue};
use yrs::types::Attrs;
use yrs::{
    Any, Array, ArrayPrelim, GetString, Map, MapPrelim, Out, ReadTxn, Text, Transact, XmlFragment,
    XmlOut, XmlTextRef,
};

use self::json::canonical_json;
pub(in crate::collaboration) use self::state::project_document_comments;
use self::state::{
    collect_comment_anchors, collect_comment_mark_segments, comment_attributes, read_comment_state,
    CommentAnchor,
};
use super::super::super::{
    collaboration_error, NativeOfficeCollaborationManifest, NativeOfficeCollaborationMode,
    NativeOfficeCollaborationMutation,
};
use super::super::{is_utf16_boundary, utf16_len};
use super::identity::{
    document_identity_attribute, is_identity_paragraph_tag, validate_paragraph_id_input,
    PARAGRAPH_ID_ATTRIBUTE, TEXT_ID_ATTRIBUTE,
};
pub(in crate::collaboration) use authorization::validate_authorized_comment_update;

const MAX_COMMENT_FIELD_UTF16: u32 = 1_048_576;

pub(super) fn validate_comment_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::DocumentCommentCreate {
            comment_id,
            paragraph_id,
            expected_text_id,
            start_utf16,
            end_utf16,
            expected_text,
            author,
            created_at,
            text,
        } => {
            validate_identifier(comment_id, "commentId", "Document comment")?;
            validate_paragraph_id_input(paragraph_id, "paragraphId")?;
            validate_paragraph_id_input(expected_text_id, "expectedTextId")?;
            if start_utf16 >= end_utf16 {
                return Err(collaboration_error(
                    "office.collaboration.mutation_invalid",
                    "A Document comment selection requires startUtf16 to be smaller than endUtf16.",
                ));
            }
            validate_non_empty_text(expected_text, "expectedText", "selected text")?;
            validate_non_empty_text(author, "author", "author")?;
            validate_timestamp(created_at, "createdAt", "comment")?;
            validate_non_empty_text(text, "text", "comment text")
        }
        NativeOfficeCollaborationMutation::DocumentCommentReply {
            comment_id,
            reply_id,
            author,
            created_at,
            text,
        } => {
            validate_identifier(comment_id, "commentId", "Document comment")?;
            validate_identifier(reply_id, "replyId", "Document comment reply")?;
            validate_non_empty_text(author, "author", "author")?;
            validate_timestamp(created_at, "createdAt", "comment reply")?;
            validate_non_empty_text(text, "text", "comment reply text")
        }
        NativeOfficeCollaborationMutation::DocumentCommentSetResolved { comment_id, .. } => {
            validate_identifier(comment_id, "commentId", "Document comment")
        }
        NativeOfficeCollaborationMutation::DocumentCommentDelete {
            comment_id,
            reply_id,
        } => {
            validate_identifier(comment_id, "commentId", "Document comment")?;
            if let Some(reply_id) = reply_id {
                validate_identifier(reply_id, "replyId", "Document comment reply")?;
            }
            Ok(())
        }
        _ => Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "The supplied mutation is not a Document comment mutation.",
        )),
    }
}

pub(super) fn apply_comment_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::DocumentCommentCreate {
            comment_id,
            paragraph_id,
            expected_text_id,
            start_utf16,
            end_utf16,
            expected_text,
            author,
            created_at,
            text,
        } => create_comment(
            doc,
            manifest,
            comment_id,
            paragraph_id,
            expected_text_id,
            *start_utf16,
            *end_utf16,
            expected_text,
            author,
            created_at,
            text,
        ),
        NativeOfficeCollaborationMutation::DocumentCommentReply {
            comment_id,
            reply_id,
            author,
            created_at,
            text,
        } => create_reply(
            doc, manifest, comment_id, reply_id, author, created_at, text,
        ),
        NativeOfficeCollaborationMutation::DocumentCommentSetResolved {
            comment_id,
            resolved,
        } => set_resolved(doc, manifest, comment_id, *resolved),
        NativeOfficeCollaborationMutation::DocumentCommentDelete {
            comment_id,
            reply_id,
        } => delete_comment_record(doc, manifest, comment_id, reply_id.as_deref()),
        _ => Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "The supplied mutation is not a Document comment mutation.",
        )),
    }
}

#[allow(clippy::too_many_arguments)]
fn create_comment(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    comment_id: &str,
    paragraph_id: &str,
    expected_text_id: &str,
    start_utf16: u32,
    end_utf16: u32,
    expected_text: &str,
    author: &str,
    created_at: &str,
    text: &str,
) -> UseResult<()> {
    let fragment =
        doc.get_or_insert_xml_fragment(format!("{}.document.content", manifest.namespace));
    let transaction = doc.transact();
    let state = read_comment_state(&transaction, manifest)?;
    let anchors = collect_comment_anchors(&transaction, &fragment)?;
    let fingerprint =
        comment_fingerprint(comment_id, &manifest.actor_id, author, created_at, text)?;
    if state
        .claims
        .claim_for("comment", None, comment_id)
        .is_some_and(|claim| claim != fingerprint)
    {
        return Err(identity_conflict(format!(
            "Document comment ID '{comment_id}' was already claimed by a different record."
        )));
    }
    if let Some(existing) = state.records.get(comment_id) {
        if existing.actor_id.as_deref() != Some(manifest.actor_id.as_str())
            || existing.author != author
            || existing.date != created_at
            || existing.text != text
        {
            return Err(identity_conflict(format!(
                "Document comment ID '{comment_id}' already belongs to a different record."
            )));
        }
        let target = resolve_comment_target(
            &transaction,
            &fragment,
            paragraph_id,
            expected_text_id,
            start_utf16,
            end_utf16,
            expected_text,
            Some(comment_id),
        )?;
        let existing_anchors = anchors.get(comment_id).cloned().unwrap_or_default();
        if existing_anchors == vec![target.anchor] {
            return Ok(());
        }
        return Err(identity_conflict(format!(
            "Document comment ID '{comment_id}' already owns a different anchor."
        )));
    }
    let target = resolve_comment_target(
        &transaction,
        &fragment,
        paragraph_id,
        expected_text_id,
        start_utf16,
        end_utf16,
        expected_text,
        None,
    )?;
    if anchors.contains_key(comment_id) {
        return Err(identity_conflict(format!(
            "Document comment ID '{comment_id}' already owns a dangling anchor."
        )));
    }
    let append_claim = state
        .claims
        .claim_for("comment", None, comment_id)
        .is_none();
    drop(transaction);

    let comments = doc.get_or_insert_map(format!("{}.document.comments", manifest.namespace));
    let order = doc.get_or_insert_array(format!("{}.document.comment-order", manifest.namespace));
    let claims = doc.get_or_insert_array(format!("{}.document.record-claims", manifest.namespace));
    let options = doc.get_or_insert_map(format!("{}.document.options", manifest.namespace));
    let claim = encoded_claim("comment", comment_id, None, &fingerprint)?;
    let mut transaction = doc.transact_mut();
    let record = comments.insert(&mut transaction, comment_id, MapPrelim::default());
    record.insert(&mut transaction, "id", comment_id);
    record.insert(&mut transaction, "actorId", manifest.actor_id.as_str());
    record.insert(&mut transaction, "author", author);
    record.insert(&mut transaction, "date", created_at);
    record.insert(&mut transaction, "text", text);
    record.insert(&mut transaction, "resolved", false);
    record.insert(&mut transaction, "replies", MapPrelim::default());
    record.insert(&mut transaction, "replyOrder", ArrayPrelim::default());
    order.push_back(&mut transaction, comment_id);
    options.insert(&mut transaction, "commentsPresent", true);
    if append_claim {
        claims.push_back(&mut transaction, claim);
    }
    target.text.format(
        &mut transaction,
        start_utf16,
        end_utf16 - start_utf16,
        comment_mark_attributes(comment_id),
    );
    drop(transaction);
    validate_written_comments(doc, manifest)
}

fn create_reply(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    comment_id: &str,
    reply_id: &str,
    author: &str,
    created_at: &str,
    text: &str,
) -> UseResult<()> {
    let transaction = doc.transact();
    let state = read_comment_state(&transaction, manifest)?;
    let comment = state.records.get(comment_id).ok_or_else(|| {
        match_conflict(format!(
            "Document comment ID '{comment_id}' does not exist for this reply."
        ))
    })?;
    let fingerprint = reply_fingerprint(reply_id, &manifest.actor_id, author, created_at, text)?;
    if state
        .claims
        .claim_for("comment-reply", Some(comment_id), reply_id)
        .is_some_and(|claim| claim != fingerprint)
    {
        return Err(identity_conflict(format!(
            "Document reply ID '{reply_id}' was already claimed by a different record in comment '{comment_id}'."
        )));
    }
    if let Some(existing) = comment.reply_records.get(reply_id) {
        if existing.actor_id.as_deref() == Some(manifest.actor_id.as_str())
            && existing.author == author
            && existing.date == created_at
            && existing.text == text
        {
            return Ok(());
        }
        return Err(identity_conflict(format!(
            "Document reply ID '{reply_id}' already belongs to a different record in comment '{comment_id}'."
        )));
    }
    let append_claim = state
        .claims
        .claim_for("comment-reply", Some(comment_id), reply_id)
        .is_none();
    let replies = comment.replies.clone();
    let reply_order = comment.reply_order_root.clone();
    let claims = state.claims_root.clone();
    drop(transaction);

    let claim = encoded_claim("comment-reply", reply_id, Some(comment_id), &fingerprint)?;
    let mut transaction = doc.transact_mut();
    let record = replies.insert(&mut transaction, reply_id, MapPrelim::default());
    record.insert(&mut transaction, "id", reply_id);
    record.insert(&mut transaction, "actorId", manifest.actor_id.as_str());
    record.insert(&mut transaction, "author", author);
    record.insert(&mut transaction, "date", created_at);
    record.insert(&mut transaction, "text", text);
    reply_order.push_back(&mut transaction, reply_id);
    if append_claim {
        claims.push_back(&mut transaction, claim);
    }
    drop(transaction);
    validate_written_comments(doc, manifest)
}

fn set_resolved(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    comment_id: &str,
    resolved: bool,
) -> UseResult<()> {
    let transaction = doc.transact();
    let state = read_comment_state(&transaction, manifest)?;
    let comment = state.records.get(comment_id).ok_or_else(|| {
        match_conflict(format!(
            "Document comment ID '{comment_id}' does not exist for a resolution change."
        ))
    })?;
    if comment.resolved == resolved {
        return Ok(());
    }
    let record = comment.record.clone();
    drop(transaction);
    record.insert(&mut doc.transact_mut(), "resolved", resolved);
    validate_written_comments(doc, manifest)
}

fn delete_comment_record(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    comment_id: &str,
    reply_id: Option<&str>,
) -> UseResult<()> {
    let fragment =
        doc.get_or_insert_xml_fragment(format!("{}.document.content", manifest.namespace));
    let transaction = doc.transact();
    let state = read_comment_state(&transaction, manifest)?;
    let comment = state.records.get(comment_id).ok_or_else(|| {
        match_conflict(format!(
            "Document comment ID '{comment_id}' does not exist for deletion."
        ))
    })?;
    if let Some(reply_id) = reply_id {
        let reply = comment.reply_records.get(reply_id).ok_or_else(|| {
            match_conflict(format!(
                "Document reply ID '{reply_id}' does not exist in comment '{comment_id}'."
            ))
        })?;
        assert_delete_owner(manifest, reply.actor_id.as_deref(), "comment reply")?;
        let replies = comment.replies.clone();
        let order = comment.reply_order_root.clone();
        drop(transaction);
        let mut transaction = doc.transact_mut();
        replies.remove(&mut transaction, reply_id);
        remove_order_entries(&order, &mut transaction, reply_id);
        drop(transaction);
        return validate_written_comments(doc, manifest);
    }

    assert_delete_owner(manifest, comment.actor_id.as_deref(), "comment")?;
    let comments = state.comments.clone();
    let order = state.order_root.clone();
    let marks = collect_comment_mark_segments(&transaction, &fragment)?;
    let mut rewrites = Vec::<CommentMarkRewrite>::new();
    for target in marks
        .iter()
        .filter(|segment| segment.comment_id == comment_id)
    {
        if rewrites
            .iter()
            .any(|rewrite| rewrite.text == target.text && rewrite.attribute == target.attribute)
        {
            continue;
        }
        rewrites.push(CommentMarkRewrite {
            text: target.text.clone(),
            attribute: target.attribute.clone(),
            length_utf16: target.text.len(&transaction),
            retained: marks
                .iter()
                .filter(|segment| {
                    segment.text == target.text
                        && segment.attribute == target.attribute
                        && segment.comment_id != comment_id
                })
                .map(|segment| {
                    (
                        segment.start_utf16,
                        segment.end_utf16,
                        segment.value.clone(),
                    )
                })
                .collect(),
        });
    }
    drop(transaction);
    let mut transaction = doc.transact_mut();
    comments.remove(&mut transaction, comment_id);
    remove_order_entries(&order, &mut transaction, comment_id);
    // Clearing a remote Y.XmlText format from an interior split can panic in
    // Yrs 0.27. Clear the touched attribute from the whole text at offset 0,
    // then restore every other comment span that used the same attribute key.
    for rewrite in &rewrites {
        if rewrite.length_utf16 > 0 {
            rewrite.text.format(
                &mut transaction,
                0,
                rewrite.length_utf16,
                Attrs::from([(rewrite.attribute.clone().into(), Any::Null)]),
            );
        }
    }
    for rewrite in rewrites {
        for (start, end, value) in rewrite.retained {
            rewrite.text.format(
                &mut transaction,
                start,
                end - start,
                Attrs::from([(rewrite.attribute.clone().into(), value)]),
            );
        }
    }
    drop(transaction);
    validate_written_comments(doc, manifest)
}

struct CommentMarkRewrite {
    text: XmlTextRef,
    attribute: String,
    length_utf16: u32,
    retained: Vec<(u32, u32, Any)>,
}

struct CommentTarget {
    text: XmlTextRef,
    anchor: CommentAnchor,
}

#[allow(clippy::too_many_arguments)]
fn resolve_comment_target<T: ReadTxn>(
    transaction: &T,
    fragment: &yrs::XmlFragmentRef,
    paragraph_id: &str,
    expected_text_id: &str,
    start_utf16: u32,
    end_utf16: u32,
    expected_text: &str,
    allowed_comment_id: Option<&str>,
) -> UseResult<CommentTarget> {
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
            return Err(invalid_shared_comments(format!(
                "The shared Document repeats paragraph ID '{paragraph_id}'."
            )));
        }
        found = Some(paragraph);
    }
    let paragraph = found.ok_or_else(|| {
        match_conflict(format!(
            "Document paragraph ID '{paragraph_id}' does not exist for this comment."
        ))
    })?;
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
                return Err(collaboration_error(
                    "office.collaboration.mutation_structure_conflict",
                    "A native Document comment can target only a paragraph with one text node and no inline objects.",
                ))
            }
        }
    }
    let text = text.ok_or_else(|| {
        collaboration_error(
            "office.collaboration.mutation_structure_conflict",
            "A native Document comment cannot target an empty paragraph without a text node.",
        )
    })?;
    let current = text.get_string(transaction);
    let current_len = utf16_len(&current)?;
    if end_utf16 > current_len
        || !is_utf16_boundary(&current, start_utf16)
        || !is_utf16_boundary(&current, end_utf16)
    {
        return Err(range_conflict(start_utf16, end_utf16, current_len));
    }
    let selected = utf16_slice(&current, start_utf16, end_utf16)
        .ok_or_else(|| range_conflict(start_utf16, end_utf16, current_len))?;
    if selected != expected_text {
        return Err(match_conflict(format!(
            "The current Document comment selection is '{selected}', not the expected text."
        )));
    }
    let mut cursor = 0_u32;
    for chunk in text.diff(transaction, |_| ()) {
        let length = match chunk.insert {
            Out::Any(Any::String(value)) => utf16_len(&value)?,
            _ => {
                return Err(collaboration_error(
                    "office.collaboration.mutation_structure_conflict",
                    "A native Document comment selection cannot cross an inline object.",
                ))
            }
        };
        let end = cursor.checked_add(length).ok_or_else(|| {
            collaboration_error(
                "office.collaboration.mutation_too_large",
                "The Document comment selection offset is too large.",
            )
        })?;
        if cursor < end_utf16 && end > start_utf16 {
            let conflicting = comment_attributes(chunk.attributes.as_deref())?
                .into_iter()
                .any(|(_, id)| allowed_comment_id != Some(id.as_str()));
            if conflicting {
                return Err(match_conflict(
                    "The Document comment selection overlaps an existing comment anchor."
                        .to_owned(),
                ));
            }
        }
        cursor = end;
    }
    Ok(CommentTarget {
        text,
        anchor: CommentAnchor {
            paragraph_id: Some(paragraph_id.to_owned()),
            text_id: Some(expected_text_id.to_owned()),
            start_utf16,
            end_utf16,
            text: selected.to_owned(),
        },
    })
}

fn validate_written_comments(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<()> {
    let fragment =
        doc.get_or_insert_xml_fragment(format!("{}.document.content", manifest.namespace));
    let transaction = doc.transact();
    read_comment_state(&transaction, manifest)?;
    collect_comment_anchors(&transaction, &fragment)?;
    Ok(())
}

fn comment_fingerprint(
    id: &str,
    actor_id: &str,
    author: &str,
    date: &str,
    text: &str,
) -> UseResult<String> {
    canonical_json(&json!({
        "actorId": actor_id,
        "author": author,
        "date": date,
        "id": id,
        "resolved": false,
        "text": text,
    }))
}

fn reply_fingerprint(
    id: &str,
    actor_id: &str,
    author: &str,
    date: &str,
    text: &str,
) -> UseResult<String> {
    canonical_json(&json!({
        "actorId": actor_id,
        "author": author,
        "date": date,
        "id": id,
        "text": text,
    }))
}

fn encoded_claim(
    kind: &str,
    id: &str,
    parent_id: Option<&str>,
    fingerprint: &str,
) -> UseResult<String> {
    let mut claim = serde_json::Map::new();
    claim.insert(
        "fingerprint".to_owned(),
        JsonValue::String(fingerprint.to_owned()),
    );
    claim.insert("id".to_owned(), JsonValue::String(id.to_owned()));
    claim.insert("kind".to_owned(), JsonValue::String(kind.to_owned()));
    if let Some(parent_id) = parent_id {
        claim.insert(
            "parentId".to_owned(),
            JsonValue::String(parent_id.to_owned()),
        );
    }
    canonical_json(&JsonValue::Object(claim))
}

fn comment_mark_attributes(comment_id: &str) -> Attrs {
    let mut value = HashMap::new();
    value.insert("id".to_owned(), Any::String(comment_id.into()));
    Attrs::from([("documentComment".into(), Any::Map(Arc::new(value)))])
}

fn remove_order_entries(
    order: &yrs::ArrayRef,
    transaction: &mut yrs::TransactionMut<'_>,
    id: &str,
) {
    for index in (0..order.len(transaction)).rev() {
        if matches!(order.get(transaction, index), Some(Out::Any(Any::String(value))) if value.as_ref() == id)
        {
            order.remove_range(transaction, index, 1);
        }
    }
}

fn assert_delete_owner(
    manifest: &NativeOfficeCollaborationManifest,
    actor_id: Option<&str>,
    label: &str,
) -> UseResult<()> {
    if manifest.mode != NativeOfficeCollaborationMode::Comment
        || actor_id == Some(manifest.actor_id.as_str())
    {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.permission_denied",
        format!("Comment mode can delete only the collaboration actor's own Document {label}."),
    ))
}

fn validate_identifier(value: &str, field: &str, label: &str) -> UseResult<()> {
    if !value.is_empty() && value == value.trim() && value.encode_utf16().count() <= 256 {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.mutation_invalid",
        format!("A native {label} field '{field}' must contain 1 to 256 non-padded characters."),
    ))
}

pub(super) fn validate_shared_identifier(value: &str, label: &str) -> UseResult<()> {
    if !value.is_empty() && value == value.trim() && value.encode_utf16().count() <= 256 {
        return Ok(());
    }
    Err(invalid_shared_comments(format!(
        "The shared Document {label} ID is invalid."
    )))
}

fn validate_non_empty_text(value: &str, field: &str, label: &str) -> UseResult<()> {
    let length = utf16_len(value)?;
    if !value.trim().is_empty() && length <= MAX_COMMENT_FIELD_UTF16 {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.mutation_invalid",
        format!(
            "A native Document comment field '{field}' must contain non-whitespace {label} no longer than {MAX_COMMENT_FIELD_UTF16} UTF-16 code units."
        ),
    ))
}

fn validate_timestamp(value: &str, field: &str, label: &str) -> UseResult<()> {
    if valid_canonical_utc_timestamp(value) {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.mutation_invalid",
        format!(
            "A native Document {label} timestamp must use canonical UTC form YYYY-MM-DDTHH:mm:ss.sssZ."
        ),
    )
    .with_detail(field, value.to_owned()))
}

fn valid_canonical_utc_timestamp(value: &str) -> bool {
    let bytes = value.as_bytes();
    bytes.len() == 24
        && bytes[10] == b'T'
        && bytes[13] == b':'
        && bytes[16] == b':'
        && bytes[19] == b'.'
        && bytes[23] == b'Z'
        && valid_date(&bytes[..10])
        && valid_time(&bytes[11..19])
        && bytes[20..23].iter().all(u8::is_ascii_digit)
}

fn valid_date(value: &[u8]) -> bool {
    if value.len() != 10 || value[4] != b'-' || value[7] != b'-' {
        return false;
    }
    let Some(year) = parse_digits(&value[..4]) else {
        return false;
    };
    let Some(month) = parse_digits(&value[5..7]) else {
        return false;
    };
    let Some(day) = parse_digits(&value[8..]) else {
        return false;
    };
    let days = match month {
        2 if year % 400 == 0 || (year % 4 == 0 && year % 100 != 0) => 29,
        2 => 28,
        4 | 6 | 9 | 11 => 30,
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        _ => return false,
    };
    (1..=days).contains(&day)
}

fn valid_time(value: &[u8]) -> bool {
    value.len() == 8
        && value[2] == b':'
        && value[5] == b':'
        && parse_digits(&value[..2]).is_some_and(|hour| hour <= 23)
        && parse_digits(&value[3..5]).is_some_and(|minute| minute <= 59)
        && parse_digits(&value[6..]).is_some_and(|second| second <= 59)
}

fn parse_digits(value: &[u8]) -> Option<u32> {
    if value.is_empty() || !value.iter().all(u8::is_ascii_digit) {
        return None;
    }
    Some(
        value
            .iter()
            .fold(0, |result, digit| result * 10 + u32::from(*digit - b'0')),
    )
}

fn utf16_slice(value: &str, start: u32, end: u32) -> Option<&str> {
    let mut cursor = 0_u32;
    let mut start_byte = None;
    let mut end_byte = None;
    if start == 0 {
        start_byte = Some(0);
    }
    if end == 0 {
        end_byte = Some(0);
    }
    for (byte_index, character) in value.char_indices() {
        cursor = cursor.checked_add(character.len_utf16() as u32)?;
        let next_byte = byte_index + character.len_utf8();
        if cursor == start {
            start_byte = Some(next_byte);
        }
        if cursor == end {
            end_byte = Some(next_byte);
            break;
        }
        if cursor > end {
            return None;
        }
    }
    Some(&value[start_byte?..end_byte?])
}

pub(super) fn invalid_shared_comments(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.content_invalid", message)
}

fn identity_conflict(message: String) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_identity_conflict", message).with_suggestion(
        "Read the latest Document comment projection and choose a fresh stable ID before retrying.",
    )
}

fn match_conflict(message: String) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_match_conflict", message).with_suggestion(
        "Read the latest Document comment projection and retry against its exact identity and UTF-16 anchor.",
    )
}

fn range_conflict(start: u32, end: u32, length: u32) -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.mutation_range_invalid",
        "The Document comment selection is outside the current paragraph or splits a UTF-16 surrogate pair.",
    )
    .with_detail("startUtf16", start as u64)
    .with_detail("endUtf16", end as u64)
    .with_detail("lengthUtf16", length as u64)
}
