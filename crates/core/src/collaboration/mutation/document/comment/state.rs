use std::collections::{HashMap, HashSet};

use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;
use yrs::types::Attrs;
use yrs::{Any, Array, ArrayRef, Map, MapRef, Out, ReadTxn, Text, XmlFragment, XmlOut, XmlTextRef};

use super::json::canonical_json;
use super::{invalid_shared_comments, validate_shared_identifier};
use crate::collaboration::mutation::document::identity::{
    document_identity_attribute, is_identity_paragraph_tag, PARAGRAPH_ID_ATTRIBUTE,
    TEXT_ID_ATTRIBUTE,
};
use crate::collaboration::{
    NativeOfficeCollaborationDocumentComment, NativeOfficeCollaborationDocumentCommentAnchor,
    NativeOfficeCollaborationDocumentCommentReply, NativeOfficeCollaborationManifest,
};

const MAX_DOCUMENT_COMMENTS: usize = 1_048_576;
const MAX_DOCUMENT_COMMENT_REPLIES: usize = 1_048_576;
const MAX_DOCUMENT_COMMENT_CLAIMS: usize = 1_048_576;
const MAX_DOCUMENT_COMMENT_TEXT_BYTES: usize = 64 * 1024 * 1024;

pub(super) struct DocumentCommentsState {
    pub(super) comments: MapRef,
    pub(super) order_root: ArrayRef,
    pub(super) claims_root: ArrayRef,
    pub(super) order: Vec<String>,
    pub(super) records: HashMap<String, DocumentCommentRecord>,
    pub(super) claims: DocumentRecordClaims,
}

pub(super) struct DocumentCommentRecord {
    pub(super) record: MapRef,
    pub(super) id: String,
    pub(super) actor_id: Option<String>,
    pub(super) author: String,
    pub(super) date: String,
    pub(super) text: String,
    pub(super) resolved: bool,
    pub(super) replies: MapRef,
    pub(super) reply_order_root: ArrayRef,
    pub(super) reply_order: Vec<String>,
    pub(super) reply_records: HashMap<String, DocumentCommentReplyRecord>,
}

pub(super) struct DocumentCommentReplyRecord {
    pub(super) id: String,
    pub(super) actor_id: Option<String>,
    pub(super) author: String,
    pub(super) date: String,
    pub(super) text: String,
}

pub(in crate::collaboration) struct DocumentRecordClaims {
    by_identity: HashMap<(String, Option<String>, String), String>,
}

impl DocumentRecordClaims {
    pub(in crate::collaboration) fn claim_for(
        &self,
        kind: &str,
        parent_id: Option<&str>,
        id: &str,
    ) -> Option<&str> {
        self.by_identity
            .get(&(kind.to_owned(), parent_id.map(str::to_owned), id.to_owned()))
            .map(String::as_str)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct CommentAnchor {
    pub(super) paragraph_id: Option<String>,
    pub(super) text_id: Option<String>,
    pub(super) start_utf16: u32,
    pub(super) end_utf16: u32,
    pub(super) text: String,
}

#[derive(Debug, Clone)]
pub(super) struct CommentMarkSegment {
    pub(super) text: XmlTextRef,
    pub(super) attribute: String,
    pub(super) value: Any,
    pub(super) start_utf16: u32,
    pub(super) end_utf16: u32,
    pub(super) comment_id: String,
}

pub(super) fn read_comment_state<T: ReadTxn>(
    transaction: &T,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<DocumentCommentsState> {
    let options = required_root_map(transaction, manifest, "document.options")?;
    let comments = required_root_map(transaction, manifest, "document.comments")?;
    let order_root = required_root_array(transaction, manifest, "document.comment-order")?;
    let claims_root = required_root_array(transaction, manifest, "document.record-claims")?;
    let comments_present = match options.get(transaction, "commentsPresent") {
        Some(Out::Any(Any::Bool(value))) => Some(value),
        Some(_) => {
            return Err(invalid_shared_comments(
                "The shared Document comment presence setting is invalid.",
            ))
        }
        None => None,
    };
    if comments_present != Some(true) && order_root.len(transaction) == 0 {
        if comments.len(transaction) > 0 {
            return Err(invalid_shared_comments(
                "The shared Document comment order and record set disagree.",
            ));
        }
    }

    let order = validated_order(
        transaction,
        &order_root,
        &comments,
        "comment",
        MAX_DOCUMENT_COMMENTS,
    )?;
    let claims = read_document_record_claims(transaction, &claims_root)?;
    let mut total_text_bytes = 0_usize;
    let mut total_replies = 0_usize;
    let mut records = HashMap::with_capacity(order.len());
    for id in &order {
        let record = required_nested_map(
            comments.get(transaction, id.as_str()),
            "Document comment record",
        )?;
        let projected = read_comment_record(
            transaction,
            record,
            id,
            &mut total_text_bytes,
            &mut total_replies,
        )?;
        if claims.claim_for("comment", None, id).is_none() {
            return Err(invalid_shared_comments(format!(
                "The shared Document comment ID '{id}' is missing its immutable claim."
            )));
        }
        for reply_id in &projected.reply_order {
            if claims
                .claim_for("comment-reply", Some(id), reply_id)
                .is_none()
            {
                return Err(invalid_shared_comments(format!(
                    "The shared Document reply ID '{reply_id}' is missing its immutable claim."
                )));
            }
        }
        records.insert(id.clone(), projected);
    }
    Ok(DocumentCommentsState {
        comments,
        order_root,
        claims_root,
        order,
        records,
        claims,
    })
}

pub(in crate::collaboration) fn project_document_comments<T: ReadTxn>(
    transaction: &T,
    manifest: &NativeOfficeCollaborationManifest,
    fragment: &yrs::XmlFragmentRef,
) -> UseResult<Vec<NativeOfficeCollaborationDocumentComment>> {
    let state = read_comment_state(transaction, manifest)?;
    let mut anchors = collect_comment_anchors(transaction, fragment)?;
    let mut comments = Vec::with_capacity(state.order.len());
    for id in state.order {
        let record = state.records.get(&id).expect("validated comment order");
        let record_anchors = anchors.remove(&id).unwrap_or_default();
        comments.push(NativeOfficeCollaborationDocumentComment {
            id: record.id.clone(),
            actor_id: record.actor_id.clone(),
            author: record.author.clone(),
            date: record.date.clone(),
            text: record.text.clone(),
            resolved: record.resolved,
            replies: record
                .reply_order
                .iter()
                .map(|reply_id| {
                    let reply = &record.reply_records[reply_id];
                    NativeOfficeCollaborationDocumentCommentReply {
                        id: reply.id.clone(),
                        actor_id: reply.actor_id.clone(),
                        author: reply.author.clone(),
                        date: reply.date.clone(),
                        text: reply.text.clone(),
                    }
                })
                .collect(),
            detached: record_anchors.is_empty(),
            anchors: record_anchors
                .into_iter()
                .map(|anchor| NativeOfficeCollaborationDocumentCommentAnchor {
                    paragraph_id: anchor.paragraph_id,
                    text_id: anchor.text_id,
                    start_utf16: anchor.start_utf16,
                    end_utf16: anchor.end_utf16,
                    text: anchor.text,
                })
                .collect(),
        });
    }
    Ok(comments)
}

pub(super) fn collect_comment_anchors<T: ReadTxn>(
    transaction: &T,
    fragment: &yrs::XmlFragmentRef,
) -> UseResult<HashMap<String, Vec<CommentAnchor>>> {
    let mut anchors = HashMap::<String, Vec<CommentAnchor>>::new();
    for node in fragment.successors(transaction) {
        let XmlOut::Element(paragraph) = node else {
            continue;
        };
        if !is_identity_paragraph_tag(paragraph.tag()) {
            continue;
        }
        let paragraph_id =
            document_identity_attribute(&paragraph, transaction, PARAGRAPH_ID_ATTRIBUTE)?;
        let text_id = document_identity_attribute(&paragraph, transaction, TEXT_ID_ATTRIBUTE)?;
        let mut paragraph_cursor = 0_u32;
        for child in paragraph.children(transaction) {
            let XmlOut::Text(text) = child else {
                paragraph_cursor = paragraph_cursor.checked_add(1).ok_or_else(|| {
                    invalid_shared_comments("A Document comment anchor offset is too large.")
                })?;
                continue;
            };
            for chunk in text.diff(transaction, |_| ()) {
                let (length, chunk_text) = match chunk.insert {
                    Out::Any(Any::String(value)) => (
                        u32::try_from(value.encode_utf16().count()).map_err(|_| {
                            invalid_shared_comments(
                                "A Document comment anchor offset is too large.",
                            )
                        })?,
                        value.to_string(),
                    ),
                    _ => (1, String::new()),
                };
                let end = paragraph_cursor.checked_add(length).ok_or_else(|| {
                    invalid_shared_comments("A Document comment anchor offset is too large.")
                })?;
                for (_, id) in comment_attributes(chunk.attributes.as_deref())? {
                    append_anchor(
                        anchors.entry(id).or_default(),
                        CommentAnchor {
                            paragraph_id: paragraph_id.clone(),
                            text_id: text_id.clone(),
                            start_utf16: paragraph_cursor,
                            end_utf16: end,
                            text: chunk_text.clone(),
                        },
                    );
                }
                paragraph_cursor = end;
            }
        }
    }
    Ok(anchors)
}

pub(super) fn collect_comment_mark_segments<T: ReadTxn>(
    transaction: &T,
    fragment: &yrs::XmlFragmentRef,
) -> UseResult<Vec<CommentMarkSegment>> {
    let mut segments = Vec::new();
    for node in fragment.successors(transaction) {
        let XmlOut::Text(text) = node else {
            continue;
        };
        let mut cursor = 0_u32;
        for chunk in text.diff(transaction, |_| ()) {
            let length = match chunk.insert {
                Out::Any(Any::String(value)) => u32::try_from(value.encode_utf16().count())
                    .map_err(|_| {
                        invalid_shared_comments("A Document comment mark offset is too large.")
                    })?,
                _ => 1,
            };
            let end = cursor.checked_add(length).ok_or_else(|| {
                invalid_shared_comments("A Document comment mark offset is too large.")
            })?;
            for (attribute, comment_id) in comment_attributes(chunk.attributes.as_deref())? {
                let value = chunk
                    .attributes
                    .as_ref()
                    .and_then(|attributes| attributes.get(attribute.as_str()))
                    .cloned()
                    .expect("validated comment attribute");
                segments.push(CommentMarkSegment {
                    text: text.clone(),
                    attribute,
                    value,
                    start_utf16: cursor,
                    end_utf16: end,
                    comment_id,
                });
            }
            cursor = end;
        }
    }
    Ok(segments)
}

pub(super) fn comment_attributes(attributes: Option<&Attrs>) -> UseResult<Vec<(String, String)>> {
    let Some(attributes) = attributes else {
        return Ok(Vec::new());
    };
    let mut result = Vec::new();
    let mut ids = HashSet::new();
    for (key, value) in attributes {
        if !is_comment_attribute(key) {
            continue;
        }
        let Any::Map(fields) = value else {
            return Err(invalid_shared_comments(
                "A shared Document comment mark is not an object.",
            ));
        };
        let Some(Any::String(id)) = fields.get("id") else {
            return Err(invalid_shared_comments(
                "A shared Document comment mark omits its string ID.",
            ));
        };
        if fields.len() != 1 {
            return Err(invalid_shared_comments(
                "A shared Document comment mark contains unsupported fields.",
            ));
        }
        validate_shared_identifier(id, "comment mark")?;
        ids.insert(id.to_string());
        result.push((key.to_string(), id.to_string()));
    }
    if ids.len() > 1 {
        return Err(invalid_shared_comments(
            "Overlapping Document comment marks carry different IDs.",
        ));
    }
    Ok(result)
}

fn read_comment_record<T: ReadTxn>(
    transaction: &T,
    record: MapRef,
    expected_id: &str,
    total_text_bytes: &mut usize,
    total_replies: &mut usize,
) -> UseResult<DocumentCommentRecord> {
    let id = required_string(record.get(transaction, "id"), "comment ID")?;
    validate_shared_identifier(&id, "comment")?;
    if id != expected_id {
        return Err(invalid_shared_comments(
            "A shared Document comment identity does not match its map key.",
        ));
    }
    let actor_id = optional_identifier(record.get(transaction, "actorId"), "comment actor")?;
    let expected_fields = if actor_id.is_some() { 8 } else { 7 };
    if record.len(transaction) != expected_fields
        || record.iter(transaction).any(|(key, _)| {
            !matches!(
                key.as_ref(),
                "id" | "actorId"
                    | "author"
                    | "date"
                    | "text"
                    | "resolved"
                    | "replies"
                    | "replyOrder"
            )
        })
    {
        return Err(invalid_shared_comments(
            "A shared Document comment record contains unsupported fields.",
        ));
    }
    let author = required_string(record.get(transaction, "author"), "comment author")?;
    let date = required_string(record.get(transaction, "date"), "comment date")?;
    let text = required_string(record.get(transaction, "text"), "comment text")?;
    account_text(total_text_bytes, [&author, &date, &text])?;
    let resolved = match record.get(transaction, "resolved") {
        Some(Out::Any(Any::Bool(value))) => value,
        _ => {
            return Err(invalid_shared_comments(
                "A shared Document comment resolution is invalid.",
            ))
        }
    };
    let replies = required_nested_map(record.get(transaction, "replies"), "comment replies")?;
    let reply_order_root =
        required_nested_array(record.get(transaction, "replyOrder"), "comment reply order")?;
    let reply_order = validated_order(
        transaction,
        &reply_order_root,
        &replies,
        "comment reply",
        MAX_DOCUMENT_COMMENT_REPLIES,
    )?;
    *total_replies = total_replies
        .checked_add(reply_order.len())
        .ok_or_else(|| {
            invalid_shared_comments("The shared Document contains too many comment replies.")
        })?;
    if *total_replies > MAX_DOCUMENT_COMMENT_REPLIES {
        return Err(invalid_shared_comments(
            "The shared Document contains too many comment replies.",
        ));
    }
    let mut reply_records = HashMap::with_capacity(reply_order.len());
    for reply_id in &reply_order {
        let reply = required_nested_map(
            replies.get(transaction, reply_id.as_str()),
            "Document comment reply record",
        )?;
        let id = required_string(reply.get(transaction, "id"), "comment reply ID")?;
        validate_shared_identifier(&id, "comment reply")?;
        if &id != reply_id {
            return Err(invalid_shared_comments(
                "A shared Document comment reply identity does not match its map key.",
            ));
        }
        let actor_id =
            optional_identifier(reply.get(transaction, "actorId"), "comment reply actor")?;
        let expected_fields = if actor_id.is_some() { 5 } else { 4 };
        if reply.len(transaction) != expected_fields
            || reply.iter(transaction).any(|(key, _)| {
                !matches!(key.as_ref(), "id" | "actorId" | "author" | "date" | "text")
            })
        {
            return Err(invalid_shared_comments(
                "A shared Document comment reply contains unsupported fields.",
            ));
        }
        let author = required_string(reply.get(transaction, "author"), "comment reply author")?;
        let date = required_string(reply.get(transaction, "date"), "comment reply date")?;
        let text = required_string(reply.get(transaction, "text"), "comment reply text")?;
        account_text(total_text_bytes, [&author, &date, &text])?;
        reply_records.insert(
            id.clone(),
            DocumentCommentReplyRecord {
                id,
                actor_id,
                author,
                date,
                text,
            },
        );
    }
    Ok(DocumentCommentRecord {
        record,
        id,
        actor_id,
        author,
        date,
        text,
        resolved,
        replies,
        reply_order_root,
        reply_order,
        reply_records,
    })
}

pub(in crate::collaboration) fn read_document_record_claims<T: ReadTxn>(
    transaction: &T,
    claims: &ArrayRef,
) -> UseResult<DocumentRecordClaims> {
    if claims.len(transaction) as usize > MAX_DOCUMENT_COMMENT_CLAIMS {
        return Err(invalid_shared_comments(
            "The shared Document contains too many immutable record claims.",
        ));
    }
    let mut by_identity = HashMap::new();
    for index in 0..claims.len(transaction) {
        let raw = match claims.get(transaction, index) {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            _ => {
                return Err(invalid_shared_comments(
                    "A Document record claim is not a string.",
                ))
            }
        };
        let value: JsonValue = serde_json::from_str(&raw)
            .map_err(|_| invalid_shared_comments("A Document record claim is not valid JSON."))?;
        if canonical_json(&value)? != raw {
            return Err(invalid_shared_comments(
                "A Document record claim is not canonical JSON.",
            ));
        }
        let object = value
            .as_object()
            .ok_or_else(|| invalid_shared_comments("A Document record claim is not an object."))?;
        let kind = object
            .get("kind")
            .and_then(JsonValue::as_str)
            .ok_or_else(|| invalid_shared_comments("A Document record claim kind is invalid."))?;
        if !matches!(
            kind,
            "bibliography-source" | "change-decision" | "comment" | "comment-reply"
        ) {
            return Err(invalid_shared_comments(
                "A Document record claim kind is unsupported.",
            ));
        }
        let parent_id = object
            .get("parentId")
            .map(|value| {
                value.as_str().map(str::to_owned).ok_or_else(|| {
                    invalid_shared_comments("A Document record claim parent is invalid.")
                })
            })
            .transpose()?;
        let expected_keys = if parent_id.is_some() { 4 } else { 3 };
        if object.len() != expected_keys
            || object
                .keys()
                .any(|key| !matches!(key.as_str(), "fingerprint" | "id" | "kind" | "parentId"))
            || (kind == "comment-reply") != parent_id.is_some()
        {
            return Err(invalid_shared_comments(
                "A Document record claim shape is invalid.",
            ));
        }
        let id = object
            .get("id")
            .and_then(JsonValue::as_str)
            .ok_or_else(|| invalid_shared_comments("A Document record claim ID is invalid."))?
            .to_owned();
        validate_shared_identifier(&id, "record claim")?;
        if let Some(parent_id) = &parent_id {
            validate_shared_identifier(parent_id, "record claim parent")?;
        }
        let fingerprint = object
            .get("fingerprint")
            .and_then(JsonValue::as_str)
            .ok_or_else(|| {
                invalid_shared_comments("A Document record claim fingerprint is invalid.")
            })?
            .to_owned();
        let fingerprint_value: JsonValue = serde_json::from_str(&fingerprint).map_err(|_| {
            invalid_shared_comments("A Document record claim fingerprint is not valid JSON.")
        })?;
        if !fingerprint_value.is_object() || canonical_json(&fingerprint_value)? != fingerprint {
            return Err(invalid_shared_comments(
                "A Document record claim fingerprint is not canonical JSON.",
            ));
        }
        let identity = (kind.to_owned(), parent_id, id.clone());
        if let Some(existing) = by_identity.insert(identity, fingerprint.clone()) {
            if existing != fingerprint {
                return Err(invalid_shared_comments(format!(
                    "The shared Document {kind} ID '{id}' was claimed by different records."
                )));
            }
        }
    }
    Ok(DocumentRecordClaims { by_identity })
}

fn validated_order<T: ReadTxn>(
    transaction: &T,
    order: &ArrayRef,
    records: &MapRef,
    label: &str,
    maximum: usize,
) -> UseResult<Vec<String>> {
    if records.len(transaction) as usize > maximum {
        return Err(invalid_shared_comments(format!(
            "The shared Document contains too many {label} records."
        )));
    }
    let mut seen = HashSet::new();
    let mut result = Vec::new();
    for index in 0..order.len(transaction) {
        let id = match order.get(transaction, index) {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            _ => {
                return Err(invalid_shared_comments(format!(
                    "The shared Document {label} order contains a non-string identity."
                )))
            }
        };
        validate_shared_identifier(&id, label)?;
        if seen.insert(id.clone()) {
            result.push(id);
        }
    }
    if seen.len() != records.len(transaction) as usize
        || records.iter(transaction).any(|(id, _)| !seen.contains(id))
    {
        return Err(invalid_shared_comments(format!(
            "The shared Document {label} order and record set disagree."
        )));
    }
    Ok(result)
}

fn required_root_map<T: ReadTxn>(
    transaction: &T,
    manifest: &NativeOfficeCollaborationManifest,
    suffix: &str,
) -> UseResult<MapRef> {
    let root = format!("{}.{}", manifest.namespace, suffix);
    match transaction.get(&root) {
        Some(Out::YMap(value)) => Ok(value),
        _ => Err(invalid_shared_comments(format!(
            "The shared Document root '{root}' is not a Y.Map."
        ))),
    }
}

fn required_root_array<T: ReadTxn>(
    transaction: &T,
    manifest: &NativeOfficeCollaborationManifest,
    suffix: &str,
) -> UseResult<ArrayRef> {
    let root = format!("{}.{}", manifest.namespace, suffix);
    match transaction.get(&root) {
        Some(Out::YArray(value)) => Ok(value),
        _ => Err(invalid_shared_comments(format!(
            "The shared Document root '{root}' is not a Y.Array."
        ))),
    }
}

fn required_nested_map(value: Option<Out>, label: &str) -> UseResult<MapRef> {
    match value {
        Some(Out::YMap(value)) => Ok(value),
        _ => Err(invalid_shared_comments(format!(
            "The shared {label} is not a Y.Map."
        ))),
    }
}

fn required_nested_array(value: Option<Out>, label: &str) -> UseResult<ArrayRef> {
    match value {
        Some(Out::YArray(value)) => Ok(value),
        _ => Err(invalid_shared_comments(format!(
            "The shared {label} is not a Y.Array."
        ))),
    }
}

fn required_string(value: Option<Out>, label: &str) -> UseResult<String> {
    match value {
        Some(Out::Any(Any::String(value))) => Ok(value.to_string()),
        _ => Err(invalid_shared_comments(format!(
            "The shared Document {label} is not a string."
        ))),
    }
}

fn optional_identifier(value: Option<Out>, label: &str) -> UseResult<Option<String>> {
    match value {
        Some(Out::Any(Any::String(value))) => {
            validate_shared_identifier(&value, label)?;
            Ok(Some(value.to_string()))
        }
        Some(_) => Err(invalid_shared_comments(format!(
            "The shared Document {label} is invalid."
        ))),
        None => Ok(None),
    }
}

fn account_text<const N: usize>(total: &mut usize, values: [&String; N]) -> UseResult<()> {
    for value in values {
        *total = total.checked_add(value.len()).ok_or_else(|| {
            invalid_shared_comments("The shared Document comment text is too large.")
        })?;
    }
    if *total > MAX_DOCUMENT_COMMENT_TEXT_BYTES {
        return Err(invalid_shared_comments(
            "The shared Document comment text is too large.",
        ));
    }
    Ok(())
}

fn append_anchor(anchors: &mut Vec<CommentAnchor>, next: CommentAnchor) {
    if let Some(previous) = anchors.last_mut() {
        if previous.paragraph_id == next.paragraph_id
            && previous.text_id == next.text_id
            && previous.end_utf16 == next.start_utf16
        {
            previous.end_utf16 = next.end_utf16;
            previous.text.push_str(&next.text);
            return;
        }
    }
    anchors.push(next);
}

fn is_comment_attribute(value: &str) -> bool {
    if value == "documentComment" {
        return true;
    }
    value
        .strip_prefix("documentComment--")
        .is_some_and(|suffix| {
            suffix.len() == 8
                && suffix.bytes().all(|value| {
                    value.is_ascii_alphanumeric() || matches!(value, b'+' | b'/' | b'=')
                })
        })
}
