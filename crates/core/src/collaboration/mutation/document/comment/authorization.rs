use std::collections::{BTreeSet, HashMap, HashSet};

use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;
use yrs::{Any, Array, Out, ReadTxn, Transact};

use super::state::{collect_comment_anchors, read_comment_state, CommentAnchor};
use super::{
    comment_fingerprint, encoded_claim, reply_fingerprint, validate_non_empty_text,
    validate_timestamp,
};
use crate::collaboration::document::{
    canonical_content_without_comment_marks_sha256, canonical_map_without_key_sha256,
    canonical_visible_root_sha256,
};
use crate::collaboration::{
    collaboration_error, NativeOfficeCollaborationArtifactKind, NativeOfficeCollaborationManifest,
};

pub(in crate::collaboration) fn validate_authorized_comment_update(
    before: &yrs::Doc,
    candidate: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    actor_name: &str,
) -> UseResult<()> {
    if manifest.kind != NativeOfficeCollaborationArtifactKind::Document {
        return Err(permission_denied(
            "Comment mode can publish review updates only for Document artifacts.",
        ));
    }
    validate_non_empty_text(actor_name, "actorName", "actor display name")?;

    let before_transaction = before.transact();
    let candidate_transaction = candidate.transact();
    if candidate_transaction.has_missing_updates() {
        return Err(permission_denied(
            "A comment update with unresolved Yjs dependencies cannot be authorized safely.",
        ));
    }
    assert_root_boundaries(&before_transaction, &candidate_transaction, manifest)?;

    let content_root = format!("{}.document.content", manifest.namespace);
    let before_content = required_root(&before_transaction, &content_root)?;
    let candidate_content = required_root(&candidate_transaction, &content_root)?;
    if canonical_content_without_comment_marks_sha256(&before_transaction, &before_content)
        != canonical_content_without_comment_marks_sha256(
            &candidate_transaction,
            &candidate_content,
        )
    {
        return Err(permission_denied(
            "Comment mode cannot change Document content, structure, or non-comment formatting.",
        ));
    }

    let options_root = format!("{}.document.options", manifest.namespace);
    let before_options = required_map(&before_transaction, &options_root)?;
    let candidate_options = required_map(&candidate_transaction, &options_root)?;
    if canonical_map_without_key_sha256(&before_transaction, &before_options, "commentsPresent")
        != canonical_map_without_key_sha256(
            &candidate_transaction,
            &candidate_options,
            "commentsPresent",
        )
    {
        return Err(permission_denied(
            "Comment mode cannot change Document options other than comment presence.",
        ));
    }

    let before_state = read_comment_state(&before_transaction, manifest)?;
    let candidate_state = read_comment_state(&candidate_transaction, manifest)?;
    let before_claims = raw_claims(&before_transaction, &before_state.claims_root)?;
    let candidate_claims = raw_claims(&candidate_transaction, &candidate_state.claims_root)?;
    if !candidate_claims.starts_with(&before_claims) {
        return Err(permission_denied(
            "Comment mode must preserve immutable Document record claims and append new claims.",
        ));
    }

    assert_retained_order(&before_state.order, &candidate_state.order, "comments")?;
    let mut allowed_new_claims = HashSet::new();
    for id in &before_state.order {
        let before_record = &before_state.records[id];
        let Some(candidate_record) = candidate_state.records.get(id) else {
            assert_owner(
                before_record.actor_id.as_deref(),
                manifest,
                "delete another actor's comment",
            )?;
            continue;
        };
        if before_record.id != candidate_record.id
            || before_record.actor_id != candidate_record.actor_id
            || before_record.author != candidate_record.author
            || before_record.date != candidate_record.date
            || before_record.text != candidate_record.text
        {
            return Err(permission_denied(
                "Comment mode cannot rewrite an existing Document comment.",
            ));
        }
        assert_replies(
            before_record,
            candidate_record,
            manifest,
            actor_name,
            &candidate_state,
            &mut allowed_new_claims,
        )?;
    }

    let before_ids = before_state.order.iter().collect::<HashSet<_>>();
    for id in &candidate_state.order {
        if before_ids.contains(id) {
            continue;
        }
        let record = &candidate_state.records[id];
        assert_new_comment(record, manifest, actor_name)?;
        let fingerprint = comment_fingerprint(
            &record.id,
            &manifest.actor_id,
            &record.author,
            &record.date,
            &record.text,
        )?;
        assert_claim(&candidate_state, "comment", None, &record.id, &fingerprint)?;
        allowed_new_claims.insert(encoded_claim("comment", &record.id, None, &fingerprint)?);
    }

    for claim in candidate_claims.iter().skip(before_claims.len()) {
        if !allowed_new_claims.contains(claim) {
            return Err(permission_denied(
                "Comment mode appended a record claim that does not belong to a new authenticated comment or reply.",
            ));
        }
    }

    let before_fragment = match before_content {
        Out::YXmlFragment(fragment) => fragment,
        _ => return Err(permission_denied("The Document content root is invalid.")),
    };
    let candidate_fragment = match candidate_content {
        Out::YXmlFragment(fragment) => fragment,
        _ => return Err(permission_denied("The Document content root is invalid.")),
    };
    let before_anchors = collect_comment_anchors(&before_transaction, &before_fragment)?;
    let candidate_anchors = collect_comment_anchors(&candidate_transaction, &candidate_fragment)?;
    assert_anchor_changes(
        &before_anchors,
        &candidate_anchors,
        &before_state,
        &candidate_state,
        manifest,
    )
}

fn assert_root_boundaries<T: ReadTxn, U: ReadTxn>(
    before: &T,
    candidate: &U,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<()> {
    let before_names = before
        .root_refs()
        .map(|(name, _)| name.to_owned())
        .collect::<BTreeSet<_>>();
    let candidate_names = candidate
        .root_refs()
        .map(|(name, _)| name.to_owned())
        .collect::<BTreeSet<_>>();
    if before_names != candidate_names {
        return Err(permission_denied(
            "Comment mode cannot add, remove, or replace Office collaboration roots.",
        ));
    }
    let allowed = [
        format!("{}.document.content", manifest.namespace),
        format!("{}.document.options", manifest.namespace),
        format!("{}.document.comments", manifest.namespace),
        format!("{}.document.comment-order", manifest.namespace),
        format!("{}.document.record-claims", manifest.namespace),
    ]
    .into_iter()
    .collect::<HashSet<_>>();
    for name in before_names {
        if allowed.contains(&name) {
            continue;
        }
        let before_value = required_root(before, &name)?;
        let candidate_value = required_root(candidate, &name)?;
        if canonical_visible_root_sha256(before, &before_value)
            != canonical_visible_root_sha256(candidate, &candidate_value)
        {
            return Err(permission_denied(format!(
                "Comment mode cannot change collaboration root '{name}'."
            )));
        }
    }
    Ok(())
}

fn assert_replies(
    before: &super::state::DocumentCommentRecord,
    candidate: &super::state::DocumentCommentRecord,
    manifest: &NativeOfficeCollaborationManifest,
    actor_name: &str,
    candidate_state: &super::state::DocumentCommentsState,
    allowed_new_claims: &mut HashSet<String>,
) -> UseResult<()> {
    assert_retained_order(
        &before.reply_order,
        &candidate.reply_order,
        &format!("replies in comment '{}'", before.id),
    )?;
    for reply_id in &before.reply_order {
        let before_reply = &before.reply_records[reply_id];
        let Some(candidate_reply) = candidate.reply_records.get(reply_id) else {
            assert_owner(
                before_reply.actor_id.as_deref(),
                manifest,
                "delete another actor's comment reply",
            )?;
            continue;
        };
        if before_reply.id != candidate_reply.id
            || before_reply.actor_id != candidate_reply.actor_id
            || before_reply.author != candidate_reply.author
            || before_reply.date != candidate_reply.date
            || before_reply.text != candidate_reply.text
        {
            return Err(permission_denied(
                "Comment mode cannot rewrite an existing Document comment reply.",
            ));
        }
    }
    let before_ids = before.reply_order.iter().collect::<HashSet<_>>();
    for reply_id in &candidate.reply_order {
        if before_ids.contains(reply_id) {
            continue;
        }
        let reply = &candidate.reply_records[reply_id];
        assert_actor_record(
            reply.actor_id.as_deref(),
            &reply.author,
            manifest,
            actor_name,
            "comment reply",
        )?;
        validate_timestamp(&reply.date, "date", "comment reply")?;
        validate_non_empty_text(&reply.text, "text", "comment reply text")?;
        let fingerprint = reply_fingerprint(
            &reply.id,
            &manifest.actor_id,
            &reply.author,
            &reply.date,
            &reply.text,
        )?;
        assert_claim(
            candidate_state,
            "comment-reply",
            Some(&candidate.id),
            &reply.id,
            &fingerprint,
        )?;
        allowed_new_claims.insert(encoded_claim(
            "comment-reply",
            &reply.id,
            Some(&candidate.id),
            &fingerprint,
        )?);
    }
    Ok(())
}

fn assert_new_comment(
    record: &super::state::DocumentCommentRecord,
    manifest: &NativeOfficeCollaborationManifest,
    actor_name: &str,
) -> UseResult<()> {
    assert_actor_record(
        record.actor_id.as_deref(),
        &record.author,
        manifest,
        actor_name,
        "comment",
    )?;
    if record.resolved || !record.reply_order.is_empty() {
        return Err(permission_denied(
            "A new Document comment must start unresolved and without embedded replies.",
        ));
    }
    validate_timestamp(&record.date, "date", "comment")?;
    validate_non_empty_text(&record.text, "text", "comment text")
}

fn assert_actor_record(
    actor_id: Option<&str>,
    author: &str,
    manifest: &NativeOfficeCollaborationManifest,
    actor_name: &str,
    label: &str,
) -> UseResult<()> {
    if actor_id == Some(manifest.actor_id.as_str()) && author == actor_name {
        return Ok(());
    }
    Err(permission_denied(format!(
        "A Document {label} created in comment mode must use authenticated actor '{}' and display name '{actor_name}'.",
        manifest.actor_id
    )))
}

fn assert_owner(
    actor_id: Option<&str>,
    manifest: &NativeOfficeCollaborationManifest,
    action: &str,
) -> UseResult<()> {
    if actor_id == Some(manifest.actor_id.as_str()) {
        return Ok(());
    }
    Err(permission_denied(format!("Comment mode cannot {action}.")))
}

fn assert_claim(
    state: &super::state::DocumentCommentsState,
    kind: &str,
    parent_id: Option<&str>,
    id: &str,
    expected: &str,
) -> UseResult<()> {
    if state.claims.claim_for(kind, parent_id, id) == Some(expected) {
        return Ok(());
    }
    Err(permission_denied(format!(
        "The immutable claim for new Document {kind} ID '{id}' does not match its authenticated record."
    )))
}

fn assert_retained_order(before: &[String], candidate: &[String], label: &str) -> UseResult<()> {
    let before_ids = before.iter().collect::<HashSet<_>>();
    let candidate_ids = candidate.iter().collect::<HashSet<_>>();
    let retained_before = before
        .iter()
        .filter(|id| candidate_ids.contains(id))
        .collect::<Vec<_>>();
    let retained_candidate = candidate
        .iter()
        .filter(|id| before_ids.contains(id))
        .collect::<Vec<_>>();
    let mut saw_new = false;
    for id in candidate {
        if before_ids.contains(id) {
            if saw_new {
                return Err(permission_denied(format!(
                    "Comment mode must preserve existing {label} order and append new records."
                )));
            }
        } else {
            saw_new = true;
        }
    }
    if retained_before == retained_candidate {
        Ok(())
    } else {
        Err(permission_denied(format!(
            "Comment mode must preserve existing {label} order and append new records."
        )))
    }
}

fn assert_anchor_changes(
    before: &HashMap<String, Vec<CommentAnchor>>,
    candidate: &HashMap<String, Vec<CommentAnchor>>,
    before_state: &super::state::DocumentCommentsState,
    candidate_state: &super::state::DocumentCommentsState,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<()> {
    let ids = before
        .keys()
        .chain(candidate.keys())
        .cloned()
        .collect::<BTreeSet<_>>();
    for id in ids {
        if before.get(&id) == candidate.get(&id) {
            continue;
        }
        let owner = candidate_state
            .records
            .get(&id)
            .and_then(|record| record.actor_id.as_deref())
            .or_else(|| {
                before_state
                    .records
                    .get(&id)
                    .and_then(|record| record.actor_id.as_deref())
            })
            .map(str::to_owned)
            .or_else(|| claim_actor_id(candidate_state, &id))
            .or_else(|| claim_actor_id(before_state, &id));
        if owner.as_deref() != Some(manifest.actor_id.as_str()) {
            return Err(permission_denied(format!(
                "Comment mode cannot change another actor's Document comment anchor '{id}'."
            )));
        }
    }
    for id in candidate.keys() {
        if candidate_state
            .claims
            .claim_for("comment", None, id)
            .is_none()
        {
            return Err(permission_denied(format!(
                "Document comment anchor '{id}' has no immutable record claim."
            )));
        }
    }
    Ok(())
}

fn claim_actor_id(state: &super::state::DocumentCommentsState, id: &str) -> Option<String> {
    let fingerprint = state.claims.claim_for("comment", None, id)?;
    let value: JsonValue = serde_json::from_str(fingerprint).ok()?;
    value
        .get("actorId")
        .and_then(JsonValue::as_str)
        .map(str::to_owned)
}

fn raw_claims<T: ReadTxn>(transaction: &T, claims: &yrs::ArrayRef) -> UseResult<Vec<String>> {
    let mut result = Vec::with_capacity(claims.len(transaction) as usize);
    for index in 0..claims.len(transaction) {
        match claims.get(transaction, index) {
            Some(Out::Any(Any::String(value))) => result.push(value.to_string()),
            _ => {
                return Err(permission_denied(
                    "A Document record claim is not a string.",
                ))
            }
        }
    }
    Ok(result)
}

fn required_root<T: ReadTxn>(transaction: &T, name: &str) -> UseResult<Out> {
    transaction
        .get(name)
        .ok_or_else(|| permission_denied(format!("Collaboration root '{name}' is missing.")))
}

fn required_map<T: ReadTxn>(transaction: &T, name: &str) -> UseResult<yrs::MapRef> {
    match transaction.get(name) {
        Some(Out::YMap(value)) => Ok(value),
        _ => Err(permission_denied(format!(
            "Collaboration root '{name}' is not a Y.Map."
        ))),
    }
}

fn permission_denied(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.permission_denied", message)
}
