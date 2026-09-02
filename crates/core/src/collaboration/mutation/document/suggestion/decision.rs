use std::collections::{HashMap, HashSet};

use a3s_use_core::UseResult;
use serde_json::json;
use yrs::{Any, Array, ArrayRef, Map, MapPrelim, Out, ReadTxn};

use super::super::comment::{
    canonical_json, read_document_record_claims, valid_canonical_utc_timestamp,
    DocumentRecordClaims,
};

use crate::collaboration::{
    collaboration_error, NativeOfficeCollaborationDocumentChangeDecision,
    NativeOfficeCollaborationDocumentChangeKind,
    NativeOfficeCollaborationDocumentSuggestionDecision, NativeOfficeCollaborationManifest,
};

const DECISION_FIELDS: &[&str] = &[
    "id",
    "changeId",
    "changeKind",
    "suggestedByActorId",
    "suggestedBy",
    "suggestedAt",
    "text",
    "decision",
    "decidedByActorId",
    "decidedBy",
    "decidedAt",
];
const MAX_DOCUMENT_CHANGE_DECISIONS: usize = 1_048_576;
const MAX_DOCUMENT_CHANGE_DECISION_TEXT_UTF16: usize = 1_048_576;
const MAX_DOCUMENT_CHANGE_DECISION_TEXT_BYTES: usize = 64 * 1024 * 1024;

pub(super) struct DocumentChangeDecisionState {
    pub records_root: yrs::MapRef,
    pub order_root: yrs::ArrayRef,
    pub claims_root: ArrayRef,
    pub records: HashMap<String, NativeOfficeCollaborationDocumentChangeDecision>,
    pub ordered_ids: Vec<String>,
    pub claims: DocumentRecordClaims,
}

pub(super) fn read_document_change_decisions<T: ReadTxn>(
    transaction: &T,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<DocumentChangeDecisionState> {
    let records_name = format!("{}.document.change-decisions", manifest.namespace);
    let order_name = format!("{}.document.change-decision-order", manifest.namespace);
    let claims_name = format!("{}.document.record-claims", manifest.namespace);
    let records_root = match transaction.get(&records_name) {
        Some(Out::YMap(value)) => value,
        _ => return Err(invalid_decisions("record root")),
    };
    let order_root = match transaction.get(&order_name) {
        Some(Out::YArray(value)) => value,
        _ => return Err(invalid_decisions("order root")),
    };
    let claims_root = match transaction.get(&claims_name) {
        Some(Out::YArray(value)) => value,
        _ => return Err(invalid_decisions("claim root")),
    };
    if records_root.len(transaction) as usize > MAX_DOCUMENT_CHANGE_DECISIONS
        || order_root.len(transaction) as usize > MAX_DOCUMENT_CHANGE_DECISIONS
    {
        return Err(invalid_decisions("record count"));
    }
    let claims = read_document_record_claims(transaction, &claims_root)?;

    let mut ordered_ids = Vec::new();
    let mut seen = HashSet::new();
    for value in order_root.iter(transaction) {
        let Out::Any(Any::String(value)) = value else {
            return Err(invalid_decisions("order"));
        };
        let id = value.to_string();
        validate_shared_identifier(&id, "decision ID")?;
        if seen.insert(id.clone()) {
            ordered_ids.push(id);
        }
    }
    let record_keys = records_root
        .keys(transaction)
        .map(ToOwned::to_owned)
        .collect::<HashSet<_>>();
    if seen != record_keys {
        return Err(invalid_decisions("order and record set"));
    }

    let mut records = HashMap::new();
    let mut change_ids = HashSet::new();
    let mut total_text_bytes = 0_usize;
    for id in &ordered_ids {
        let record = match records_root.get(transaction, id) {
            Some(Out::YMap(value)) => value,
            _ => return Err(invalid_decisions("record")),
        };
        let parsed = read_decision_record(&record, transaction, id)?;
        total_text_bytes = total_text_bytes
            .checked_add(parsed.text.len())
            .ok_or_else(|| invalid_decisions("text size"))?;
        if total_text_bytes > MAX_DOCUMENT_CHANGE_DECISION_TEXT_BYTES {
            return Err(invalid_decisions("text size"));
        }
        let fingerprint = decision_fingerprint(&parsed)?;
        match claims.claim_for("change-decision", None, id) {
            Some(current) if current == fingerprint => {}
            Some(_) => return Err(invalid_decisions("immutable claim fingerprint")),
            None => return Err(invalid_decisions("immutable claim")),
        }
        let change_identity = format!("{}:{}", parsed.change_kind.as_str(), parsed.change_id);
        if !change_ids.insert(change_identity) {
            return Err(invalid_decisions("duplicate tracked-change decision"));
        }
        records.insert(id.clone(), parsed);
    }
    Ok(DocumentChangeDecisionState {
        records_root,
        order_root,
        claims_root,
        records,
        ordered_ids,
        claims,
    })
}

pub(in crate::collaboration) fn project_document_change_decisions<T: ReadTxn>(
    transaction: &T,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<Vec<NativeOfficeCollaborationDocumentChangeDecision>> {
    let state = read_document_change_decisions(transaction, manifest)?;
    state
        .ordered_ids
        .iter()
        .map(|id| {
            state
                .records
                .get(id)
                .cloned()
                .ok_or_else(|| invalid_decisions("order and record set"))
        })
        .collect()
}

pub(super) fn insert_decision_record(
    transaction: &mut yrs::TransactionMut<'_>,
    records: &yrs::MapRef,
    order: &yrs::ArrayRef,
    decision: &NativeOfficeCollaborationDocumentChangeDecision,
) {
    let record = records.insert(transaction, decision.id.as_str(), MapPrelim::default());
    record.insert(transaction, "id", decision.id.as_str());
    record.insert(transaction, "changeId", decision.change_id.as_str());
    record.insert(transaction, "changeKind", decision.change_kind.as_str());
    if let Some(actor_id) = &decision.suggested_by_actor_id {
        record.insert(transaction, "suggestedByActorId", actor_id.as_str());
    }
    record.insert(transaction, "suggestedBy", decision.suggested_by.as_str());
    record.insert(transaction, "suggestedAt", decision.suggested_at.as_str());
    record.insert(transaction, "text", decision.text.as_str());
    record.insert(transaction, "decision", decision.decision.as_str());
    if let Some(actor_id) = &decision.decided_by_actor_id {
        record.insert(transaction, "decidedByActorId", actor_id.as_str());
    }
    record.insert(transaction, "decidedBy", decision.decided_by.as_str());
    record.insert(transaction, "decidedAt", decision.decided_at.as_str());
    order.push_back(transaction, decision.id.as_str());
}

pub(super) fn decision_record_id(
    kind: NativeOfficeCollaborationDocumentChangeKind,
    change_id: &str,
) -> String {
    format!("{}:{change_id}", kind.as_str())
}

pub(super) fn decision_fingerprint(
    decision: &NativeOfficeCollaborationDocumentChangeDecision,
) -> UseResult<String> {
    let value = serde_json::to_value(decision).map_err(|error| {
        collaboration_error(
            "office.collaboration.content_invalid",
            format!("Failed to encode a tracked-change decision fingerprint: {error}"),
        )
    })?;
    canonical_json(&value)
}

pub(super) fn encoded_decision_claim(
    decision: &NativeOfficeCollaborationDocumentChangeDecision,
) -> UseResult<String> {
    canonical_json(&json!({
        "fingerprint": decision_fingerprint(decision)?,
        "id": decision.id,
        "kind": "change-decision",
    }))
}

fn read_decision_record<T: ReadTxn>(
    record: &yrs::MapRef,
    transaction: &T,
    expected_id: &str,
) -> UseResult<NativeOfficeCollaborationDocumentChangeDecision> {
    if record
        .keys(transaction)
        .any(|key| !DECISION_FIELDS.contains(&key))
    {
        return Err(invalid_decisions("record fields"));
    }
    let id = required_string(record.get(transaction, "id"), "decision ID")?;
    validate_shared_identifier(&id, "decision ID")?;
    if id != expected_id {
        return Err(invalid_decisions("record identity"));
    }
    let change_id = required_string(record.get(transaction, "changeId"), "change ID")?;
    validate_shared_identifier(&change_id, "change ID")?;
    let change_kind = parse_kind(record.get(transaction, "changeKind"))?;
    if id != decision_record_id(change_kind, &change_id) {
        return Err(invalid_decisions("canonical identity"));
    }
    let suggested_by_actor_id = optional_identifier(
        record.get(transaction, "suggestedByActorId"),
        "suggestion actor ID",
    )?;
    let suggested_by =
        required_string(record.get(transaction, "suggestedBy"), "suggestion author")?;
    let suggested_at = required_string(record.get(transaction, "suggestedAt"), "suggestion date")?;
    let text = required_string(record.get(transaction, "text"), "suggestion text")?;
    validate_shared_identifier(&suggested_by, "suggestion author")?;
    validate_shared_timestamp(&suggested_at, "suggestion date")?;
    validate_shared_text(&text)?;
    let decision = parse_decision(record.get(transaction, "decision"))?;
    let decided_by_actor_id = optional_identifier(
        record.get(transaction, "decidedByActorId"),
        "decision actor ID",
    )?;
    let decided_by = required_string(record.get(transaction, "decidedBy"), "decision author")?;
    let decided_at = required_string(record.get(transaction, "decidedAt"), "decision date")?;
    validate_shared_identifier(&decided_by, "decision author")?;
    validate_shared_timestamp(&decided_at, "decision date")?;
    Ok(NativeOfficeCollaborationDocumentChangeDecision {
        id,
        change_id,
        change_kind,
        suggested_by_actor_id,
        suggested_by,
        suggested_at,
        text,
        decision,
        decided_by_actor_id,
        decided_by,
        decided_at,
    })
}

fn parse_kind(value: Option<Out>) -> UseResult<NativeOfficeCollaborationDocumentChangeKind> {
    match required_string(value, "change kind")?.as_str() {
        "insertion" => Ok(NativeOfficeCollaborationDocumentChangeKind::Insertion),
        "deletion" => Ok(NativeOfficeCollaborationDocumentChangeKind::Deletion),
        "formatting" => Ok(NativeOfficeCollaborationDocumentChangeKind::Formatting),
        "paragraph-formatting" => {
            Ok(NativeOfficeCollaborationDocumentChangeKind::ParagraphFormatting)
        }
        "numbering" => Ok(NativeOfficeCollaborationDocumentChangeKind::Numbering),
        _ => Err(invalid_decisions("change kind")),
    }
}

fn parse_decision(
    value: Option<Out>,
) -> UseResult<NativeOfficeCollaborationDocumentSuggestionDecision> {
    match required_string(value, "decision")?.as_str() {
        "accept" => Ok(NativeOfficeCollaborationDocumentSuggestionDecision::Accept),
        "reject" => Ok(NativeOfficeCollaborationDocumentSuggestionDecision::Reject),
        _ => Err(invalid_decisions("decision action")),
    }
}

fn optional_identifier(value: Option<Out>, label: &str) -> UseResult<Option<String>> {
    match value {
        None => Ok(None),
        Some(value) => {
            let value = required_string(Some(value), label)?;
            validate_shared_identifier(&value, label)?;
            Ok(Some(value))
        }
    }
}

fn required_string(value: Option<Out>, label: &str) -> UseResult<String> {
    match value {
        Some(Out::Any(Any::String(value))) => Ok(value.to_string()),
        _ => Err(invalid_decisions(label)),
    }
}

fn validate_shared_identifier(value: &str, label: &str) -> UseResult<()> {
    if !value.is_empty() && value == value.trim() && value.encode_utf16().count() <= 256 {
        return Ok(());
    }
    Err(invalid_decisions(label))
}

fn validate_shared_timestamp(value: &str, label: &str) -> UseResult<()> {
    if valid_canonical_utc_timestamp(value) && value.encode_utf16().count() <= 256 {
        return Ok(());
    }
    Err(invalid_decisions(label))
}

fn validate_shared_text(value: &str) -> UseResult<()> {
    if !value.is_empty() && value.encode_utf16().count() <= MAX_DOCUMENT_CHANGE_DECISION_TEXT_UTF16
    {
        return Ok(());
    }
    Err(invalid_decisions("suggestion text"))
}

fn invalid_decisions(label: &str) -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.content_invalid",
        format!("The shared Document tracked-change decision {label} is invalid."),
    )
}
