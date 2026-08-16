use std::collections::{BTreeSet, HashMap};

use a3s_use_core::UseResult;
use yrs::types::Attrs;
use yrs::{Any, Out, ReadTxn, Text, Transact, XmlFragment, XmlOut};

use super::comment::valid_canonical_utc_timestamp;
use crate::collaboration::document::{
    canonical_content_without_suggestion_effects_sha256, canonical_visible_root_sha256,
    is_document_change_attribute,
};
use crate::collaboration::{
    collaboration_error, NativeOfficeCollaborationArtifactKind, NativeOfficeCollaborationManifest,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DocumentSuggestionKind {
    Insertion,
    Deletion,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DocumentSuggestionIdentity {
    id: String,
    kind: DocumentSuggestionKind,
    actor_id: Option<String>,
    author: String,
    date: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DocumentSuggestionPlacement {
    text_node: u32,
    baseline_start_utf16: u32,
    baseline_end_utf16: u32,
    text: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DocumentSuggestion {
    identity: DocumentSuggestionIdentity,
    text: String,
    placements: Vec<DocumentSuggestionPlacement>,
}

pub(in crate::collaboration) fn validate_authorized_suggestion_update(
    before: &yrs::Doc,
    candidate: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    actor_name: &str,
) -> UseResult<()> {
    if manifest.kind != NativeOfficeCollaborationArtifactKind::Document {
        return Err(permission_denied(
            "Suggest mode can publish tracked text proposals only for Document artifacts.",
        ));
    }
    validate_identifier(actor_name, "actor display name")?;

    let before_transaction = before.transact();
    let candidate_transaction = candidate.transact();
    if candidate_transaction.has_missing_updates() {
        return Err(permission_denied(
            "A suggestion update with unresolved Yjs dependencies cannot be authorized safely.",
        ));
    }
    assert_root_boundaries(&before_transaction, &candidate_transaction, manifest)?;

    let content_root = format!("{}.document.content", manifest.namespace);
    let before_content = required_content_root(&before_transaction, &content_root)?;
    let candidate_content = required_content_root(&candidate_transaction, &content_root)?;
    let before_suggestions = collect_suggestions(&before_transaction, &before_content)?;
    let candidate_suggestions = collect_suggestions(&candidate_transaction, &candidate_content)?;
    if canonical_content_without_suggestion_effects_sha256(
        &before_transaction,
        &Out::YXmlFragment(before_content),
    ) != canonical_content_without_suggestion_effects_sha256(
        &candidate_transaction,
        &Out::YXmlFragment(candidate_content),
    ) {
        return Err(permission_denied(
            "Suggest mode cannot change canonical Document text, structure, or non-suggestion formatting.",
        ));
    }

    assert_suggestion_changes(
        &before_suggestions,
        &candidate_suggestions,
        manifest,
        actor_name,
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
            "Suggest mode cannot add, remove, or replace Office collaboration roots.",
        ));
    }
    let content_root = format!("{}.document.content", manifest.namespace);
    for name in before_names {
        if name == content_root {
            continue;
        }
        let before_value = required_root(before, &name)?;
        let candidate_value = required_root(candidate, &name)?;
        if canonical_visible_root_sha256(before, &before_value)
            != canonical_visible_root_sha256(candidate, &candidate_value)
        {
            return Err(permission_denied(format!(
                "Suggest mode cannot change collaboration root '{name}'."
            )));
        }
    }
    Ok(())
}

fn collect_suggestions<T: ReadTxn>(
    transaction: &T,
    fragment: &yrs::XmlFragmentRef,
) -> UseResult<HashMap<String, DocumentSuggestion>> {
    let mut suggestions = HashMap::<String, DocumentSuggestion>::new();
    let mut text_node = 0_u32;
    for node in fragment.successors(transaction) {
        let XmlOut::Text(text) = node else {
            continue;
        };
        let current_text_node = text_node;
        text_node = text_node.checked_add(1).ok_or_else(|| {
            permission_denied("The Document suggestion tree contains too many text nodes.")
        })?;
        let mut baseline_offset = 0_u32;
        for chunk in text.diff(transaction, |_| ()) {
            let identity = suggestion_identity(chunk.attributes.as_deref())?;
            let (chunk_text, chunk_length) = match chunk.insert {
                Out::Any(Any::String(value)) => {
                    let length = u32::try_from(value.encode_utf16().count()).map_err(|_| {
                        permission_denied("A Document suggestion text segment is too large.")
                    })?;
                    (Some(value.to_string()), length)
                }
                _ => (None, 1),
            };
            let is_insertion = identity
                .as_ref()
                .is_some_and(|value| value.kind == DocumentSuggestionKind::Insertion);
            let baseline_end = if is_insertion {
                baseline_offset
            } else {
                baseline_offset.checked_add(chunk_length).ok_or_else(|| {
                    permission_denied("A Document suggestion offset is too large.")
                })?
            };
            if let Some(identity) = identity {
                let Some(chunk_text) = chunk_text else {
                    return Err(permission_denied(
                        "Document suggestions may mark text only.",
                    ));
                };
                if chunk_text.is_empty() {
                    return Err(permission_denied(
                        "A Document suggestion cannot contain an empty marked segment.",
                    ));
                }
                append_suggestion(
                    &mut suggestions,
                    identity,
                    DocumentSuggestionPlacement {
                        text_node: current_text_node,
                        baseline_start_utf16: baseline_offset,
                        baseline_end_utf16: baseline_end,
                        text: chunk_text,
                    },
                )?;
            }
            baseline_offset = baseline_end;
        }
    }
    Ok(suggestions)
}

fn append_suggestion(
    suggestions: &mut HashMap<String, DocumentSuggestion>,
    identity: DocumentSuggestionIdentity,
    placement: DocumentSuggestionPlacement,
) -> UseResult<()> {
    let suggestion = suggestions
        .entry(identity.id.clone())
        .or_insert_with(|| DocumentSuggestion {
            identity: identity.clone(),
            text: String::new(),
            placements: Vec::new(),
        });
    if suggestion.identity != identity {
        return Err(permission_denied(format!(
            "Document suggestion ID '{}' carries conflicting identities.",
            identity.id
        )));
    }
    suggestion.text.push_str(&placement.text);
    if let Some(previous) = suggestion.placements.last_mut() {
        let adjacent = match identity.kind {
            DocumentSuggestionKind::Insertion => {
                previous.text_node == placement.text_node
                    && previous.baseline_start_utf16 == placement.baseline_start_utf16
                    && previous.baseline_end_utf16 == placement.baseline_end_utf16
            }
            DocumentSuggestionKind::Deletion => {
                previous.text_node == placement.text_node
                    && previous.baseline_end_utf16 == placement.baseline_start_utf16
            }
        };
        if adjacent {
            previous.baseline_end_utf16 = placement.baseline_end_utf16;
            previous.text.push_str(&placement.text);
            return Ok(());
        }
    }
    suggestion.placements.push(placement);
    Ok(())
}

fn suggestion_identity(
    attributes: Option<&Attrs>,
) -> UseResult<Option<DocumentSuggestionIdentity>> {
    let Some(attributes) = attributes else {
        return Ok(None);
    };
    let mut result: Option<DocumentSuggestionIdentity> = None;
    for (attribute, value) in attributes {
        if !is_document_change_attribute(attribute) {
            continue;
        }
        let parsed = parse_suggestion_identity(value)?;
        if result.as_ref().is_some_and(|current| current != &parsed) {
            return Err(permission_denied(
                "Overlapping Document suggestion marks carry different identities.",
            ));
        }
        result = Some(parsed);
    }
    Ok(result)
}

fn parse_suggestion_identity(value: &Any) -> UseResult<DocumentSuggestionIdentity> {
    let Any::Map(fields) = value else {
        return Err(permission_denied(
            "A shared Document suggestion mark is not an object.",
        ));
    };
    let actor_id = optional_string(fields.get("actorId"), "suggestion actor ID")?;
    let expected_fields = if actor_id.is_some() { 5 } else { 4 };
    if fields.len() != expected_fields
        || fields
            .keys()
            .any(|key| !matches!(key.as_str(), "actorId" | "author" | "date" | "id" | "kind"))
    {
        return Err(permission_denied(
            "A shared Document suggestion mark contains unsupported fields.",
        ));
    }
    let id = required_string(fields.get("id"), "suggestion ID")?;
    let author = required_string(fields.get("author"), "suggestion author")?;
    let date = required_string(fields.get("date"), "suggestion date")?;
    let kind = match required_string(fields.get("kind"), "suggestion kind")?.as_str() {
        "insertion" => DocumentSuggestionKind::Insertion,
        "deletion" => DocumentSuggestionKind::Deletion,
        _ => {
            return Err(permission_denied(
                "A shared Document suggestion kind must be 'insertion' or 'deletion'.",
            ))
        }
    };
    validate_identifier(&id, "suggestion ID")?;
    validate_identifier(&author, "suggestion author")?;
    validate_identifier(&date, "suggestion date")?;
    Ok(DocumentSuggestionIdentity {
        id,
        kind,
        actor_id,
        author,
        date,
    })
}

fn assert_suggestion_changes(
    before: &HashMap<String, DocumentSuggestion>,
    candidate: &HashMap<String, DocumentSuggestion>,
    manifest: &NativeOfficeCollaborationManifest,
    actor_name: &str,
) -> UseResult<()> {
    for (id, previous) in before {
        let Some(current) = candidate.get(id) else {
            if previous.identity.actor_id.as_deref() == Some(manifest.actor_id.as_str()) {
                continue;
            }
            return Err(permission_denied(format!(
                "Suggest mode cannot withdraw another actor's Document suggestion '{id}'."
            )));
        };
        if previous.identity != current.identity {
            return Err(permission_denied(format!(
                "Suggest mode cannot rewrite the identity of Document suggestion '{id}'."
            )));
        }
        if previous.identity.actor_id.as_deref() != Some(manifest.actor_id.as_str()) {
            if previous != current {
                return Err(permission_denied(format!(
                    "Suggest mode cannot rewrite another actor's Document suggestion '{id}'."
                )));
            }
        } else if previous.identity.kind == DocumentSuggestionKind::Deletion
            && (previous.text != current.text || previous.placements != current.placements)
        {
            return Err(permission_denied(format!(
                "Suggest mode cannot rewrite the canonical text targeted by deletion suggestion '{id}'."
            )));
        }
    }

    for (id, current) in candidate {
        if before.contains_key(id) {
            continue;
        }
        if current.identity.actor_id.as_deref() != Some(manifest.actor_id.as_str())
            || current.identity.author != actor_name
        {
            return Err(permission_denied(format!(
                "A new Document suggestion '{id}' must use authenticated actor '{}' and display name '{actor_name}'.",
                manifest.actor_id
            )));
        }
        if !valid_canonical_utc_timestamp(&current.identity.date) {
            return Err(permission_denied(format!(
                "A new Document suggestion '{id}' must use a canonical UTC timestamp."
            )));
        }
    }
    Ok(())
}

fn required_root<T: ReadTxn>(transaction: &T, name: &str) -> UseResult<Out> {
    transaction
        .get(name)
        .ok_or_else(|| permission_denied(format!("Collaboration root '{name}' is missing.")))
}

fn required_content_root<T: ReadTxn>(
    transaction: &T,
    name: &str,
) -> UseResult<yrs::XmlFragmentRef> {
    match transaction.get(name) {
        Some(Out::YXmlFragment(value)) => Ok(value),
        _ => Err(permission_denied(format!(
            "Collaboration root '{name}' is not a Y.XmlFragment."
        ))),
    }
}

fn required_string(value: Option<&Any>, label: &str) -> UseResult<String> {
    match value {
        Some(Any::String(value)) => Ok(value.to_string()),
        _ => Err(permission_denied(format!(
            "A shared Document {label} is not a string."
        ))),
    }
}

fn optional_string(value: Option<&Any>, label: &str) -> UseResult<Option<String>> {
    match value {
        Some(Any::String(value)) => {
            validate_identifier(value, label)?;
            Ok(Some(value.to_string()))
        }
        Some(_) => Err(permission_denied(format!(
            "A shared Document {label} is not a string."
        ))),
        None => Ok(None),
    }
}

fn validate_identifier(value: &str, label: &str) -> UseResult<()> {
    if !value.is_empty() && value == value.trim() && value.encode_utf16().count() <= 256 {
        return Ok(());
    }
    Err(permission_denied(format!(
        "A shared Document {label} must contain 1 to 256 non-padded characters."
    )))
}

fn permission_denied(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.permission_denied", message)
}
