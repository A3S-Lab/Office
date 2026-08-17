use std::collections::HashSet;

use a3s_use_core::UseResult;
use yrs::types::Attrs;
use yrs::{Any, Array, Text, Transact, Xml, XmlElementRef, XmlTextRef};

use super::super::super::identity::{
    ancestor_table_rows, paragraph_text_id_rotations, table_row_text_id_rotations,
    ROW_TEXT_ID_ATTRIBUTE, TEXT_ID_ATTRIBUTE,
};
use super::super::decision::{
    decision_record_id, encoded_decision_claim, insert_decision_record,
    read_document_change_decisions,
};
use super::super::{collect_suggestions, DocumentSuggestion};
use super::shared::{
    collect_live_suggestion_segments, content_invalid, identity_conflict, match_conflict,
    validate_change_id, validate_identifier, validate_text, validate_timestamp,
    LiveSuggestionSegment, MAX_DOCUMENT_SUGGESTION_DECISIONS,
};
use crate::collaboration::{
    collaboration_error, NativeOfficeCollaborationDocumentChangeDecision,
    NativeOfficeCollaborationDocumentSuggestionDecision,
    NativeOfficeCollaborationDocumentSuggestionKind,
    NativeOfficeCollaborationDocumentSuggestionMatch, NativeOfficeCollaborationManifest,
    NativeOfficeCollaborationMutation,
};

pub(super) fn validate_decide_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    let NativeOfficeCollaborationMutation::DocumentSuggestionDecide {
        suggestions,
        decided_by,
        decided_at,
        ..
    } = mutation
    else {
        return Err(invalid_mutation());
    };
    if suggestions.is_empty() || suggestions.len() > MAX_DOCUMENT_SUGGESTION_DECISIONS {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            format!(
                "A Document suggestion decision requires 1 to {MAX_DOCUMENT_SUGGESTION_DECISIONS} exact suggestions."
            ),
        ));
    }
    let mut ids = HashSet::new();
    for suggestion in suggestions {
        validate_change_id(&suggestion.id, "suggestions[].id")?;
        if !ids.insert(suggestion.id.as_str()) {
            return Err(collaboration_error(
                "office.collaboration.mutation_invalid",
                format!(
                    "Document suggestion ID '{}' is repeated in one decision batch.",
                    suggestion.id
                ),
            ));
        }
        if let Some(actor_id) = &suggestion.expected_actor_id {
            validate_identifier(actor_id, "suggestions[].expectedActorId")?;
        }
        validate_identifier(&suggestion.expected_author, "suggestions[].expectedAuthor")?;
        validate_timestamp(
            &suggestion.expected_created_at,
            "suggestions[].expectedCreatedAt",
        )?;
        if validate_text(&suggestion.expected_text, "suggestions[].expectedText")? == 0 {
            return Err(collaboration_error(
                "office.collaboration.mutation_invalid",
                "A Document suggestion decision must match non-empty tracked text.",
            ));
        }
    }
    validate_identifier(decided_by, "decidedBy")?;
    validate_timestamp(decided_at, "decidedAt")
}

pub(super) fn apply_decide_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    let NativeOfficeCollaborationMutation::DocumentSuggestionDecide {
        suggestions,
        decision,
        decided_by,
        decided_at,
    } = mutation
    else {
        return Err(invalid_mutation());
    };
    let fragment =
        doc.get_or_insert_xml_fragment(format!("{}.document.content", manifest.namespace));
    let transaction = doc.transact();
    let current = collect_suggestions(&transaction, &fragment)?;
    let decision_state = read_document_change_decisions(&transaction, manifest)?;
    let live_segments = collect_live_suggestion_segments(&transaction, &fragment)?;
    let mut pending = Vec::new();
    for expected in suggestions {
        let candidate = decision_record(manifest, expected, *decision, decided_by, decided_at);
        let existing = decision_state
            .records
            .values()
            .find(|record| record.change_id == expected.id && record.change_kind == expected.kind);
        if let Some(existing) = existing {
            if existing != &candidate {
                return Err(identity_conflict(format!(
                    "Tracked change '{}:{}' already has a different final decision.",
                    expected.kind.as_str(),
                    expected.id
                )));
            }
            if current.contains_key(&expected.id) {
                return Err(content_invalid(format!(
                    "Tracked change '{}:{}' has both a final decision and a live mark.",
                    expected.kind.as_str(),
                    expected.id
                )));
            }
            continue;
        }
        if decision_state.records.contains_key(&candidate.id) {
            return Err(identity_conflict(format!(
                "Tracked-change decision ID '{}' belongs to another record.",
                candidate.id
            )));
        }
        if decision_state
            .claims
            .claim_for("change-decision", None, &candidate.id)
            .is_some()
        {
            return Err(identity_conflict(format!(
                "Tracked-change decision ID '{}' already has an immutable claim.",
                candidate.id
            )));
        }
        let live = current.get(&expected.id).ok_or_else(|| {
            match_conflict(format!(
                "Document suggestion '{}:{}' is no longer live.",
                expected.kind.as_str(),
                expected.id
            ))
        })?;
        if !matches_exact_suggestion(live, expected) {
            return Err(match_conflict(format!(
                "Document suggestion '{}:{}' no longer matches the reviewed identity and text.",
                expected.kind.as_str(),
                expected.id
            )));
        }
        pending.push((expected, candidate));
    }
    if pending.is_empty() {
        return Ok(());
    }

    let targeted = |segment: &LiveSuggestionSegment| {
        pending.iter().any(|(expected, _)| {
            segment.identity.id == expected.id && segment.identity.kind == expected.kind
        })
    };
    for (expected, _) in &pending {
        if !live_segments.iter().any(|segment| {
            segment.identity.id == expected.id && segment.identity.kind == expected.kind
        }) {
            return Err(content_invalid(format!(
                "Document suggestion '{}:{}' has no live Yjs text segment.",
                expected.kind.as_str(),
                expected.id
            )));
        }
    }

    let mut paragraphs = Vec::<XmlElementRef>::new();
    for segment in live_segments.iter().filter(|segment| targeted(segment)) {
        if let Some(paragraph) = &segment.paragraph {
            if !paragraphs.contains(paragraph) {
                paragraphs.push(paragraph.clone());
            }
        }
    }
    let paragraph_rotations = paragraph_text_id_rotations(&paragraphs, &transaction)?;
    let table_rows = ancestor_table_rows(&paragraphs)?;
    let row_rotations = table_row_text_id_rotations(&table_rows, &transaction)?;
    let rewrites = attribute_rewrites(&live_segments, &targeted, &transaction);
    let deletions = deletion_groups(&live_segments, &targeted, *decision);
    let records_root = decision_state.records_root.clone();
    let order_root = decision_state.order_root.clone();
    let claims_root = decision_state.claims_root.clone();
    let records = pending
        .into_iter()
        .map(|(_, record)| record)
        .collect::<Vec<_>>();
    let claims = records
        .iter()
        .map(encoded_decision_claim)
        .collect::<UseResult<Vec<_>>>()?;
    drop(transaction);

    let mut transaction = doc.transact_mut();
    // Clearing an interior attribute written by another Yjs client can panic
    // in Yrs. Clear the key across the whole text and then restore every
    // non-target span that used that key before applying descending deletes.
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
    for rewrite in &rewrites {
        for retained in &rewrite.retained {
            rewrite.text.format(
                &mut transaction,
                retained.start_utf16,
                retained.end_utf16 - retained.start_utf16,
                Attrs::from([(rewrite.attribute.clone().into(), retained.value.clone())]),
            );
        }
    }
    for group in deletions {
        for (start, end) in group.ranges.into_iter().rev() {
            group
                .text
                .remove_range(&mut transaction, start, end - start);
        }
    }
    for (paragraph, next_text_id) in paragraph_rotations {
        paragraph.insert_attribute(&mut transaction, TEXT_ID_ATTRIBUTE, next_text_id);
    }
    for (row, next_text_id) in row_rotations {
        row.insert_attribute(&mut transaction, ROW_TEXT_ID_ATTRIBUTE, next_text_id);
    }
    for (record, claim) in records.iter().zip(claims) {
        claims_root.push_back(&mut transaction, claim);
        insert_decision_record(&mut transaction, &records_root, &order_root, record);
    }
    drop(transaction);

    let transaction = doc.transact();
    let remaining = collect_suggestions(&transaction, &fragment)?;
    let written = read_document_change_decisions(&transaction, manifest)?;
    if records.iter().any(|record| {
        remaining
            .get(&record.change_id)
            .is_some_and(|suggestion| suggestion.identity.kind == record.change_kind)
            || written.records.get(&record.id) != Some(record)
    }) {
        return Err(content_invalid(
            "A native Document tracked-change decision did not round-trip through the browser-compatible Yjs schema."
                .to_owned(),
        ));
    }
    Ok(())
}

fn matches_exact_suggestion(
    current: &DocumentSuggestion,
    expected: &NativeOfficeCollaborationDocumentSuggestionMatch,
) -> bool {
    current.identity.id == expected.id
        && current.identity.kind == expected.kind
        && current.identity.actor_id == expected.expected_actor_id
        && current.identity.author == expected.expected_author
        && current.identity.date == expected.expected_created_at
        && current.text == expected.expected_text
}

fn decision_record(
    manifest: &NativeOfficeCollaborationManifest,
    suggestion: &NativeOfficeCollaborationDocumentSuggestionMatch,
    decision: NativeOfficeCollaborationDocumentSuggestionDecision,
    decided_by: &str,
    decided_at: &str,
) -> NativeOfficeCollaborationDocumentChangeDecision {
    NativeOfficeCollaborationDocumentChangeDecision {
        id: decision_record_id(suggestion.kind, &suggestion.id),
        change_id: suggestion.id.clone(),
        change_kind: suggestion.kind,
        suggested_by_actor_id: suggestion.expected_actor_id.clone(),
        suggested_by: suggestion.expected_author.clone(),
        suggested_at: suggestion.expected_created_at.clone(),
        text: suggestion.expected_text.clone(),
        decision,
        decided_by_actor_id: Some(manifest.actor_id.clone()),
        decided_by: decided_by.to_owned(),
        decided_at: decided_at.to_owned(),
    }
}

struct AttributeRewrite {
    text: XmlTextRef,
    attribute: String,
    length_utf16: u32,
    retained: Vec<RetainedAttribute>,
}

struct RetainedAttribute {
    start_utf16: u32,
    end_utf16: u32,
    value: Any,
}

fn attribute_rewrites<T: yrs::ReadTxn>(
    segments: &[LiveSuggestionSegment],
    targeted: &impl Fn(&LiveSuggestionSegment) -> bool,
    transaction: &T,
) -> Vec<AttributeRewrite> {
    let mut rewrites = Vec::<AttributeRewrite>::new();
    for segment in segments.iter().filter(|segment| targeted(segment)) {
        for (attribute, _) in &segment.attributes {
            if !rewrites
                .iter()
                .any(|rewrite| rewrite.text == segment.text && rewrite.attribute == *attribute)
            {
                rewrites.push(AttributeRewrite {
                    text: segment.text.clone(),
                    attribute: attribute.clone(),
                    length_utf16: segment.text.len(transaction),
                    retained: Vec::new(),
                });
            }
        }
    }
    for segment in segments.iter().filter(|segment| !targeted(segment)) {
        for (attribute, value) in &segment.attributes {
            if let Some(rewrite) = rewrites
                .iter_mut()
                .find(|rewrite| rewrite.text == segment.text && rewrite.attribute == *attribute)
            {
                rewrite.retained.push(RetainedAttribute {
                    start_utf16: segment.start_utf16,
                    end_utf16: segment.end_utf16,
                    value: value.clone(),
                });
            }
        }
    }
    rewrites
}

struct DeletionGroup {
    text: XmlTextRef,
    ranges: Vec<(u32, u32)>,
}

fn deletion_groups(
    segments: &[LiveSuggestionSegment],
    targeted: &impl Fn(&LiveSuggestionSegment) -> bool,
    decision: NativeOfficeCollaborationDocumentSuggestionDecision,
) -> Vec<DeletionGroup> {
    let mut groups = Vec::<DeletionGroup>::new();
    for segment in segments
        .iter()
        .filter(|segment| targeted(segment) && should_delete(segment.identity.kind, decision))
    {
        let group = if let Some(index) = groups.iter().position(|group| group.text == segment.text)
        {
            &mut groups[index]
        } else {
            groups.push(DeletionGroup {
                text: segment.text.clone(),
                ranges: Vec::new(),
            });
            groups.last_mut().expect("just inserted deletion group")
        };
        group.ranges.push((segment.start_utf16, segment.end_utf16));
    }
    for group in &mut groups {
        group.ranges.sort_unstable();
        let mut merged = Vec::<(u32, u32)>::new();
        for (start, end) in group.ranges.drain(..) {
            if let Some(previous) = merged.last_mut() {
                if start <= previous.1 {
                    previous.1 = previous.1.max(end);
                    continue;
                }
            }
            merged.push((start, end));
        }
        group.ranges = merged;
    }
    groups
}

fn should_delete(
    kind: NativeOfficeCollaborationDocumentSuggestionKind,
    decision: NativeOfficeCollaborationDocumentSuggestionDecision,
) -> bool {
    matches!(
        (kind, decision),
        (
            NativeOfficeCollaborationDocumentSuggestionKind::Deletion,
            NativeOfficeCollaborationDocumentSuggestionDecision::Accept
        ) | (
            NativeOfficeCollaborationDocumentSuggestionKind::Insertion,
            NativeOfficeCollaborationDocumentSuggestionDecision::Reject
        )
    )
}

fn invalid_mutation() -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.mutation_invalid",
        "The supplied mutation is not a Document suggestion decision.",
    )
}
