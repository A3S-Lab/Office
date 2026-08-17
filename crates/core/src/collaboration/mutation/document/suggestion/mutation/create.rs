use std::collections::HashMap;
use std::sync::Arc;

use a3s_use_core::UseResult;
use yrs::types::Attrs;
use yrs::{Any, Text, Transact, XmlFragment, XmlTextPrelim};

use super::super::super::identity::validate_paragraph_id_input;
use super::super::decision::read_document_change_decisions;
use super::super::{collect_suggestions, DocumentSuggestion};
use super::shared::{
    identity_conflict, insertion_attributes, match_conflict, resolve_suggestion_target,
    selection_overlaps_suggestion, validate_change_id, validate_identifier, validate_text,
    validate_timestamp,
};
use crate::collaboration::{
    collaboration_error, NativeOfficeCollaborationDocumentSuggestionKind,
    NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation,
};

pub(super) fn validate_create_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    let NativeOfficeCollaborationMutation::DocumentSuggestionCreate {
        paragraph_id,
        expected_text_id,
        start_utf16,
        end_utf16,
        expected_text,
        replacement,
        insertion_id,
        deletion_id,
        author,
        created_at,
    } = mutation
    else {
        return Err(invalid_mutation());
    };
    validate_paragraph_id_input(paragraph_id, "paragraphId")?;
    validate_paragraph_id_input(expected_text_id, "expectedTextId")?;
    if start_utf16 > end_utf16 {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "A Document suggestion requires startUtf16 to be no greater than endUtf16.",
        ));
    }
    let expected_length = validate_text(expected_text, "expectedText")?;
    let replacement_length = validate_text(replacement, "replacement")?;
    let range_length = end_utf16 - start_utf16;
    if expected_length != range_length {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "Document suggestion expectedText must have the same UTF-16 length as its guarded range.",
        )
        .with_detail("expectedTextLengthUtf16", expected_length as u64)
        .with_detail("rangeLengthUtf16", range_length as u64));
    }
    if range_length == 0 && replacement_length == 0 {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "A Document suggestion must insert or delete at least one character.",
        ));
    }
    validate_optional_change_id(
        insertion_id.as_deref(),
        replacement_length > 0,
        "insertionId",
        "non-empty replacement",
    )?;
    validate_optional_change_id(
        deletion_id.as_deref(),
        range_length > 0,
        "deletionId",
        "non-empty guarded range",
    )?;
    if insertion_id.is_some() && insertion_id == deletion_id {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "A replacement suggestion requires different insertionId and deletionId values.",
        ));
    }
    validate_identifier(author, "author")?;
    validate_timestamp(created_at, "createdAt")
}

pub(super) fn apply_create_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    let NativeOfficeCollaborationMutation::DocumentSuggestionCreate {
        paragraph_id,
        expected_text_id,
        start_utf16,
        end_utf16,
        expected_text,
        replacement,
        insertion_id,
        deletion_id,
        author,
        created_at,
    } = mutation
    else {
        return Err(invalid_mutation());
    };
    let root = format!("{}.document.content", manifest.namespace);
    let fragment = doc.get_or_insert_xml_fragment(root);
    let transaction = doc.transact();
    let target = resolve_suggestion_target(
        &transaction,
        &fragment,
        paragraph_id,
        expected_text_id,
        *start_utf16,
        *end_utf16,
        expected_text,
    )?;
    let suggestions = collect_suggestions(&transaction, &fragment)?;
    let decisions = read_document_change_decisions(&transaction, manifest)?;
    let expected = expected_suggestions(
        insertion_id.as_deref(),
        deletion_id.as_deref(),
        manifest,
        author,
        created_at,
        expected_text,
        replacement,
        paragraph_id,
        expected_text_id,
        *start_utf16,
        *end_utf16,
    )?;
    let has_existing_suggestion = expected
        .iter()
        .any(|candidate| suggestions.contains_key(candidate.id));
    let has_existing_decision = expected.iter().any(|candidate| {
        decisions.records.values().any(|decision| {
            decision.change_id == candidate.id && decision.change_kind == candidate.kind
        })
    });
    if has_existing_suggestion || has_existing_decision {
        if !has_existing_decision
            && expected.iter().all(|candidate| {
                suggestions
                    .get(candidate.id)
                    .is_some_and(|current| candidate.matches(current))
            })
        {
            return Ok(());
        }
        return Err(identity_conflict(
            "One or more Document suggestion IDs were already claimed by a different or final tracked change.",
        ));
    }
    if selection_overlaps_suggestion(&target, *start_utf16, *end_utf16) {
        return Err(match_conflict(
            "The guarded Document range overlaps an existing tracked-change mark.",
        ));
    }
    let mut inserted_attributes = insertion_attributes(&target, *end_utf16);
    if let Some(insertion_id) = insertion_id {
        inserted_attributes.insert(
            "documentChange".into(),
            suggestion_mark_value(
                insertion_id,
                NativeOfficeCollaborationDocumentSuggestionKind::Insertion,
                manifest,
                author,
                created_at,
            ),
        );
    }
    let deletion_attributes = deletion_id.as_deref().map(|deletion_id| {
        Attrs::from([(
            "documentChange".into(),
            suggestion_mark_value(
                deletion_id,
                NativeOfficeCollaborationDocumentSuggestionKind::Deletion,
                manifest,
                author,
                created_at,
            ),
        )])
    });
    drop(transaction);

    let mut transaction = doc.transact_mut();
    let text = match target.text {
        Some(text) => text,
        None => target
            .paragraph
            .push_back(&mut transaction, XmlTextPrelim::default()),
    };
    if let Some(attributes) = deletion_attributes {
        text.format(
            &mut transaction,
            *start_utf16,
            *end_utf16 - *start_utf16,
            attributes,
        );
    }
    if !replacement.is_empty() {
        text.insert_with_attributes(
            &mut transaction,
            *end_utf16,
            replacement,
            inserted_attributes,
        );
    }
    drop(transaction);

    let transaction = doc.transact();
    let written = collect_suggestions(&transaction, &fragment)?;
    read_document_change_decisions(&transaction, manifest)?;
    if expected.iter().all(|candidate| {
        written
            .get(candidate.id)
            .is_some_and(|current| candidate.matches(current))
    }) {
        Ok(())
    } else {
        Err(collaboration_error(
            "office.collaboration.content_invalid",
            "The native Document suggestion did not round-trip through the browser-compatible Yjs schema.",
        ))
    }
}

struct ExpectedSuggestion<'a> {
    id: &'a str,
    kind: NativeOfficeCollaborationDocumentSuggestionKind,
    actor_id: &'a str,
    author: &'a str,
    created_at: &'a str,
    text: &'a str,
    paragraph_id: &'a str,
    text_id: &'a str,
    start_utf16: u32,
    end_utf16: u32,
}

impl ExpectedSuggestion<'_> {
    fn matches(&self, current: &DocumentSuggestion) -> bool {
        current.identity.id == self.id
            && current.identity.kind == self.kind
            && current.identity.actor_id.as_deref() == Some(self.actor_id)
            && current.identity.author == self.author
            && current.identity.date == self.created_at
            && current.text == self.text
            && matches!(
                current.placements.as_slice(),
                [placement]
                    if placement.paragraph_id.as_deref() == Some(self.paragraph_id)
                        && placement.text_id.as_deref() == Some(self.text_id)
                        && placement.start_utf16 == self.start_utf16
                        && placement.end_utf16 == self.end_utf16
                        && placement.text == self.text
            )
    }
}

#[allow(clippy::too_many_arguments)]
fn expected_suggestions<'a>(
    insertion_id: Option<&'a str>,
    deletion_id: Option<&'a str>,
    manifest: &'a NativeOfficeCollaborationManifest,
    author: &'a str,
    created_at: &'a str,
    expected_text: &'a str,
    replacement: &'a str,
    paragraph_id: &'a str,
    text_id: &'a str,
    start_utf16: u32,
    end_utf16: u32,
) -> UseResult<Vec<ExpectedSuggestion<'a>>> {
    let mut expected = Vec::with_capacity(2);
    if let Some(deletion_id) = deletion_id {
        expected.push(ExpectedSuggestion {
            id: deletion_id,
            kind: NativeOfficeCollaborationDocumentSuggestionKind::Deletion,
            actor_id: &manifest.actor_id,
            author,
            created_at,
            text: expected_text,
            paragraph_id,
            text_id,
            start_utf16,
            end_utf16,
        });
    }
    if let Some(insertion_id) = insertion_id {
        let replacement_length =
            u32::try_from(replacement.encode_utf16().count()).map_err(|_| {
                collaboration_error(
                    "office.collaboration.mutation_too_large",
                    "The Document suggestion replacement is too large.",
                )
            })?;
        expected.push(ExpectedSuggestion {
            id: insertion_id,
            kind: NativeOfficeCollaborationDocumentSuggestionKind::Insertion,
            actor_id: &manifest.actor_id,
            author,
            created_at,
            text: replacement,
            paragraph_id,
            text_id,
            start_utf16: end_utf16,
            end_utf16: end_utf16.checked_add(replacement_length).ok_or_else(|| {
                collaboration_error(
                    "office.collaboration.mutation_too_large",
                    "The Document suggestion insertion offset is too large.",
                )
            })?,
        });
    }
    Ok(expected)
}

fn suggestion_mark_value(
    id: &str,
    kind: NativeOfficeCollaborationDocumentSuggestionKind,
    manifest: &NativeOfficeCollaborationManifest,
    author: &str,
    created_at: &str,
) -> Any {
    let mut fields = HashMap::new();
    fields.insert(
        "actorId".to_owned(),
        Any::String(manifest.actor_id.clone().into()),
    );
    fields.insert("author".to_owned(), Any::String(author.into()));
    fields.insert("date".to_owned(), Any::String(created_at.into()));
    fields.insert("id".to_owned(), Any::String(id.into()));
    fields.insert("kind".to_owned(), Any::String(kind.as_str().into()));
    Any::Map(Arc::new(fields))
}

fn validate_optional_change_id(
    value: Option<&str>,
    required: bool,
    field: &str,
    condition: &str,
) -> UseResult<()> {
    match (value, required) {
        (Some(value), true) => validate_change_id(value, field),
        (None, false) => Ok(()),
        (None, true) => Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            format!("Document suggestion field '{field}' is required for a {condition}."),
        )),
        (Some(_), false) => Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            format!("Document suggestion field '{field}' must be absent without a {condition}."),
        )),
    }
}

fn invalid_mutation() -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.mutation_invalid",
        "The supplied mutation is not a Document suggestion creation.",
    )
}
