use yrs::updates::decoder::Decode;
use yrs::{Transact, Update};

use super::*;

const SUGGESTER_ID: &str = "agent-suggester";
const SUGGESTER_NAME: &str = "A3S Agent";
const EDITOR_ID: &str = "human-editor";
const EDITOR_NAME: &str = "Grace Editor";
const CREATED_AT: &str = "2026-08-17T10:00:00.000Z";
const DECIDED_AT: &str = "2026-08-17T10:01:00.000Z";
const INSERTION_ID: &str = "native-replacement-insertion";
const DELETION_ID: &str = "native-replacement-deletion";

#[test]
fn native_suggestion_mutation_creates_and_projects_an_atomic_replacement() {
    let temp = tempfile::tempdir().unwrap();
    let store = suggestion_store(
        &temp.path().join("native-suggestion-create"),
        NativeOfficeCollaborationMode::Suggest,
        SUGGESTER_ID,
        901_001,
    );
    bootstrap(&store, NativeOfficeCollaborationMode::Suggest, SUGGESTER_ID);

    let result = store
        .mutate(suggestion_request(
            "native-suggestion-replacement",
            NativeOfficeCollaborationMode::Suggest,
            SUGGESTER_ID,
            replacement_mutation(),
        ))
        .unwrap();
    assert!(result.state_changed);

    let projection = store.project().unwrap();
    assert_eq!(projection.version, 4);
    let NativeOfficeCollaborationProjectedContent::Document {
        plain_text,
        paragraphs,
        suggestions,
        change_decisions,
        ..
    } = &projection.content
    else {
        panic!("expected Document projection");
    };
    assert_eq!(plain_text, "Hello 😀collaborative world");
    assert_eq!(paragraphs[0].text_id.as_deref(), Some("00000002"));
    assert!(paragraphs[0].has_review_marks);
    assert!(change_decisions.is_empty());
    assert_eq!(suggestions.len(), 2);
    assert_eq!(
        suggestions
            .iter()
            .map(|suggestion| (
                suggestion.id.as_str(),
                suggestion.kind,
                suggestion.actor_id.as_deref(),
                suggestion.author.as_str(),
                suggestion.created_at.as_str(),
                suggestion.text.as_str(),
            ))
            .collect::<Vec<_>>(),
        vec![
            (
                DELETION_ID,
                NativeOfficeCollaborationDocumentSuggestionKind::Deletion,
                Some(SUGGESTER_ID),
                SUGGESTER_NAME,
                CREATED_AT,
                "😀",
            ),
            (
                INSERTION_ID,
                NativeOfficeCollaborationDocumentSuggestionKind::Insertion,
                Some(SUGGESTER_ID),
                SUGGESTER_NAME,
                CREATED_AT,
                "collaborative",
            ),
        ]
    );
    assert_eq!(suggestions[0].placements[0].start_utf16, 6);
    assert_eq!(suggestions[0].placements[0].end_utf16, 8);
    assert_eq!(suggestions[1].placements[0].start_utf16, 8);
    assert_eq!(suggestions[1].placements[0].end_utf16, 21);

    let retry = store
        .mutate(suggestion_request(
            "native-suggestion-replacement-stable-id-retry",
            NativeOfficeCollaborationMode::Suggest,
            SUGGESTER_ID,
            replacement_mutation(),
        ))
        .unwrap();
    assert!(!retry.state_changed);
}

#[test]
fn native_suggestion_decision_accepts_replacement_and_appends_immutable_audit() {
    let temp = tempfile::tempdir().unwrap();
    let suggester = suggestion_store(
        &temp.path().join("native-suggestion-producer"),
        NativeOfficeCollaborationMode::Suggest,
        SUGGESTER_ID,
        901_002,
    );
    bootstrap(
        &suggester,
        NativeOfficeCollaborationMode::Suggest,
        SUGGESTER_ID,
    );
    suggester
        .mutate(suggestion_request(
            "produce-native-replacement",
            NativeOfficeCollaborationMode::Suggest,
            SUGGESTER_ID,
            replacement_mutation(),
        ))
        .unwrap();

    let editor_root = temp.path().join("native-suggestion-editor");
    let editor = suggestion_store(
        &editor_root,
        NativeOfficeCollaborationMode::Edit,
        EDITOR_ID,
        901_003,
    );
    editor
        .apply(suggestion_apply_request(
            "synchronize-native-replacement",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            suggester.synchronize(None).unwrap().update,
        ))
        .unwrap();

    let result = editor
        .mutate(suggestion_request(
            "accept-native-replacement",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            decision_mutation(NativeOfficeCollaborationDocumentSuggestionDecision::Accept),
        ))
        .unwrap();
    assert!(result.state_changed);

    let projection = editor.project().unwrap();
    let NativeOfficeCollaborationProjectedContent::Document {
        plain_text,
        paragraphs,
        suggestions,
        change_decisions,
        ..
    } = &projection.content
    else {
        panic!("expected Document projection");
    };
    assert_eq!(plain_text, "Hello collaborative world");
    assert_eq!(paragraphs[0].text_id.as_deref(), Some("00000003"));
    assert!(suggestions.is_empty());
    assert_eq!(change_decisions.len(), 2);
    assert!(change_decisions.iter().all(|record| {
        record.decision == NativeOfficeCollaborationDocumentSuggestionDecision::Accept
            && record.decided_by_actor_id.as_deref() == Some(EDITOR_ID)
            && record.decided_by == EDITOR_NAME
            && record.decided_at == DECIDED_AT
    }));

    drop(editor);
    let reopened = NativeOfficeCollaborationStore::open(&editor_root).unwrap();
    assert_eq!(reopened.project().unwrap(), projection);
    let duplicate = reopened
        .mutate(suggestion_request(
            "accept-native-replacement-idempotent",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            decision_mutation(NativeOfficeCollaborationDocumentSuggestionDecision::Accept),
        ))
        .unwrap();
    assert!(!duplicate.state_changed);
}

#[test]
fn native_suggestion_mutations_enforce_review_modes_and_exact_matches() {
    let temp = tempfile::tempdir().unwrap();
    let suggester = suggestion_store(
        &temp.path().join("native-suggestion-mode"),
        NativeOfficeCollaborationMode::Suggest,
        SUGGESTER_ID,
        901_004,
    );
    bootstrap(
        &suggester,
        NativeOfficeCollaborationMode::Suggest,
        SUGGESTER_ID,
    );
    let mut stale = replacement_mutation();
    let NativeOfficeCollaborationMutation::DocumentSuggestionCreate {
        expected_text_id, ..
    } = &mut stale
    else {
        unreachable!("replacement fixture must create a suggestion")
    };
    *expected_text_id = "00000009".to_owned();
    let before = suggester.inspect().unwrap().document_state_sha256;
    let conflict = suggester
        .mutate(suggestion_request(
            "stale-native-suggestion",
            NativeOfficeCollaborationMode::Suggest,
            SUGGESTER_ID,
            stale,
        ))
        .unwrap_err();
    assert_eq!(
        conflict.code,
        "office.collaboration.mutation_match_conflict"
    );
    assert_eq!(suggester.inspect().unwrap().document_state_sha256, before);

    let forbidden = suggester
        .mutate(suggestion_request(
            "suggest-mode-decision",
            NativeOfficeCollaborationMode::Suggest,
            SUGGESTER_ID,
            decision_mutation(NativeOfficeCollaborationDocumentSuggestionDecision::Reject),
        ))
        .unwrap_err();
    assert_eq!(forbidden.code, "office.collaboration.mutation_forbidden");

    let editor = suggestion_store(
        &temp.path().join("native-suggestion-edit-mode"),
        NativeOfficeCollaborationMode::Edit,
        EDITOR_ID,
        901_005,
    );
    bootstrap(&editor, NativeOfficeCollaborationMode::Edit, EDITOR_ID);
    let forbidden = editor
        .mutate(suggestion_request(
            "edit-mode-suggestion-create",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            replacement_mutation(),
        ))
        .unwrap_err();
    assert_eq!(forbidden.code, "office.collaboration.mutation_forbidden");
}

#[test]
fn native_suggestion_rejection_is_exact_final_and_restores_the_baseline() {
    let temp = tempfile::tempdir().unwrap();
    let suggester = suggestion_store(
        &temp.path().join("native-suggestion-reject-producer"),
        NativeOfficeCollaborationMode::Suggest,
        SUGGESTER_ID,
        901_006,
    );
    bootstrap(
        &suggester,
        NativeOfficeCollaborationMode::Suggest,
        SUGGESTER_ID,
    );
    suggester
        .mutate(suggestion_request(
            "produce-native-rejection",
            NativeOfficeCollaborationMode::Suggest,
            SUGGESTER_ID,
            replacement_mutation(),
        ))
        .unwrap();

    let editor = suggestion_store(
        &temp.path().join("native-suggestion-reject-editor"),
        NativeOfficeCollaborationMode::Edit,
        EDITOR_ID,
        901_007,
    );
    editor
        .apply(suggestion_apply_request(
            "synchronize-native-rejection",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            suggester.synchronize(None).unwrap().update,
        ))
        .unwrap();

    let mut stale = decision_mutation(NativeOfficeCollaborationDocumentSuggestionDecision::Reject);
    let NativeOfficeCollaborationMutation::DocumentSuggestionDecide { suggestions, .. } =
        &mut stale
    else {
        unreachable!("decision fixture must decide suggestions")
    };
    suggestions[0].expected_author = "Stale Reviewer".to_owned();
    let before = editor.inspect().unwrap().document_state_sha256;
    let conflict = editor
        .mutate(suggestion_request(
            "reject-stale-native-replacement",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            stale,
        ))
        .unwrap_err();
    assert_eq!(
        conflict.code,
        "office.collaboration.mutation_match_conflict"
    );
    assert_eq!(editor.inspect().unwrap().document_state_sha256, before);

    editor
        .mutate(suggestion_request(
            "reject-native-replacement",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            decision_mutation(NativeOfficeCollaborationDocumentSuggestionDecision::Reject),
        ))
        .unwrap();
    let projection = editor.project().unwrap();
    let NativeOfficeCollaborationProjectedContent::Document {
        plain_text,
        paragraphs,
        suggestions,
        change_decisions,
        ..
    } = &projection.content
    else {
        panic!("expected Document projection");
    };
    assert_eq!(plain_text, "Hello 😀 world");
    assert_eq!(paragraphs[0].text_id.as_deref(), Some("00000003"));
    assert!(suggestions.is_empty());
    assert_eq!(change_decisions.len(), 2);
    assert!(change_decisions.iter().all(|record| {
        record.decision == NativeOfficeCollaborationDocumentSuggestionDecision::Reject
    }));

    let finalized = editor.inspect().unwrap().document_state_sha256;
    let conflict = editor
        .mutate(suggestion_request(
            "accept-final-native-replacement",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            decision_mutation(NativeOfficeCollaborationDocumentSuggestionDecision::Accept),
        ))
        .unwrap_err();
    assert_eq!(
        conflict.code,
        "office.collaboration.mutation_identity_conflict"
    );
    assert_eq!(editor.inspect().unwrap().document_state_sha256, finalized);
}

fn replacement_mutation() -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::DocumentSuggestionCreate {
        paragraph_id: "00000001".to_owned(),
        expected_text_id: "00000002".to_owned(),
        start_utf16: 6,
        end_utf16: 8,
        expected_text: "😀".to_owned(),
        replacement: "collaborative".to_owned(),
        insertion_id: Some(INSERTION_ID.to_owned()),
        deletion_id: Some(DELETION_ID.to_owned()),
        author: SUGGESTER_NAME.to_owned(),
        created_at: CREATED_AT.to_owned(),
    }
}

fn decision_mutation(
    decision: NativeOfficeCollaborationDocumentSuggestionDecision,
) -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::DocumentSuggestionDecide {
        suggestions: vec![
            suggestion_match(
                DELETION_ID,
                NativeOfficeCollaborationDocumentSuggestionKind::Deletion,
                "😀",
            ),
            suggestion_match(
                INSERTION_ID,
                NativeOfficeCollaborationDocumentSuggestionKind::Insertion,
                "collaborative",
            ),
        ],
        decision,
        decided_by: EDITOR_NAME.to_owned(),
        decided_at: DECIDED_AT.to_owned(),
    }
}

fn suggestion_match(
    id: &str,
    kind: NativeOfficeCollaborationDocumentSuggestionKind,
    text: &str,
) -> NativeOfficeCollaborationDocumentSuggestionMatch {
    NativeOfficeCollaborationDocumentSuggestionMatch {
        id: id.to_owned(),
        kind,
        expected_actor_id: Some(SUGGESTER_ID.to_owned()),
        expected_author: SUGGESTER_NAME.to_owned(),
        expected_created_at: CREATED_AT.to_owned(),
        expected_text: text.to_owned(),
    }
}

fn suggestion_store(
    root: &std::path::Path,
    mode: NativeOfficeCollaborationMode,
    actor_id: &str,
    client_id: u64,
) -> NativeOfficeCollaborationStore {
    NativeOfficeCollaborationStore::create(NativeOfficeCollaborationCreateRequest {
        store: root.to_path_buf(),
        artifact_id: "fixture-document".to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Document,
        actor_id: actor_id.to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Agent,
        mode,
        operation_id: format!("create-{client_id}"),
        namespace: None,
        client_id: Some(client_id),
        initial_update: None,
    })
    .unwrap()
}

fn bootstrap(
    store: &NativeOfficeCollaborationStore,
    mode: NativeOfficeCollaborationMode,
    actor_id: &str,
) {
    store
        .apply(suggestion_apply_request(
            "bootstrap-document",
            mode,
            actor_id,
            STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
}

fn suggestion_apply_request(
    operation_id: &str,
    mode: NativeOfficeCollaborationMode,
    actor_id: &str,
    update: Vec<u8>,
) -> NativeOfficeCollaborationApplyRequest {
    NativeOfficeCollaborationApplyRequest {
        operation_id: operation_id.to_owned(),
        actor_id: actor_id.to_owned(),
        mode,
        expected_artifact_id: "fixture-document".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Document,
        update,
        if_state_vector: None,
        origin: None,
    }
}

fn suggestion_request(
    operation_id: &str,
    mode: NativeOfficeCollaborationMode,
    actor_id: &str,
    mutation: NativeOfficeCollaborationMutation,
) -> NativeOfficeCollaborationMutationRequest {
    NativeOfficeCollaborationMutationRequest {
        operation_id: operation_id.to_owned(),
        actor_id: actor_id.to_owned(),
        mode,
        expected_artifact_id: "fixture-document".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Document,
        mutation,
        if_state_vector: None,
    }
}

#[test]
fn reordered_suggestion_decisions_converge_or_write_nothing() {
    let temp = tempfile::tempdir().unwrap();
    let producer = suggestion_store(
        &temp.path().join("decision-producer"),
        NativeOfficeCollaborationMode::Suggest,
        SUGGESTER_ID,
        901_110,
    );
    bootstrap(
        &producer,
        NativeOfficeCollaborationMode::Suggest,
        SUGGESTER_ID,
    );
    producer
        .mutate(suggestion_request(
            "suggest-bang",
            NativeOfficeCollaborationMode::Suggest,
            SUGGESTER_ID,
            insertion_mutation(14, "suggestion-offline-bang", "!"),
        ))
        .unwrap();
    producer
        .mutate(suggestion_request(
            "suggest-question",
            NativeOfficeCollaborationMode::Suggest,
            SUGGESTER_ID,
            insertion_mutation(0, "suggestion-offline-question", "?"),
        ))
        .unwrap();
    let base = producer.synchronize(None).unwrap().update;

    let accept = decision_update(
        &temp,
        901_111,
        "accept-bang",
        &base,
        decision_for(
            "suggestion-offline-bang",
            "!",
            NativeOfficeCollaborationDocumentSuggestionDecision::Accept,
        ),
    );
    let reject = decision_update(
        &temp,
        901_112,
        "reject-question",
        &base,
        decision_for(
            "suggestion-offline-question",
            "?",
            NativeOfficeCollaborationDocumentSuggestionDecision::Reject,
        ),
    );

    let mut hashes = Vec::new();
    for (index, order) in [[0_usize, 1], [1, 0]].into_iter().enumerate() {
        let updates = [accept.clone(), reject.clone()];
        let store = deliver_decisions(
            &temp.path().join(format!("decision-order-{index}")),
            901_200 + index as u64,
            &base,
            &[updates[order[0]].clone(), updates[order[1]].clone()],
        );
        let projection = store.project().unwrap();
        let NativeOfficeCollaborationProjectedContent::Document {
            plain_text,
            paragraphs,
            suggestions,
            change_decisions,
            ..
        } = projection.content
        else {
            panic!("expected Document projection");
        };
        assert_eq!(plain_text, "Hello 😀 world!");
        assert_eq!(paragraphs[0].text, "Hello 😀 world!");
        assert!(suggestions.is_empty(), "order {index}: {suggestions:?}");
        assert_eq!(change_decisions.len(), 2);
        assert!(change_decisions.iter().any(|record| {
            record.change_id == "suggestion-offline-bang"
                && record.decision == NativeOfficeCollaborationDocumentSuggestionDecision::Accept
                && record.text == "!"
        }));
        assert!(change_decisions.iter().any(|record| {
            record.change_id == "suggestion-offline-question"
                && record.decision == NativeOfficeCollaborationDocumentSuggestionDecision::Reject
                && record.text == "?"
        }));
        hashes.push(store.inspect().unwrap().document_state_sha256);
    }
    assert_eq!(hashes[0], hashes[1]);

    let replayed = deliver_decisions(
        &temp.path().join("decision-repeated"),
        901_220,
        &base,
        &[accept.clone(), accept.clone(), reject.clone()],
    );
    assert_eq!(replayed.inspect().unwrap().document_state_sha256, hashes[0]);

    let conflict = suggestion_store(
        &temp.path().join("decision-conflict"),
        NativeOfficeCollaborationMode::Edit,
        EDITOR_ID,
        901_221,
    );
    conflict
        .apply(suggestion_apply_request(
            "bootstrap-conflict",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            base,
        ))
        .unwrap();
    conflict
        .mutate(suggestion_request(
            "accept-before-conflict",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            decision_for(
                "suggestion-offline-bang",
                "!",
                NativeOfficeCollaborationDocumentSuggestionDecision::Accept,
            ),
        ))
        .unwrap();
    let before_conflict = conflict.inspect().unwrap();
    let rejected = conflict
        .mutate(suggestion_request(
            "reject-after-accept",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            decision_for(
                "suggestion-offline-bang",
                "!",
                NativeOfficeCollaborationDocumentSuggestionDecision::Reject,
            ),
        ))
        .unwrap_err();
    assert_eq!(
        rejected.code,
        "office.collaboration.mutation_identity_conflict"
    );
    assert_eq!(
        conflict.inspect().unwrap().document_state_sha256,
        before_conflict.document_state_sha256
    );
}

#[test]
fn reordered_replacement_half_decisions_converge_or_write_nothing() {
    let temp = tempfile::tempdir().unwrap();
    let producer = suggestion_store(
        &temp.path().join("replacement-half-producer"),
        NativeOfficeCollaborationMode::Suggest,
        SUGGESTER_ID,
        901_510,
    );
    bootstrap(
        &producer,
        NativeOfficeCollaborationMode::Suggest,
        SUGGESTER_ID,
    );
    producer
        .mutate(suggestion_request(
            "suggest-replacement",
            NativeOfficeCollaborationMode::Suggest,
            SUGGESTER_ID,
            replacement_mutation(),
        ))
        .unwrap();
    let base = producer.synchronize(None).unwrap().update;

    let accept_deletion = decision_update(
        &temp,
        901_511,
        "accept-emoji-deletion",
        &base,
        tracked_decision(
            DELETION_ID,
            NativeOfficeCollaborationDocumentSuggestionKind::Deletion,
            "😀",
            NativeOfficeCollaborationDocumentSuggestionDecision::Accept,
        ),
    );
    let reject_insertion = decision_update(
        &temp,
        901_512,
        "reject-collaborative-insertion",
        &base,
        tracked_decision(
            INSERTION_ID,
            NativeOfficeCollaborationDocumentSuggestionKind::Insertion,
            "collaborative",
            NativeOfficeCollaborationDocumentSuggestionDecision::Reject,
        ),
    );

    let mut hashes = Vec::new();
    for (index, order) in [[0_usize, 1], [1, 0]].into_iter().enumerate() {
        let updates = [accept_deletion.clone(), reject_insertion.clone()];
        let store = deliver_decisions(
            &temp.path().join(format!("replacement-half-order-{index}")),
            901_600 + index as u64,
            &base,
            &[updates[order[0]].clone(), updates[order[1]].clone()],
        );
        let projection = store.project().unwrap();
        let NativeOfficeCollaborationProjectedContent::Document {
            plain_text,
            paragraphs,
            suggestions,
            change_decisions,
            ..
        } = projection.content
        else {
            panic!("expected Document projection");
        };
        assert_eq!(plain_text, "Hello  world", "order {index}: {plain_text}");
        assert_eq!(paragraphs[0].text, "Hello  world");
        assert!(suggestions.is_empty(), "order {index}: {suggestions:?}");
        assert_eq!(change_decisions.len(), 2);
        assert!(change_decisions.iter().any(|record| {
            record.change_id == DELETION_ID
                && record.decision == NativeOfficeCollaborationDocumentSuggestionDecision::Accept
                && record.text == "😀"
        }));
        assert!(change_decisions.iter().any(|record| {
            record.change_id == INSERTION_ID
                && record.decision == NativeOfficeCollaborationDocumentSuggestionDecision::Reject
                && record.text == "collaborative"
        }));
        hashes.push(store.inspect().unwrap().document_state_sha256);
    }
    assert_eq!(hashes[0], hashes[1]);

    let replayed = deliver_decisions(
        &temp.path().join("replacement-half-repeated"),
        901_620,
        &base,
        &[
            accept_deletion.clone(),
            accept_deletion.clone(),
            reject_insertion.clone(),
        ],
    );
    assert_eq!(replayed.inspect().unwrap().document_state_sha256, hashes[0]);

    let conflict = suggestion_store(
        &temp.path().join("replacement-half-conflict"),
        NativeOfficeCollaborationMode::Edit,
        EDITOR_ID,
        901_621,
    );
    conflict
        .apply(suggestion_apply_request(
            "bootstrap-replacement-conflict",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            base,
        ))
        .unwrap();
    conflict
        .mutate(suggestion_request(
            "accept-deletion-before-conflict",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            tracked_decision(
                DELETION_ID,
                NativeOfficeCollaborationDocumentSuggestionKind::Deletion,
                "😀",
                NativeOfficeCollaborationDocumentSuggestionDecision::Accept,
            ),
        ))
        .unwrap();
    let before_conflict = conflict.inspect().unwrap();
    let rejected = conflict
        .mutate(suggestion_request(
            "reject-deletion-after-accept",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            tracked_decision(
                DELETION_ID,
                NativeOfficeCollaborationDocumentSuggestionKind::Deletion,
                "😀",
                NativeOfficeCollaborationDocumentSuggestionDecision::Reject,
            ),
        ))
        .unwrap_err();
    assert_eq!(
        rejected.code,
        "office.collaboration.mutation_identity_conflict"
    );
    assert_eq!(
        conflict.inspect().unwrap().document_state_sha256,
        before_conflict.document_state_sha256
    );
}

#[test]
fn reordered_deletion_decisions_converge_or_write_nothing() {
    let temp = tempfile::tempdir().unwrap();
    let producer = suggestion_store(
        &temp.path().join("deletion-producer"),
        NativeOfficeCollaborationMode::Suggest,
        SUGGESTER_ID,
        901_310,
    );
    bootstrap(
        &producer,
        NativeOfficeCollaborationMode::Suggest,
        SUGGESTER_ID,
    );
    producer
        .mutate(suggestion_request(
            "suggest-delete-hello",
            NativeOfficeCollaborationMode::Suggest,
            SUGGESTER_ID,
            deletion_suggestion(0, 5, "Hello", "suggestion-offline-hello"),
        ))
        .unwrap();
    producer
        .mutate(suggestion_request(
            "suggest-delete-world",
            NativeOfficeCollaborationMode::Suggest,
            SUGGESTER_ID,
            deletion_suggestion(9, 14, "world", "suggestion-offline-world"),
        ))
        .unwrap();
    let base = producer.synchronize(None).unwrap().update;

    let accept = decision_update(
        &temp,
        901_311,
        "accept-hello",
        &base,
        tracked_decision(
            "suggestion-offline-hello",
            NativeOfficeCollaborationDocumentSuggestionKind::Deletion,
            "Hello",
            NativeOfficeCollaborationDocumentSuggestionDecision::Accept,
        ),
    );
    let reject = decision_update(
        &temp,
        901_312,
        "reject-world",
        &base,
        tracked_decision(
            "suggestion-offline-world",
            NativeOfficeCollaborationDocumentSuggestionKind::Deletion,
            "world",
            NativeOfficeCollaborationDocumentSuggestionDecision::Reject,
        ),
    );

    let mut hashes = Vec::new();
    for (index, order) in [[0_usize, 1], [1, 0]].into_iter().enumerate() {
        let updates = [accept.clone(), reject.clone()];
        let store = deliver_decisions(
            &temp.path().join(format!("deletion-order-{index}")),
            901_400 + index as u64,
            &base,
            &[updates[order[0]].clone(), updates[order[1]].clone()],
        );
        let projection = store.project().unwrap();
        let NativeOfficeCollaborationProjectedContent::Document {
            plain_text,
            paragraphs,
            suggestions,
            change_decisions,
            ..
        } = projection.content
        else {
            panic!("expected Document projection");
        };
        assert_eq!(plain_text, " 😀 world");
        assert_eq!(paragraphs[0].text, " 😀 world");
        assert!(suggestions.is_empty(), "order {index}: {suggestions:?}");
        assert_eq!(change_decisions.len(), 2);
        assert!(change_decisions.iter().any(|record| {
            record.change_id == "suggestion-offline-hello"
                && record.decision == NativeOfficeCollaborationDocumentSuggestionDecision::Accept
                && record.text == "Hello"
        }));
        assert!(change_decisions.iter().any(|record| {
            record.change_id == "suggestion-offline-world"
                && record.decision == NativeOfficeCollaborationDocumentSuggestionDecision::Reject
                && record.text == "world"
        }));
        hashes.push(store.inspect().unwrap().document_state_sha256);
    }
    assert_eq!(hashes[0], hashes[1]);

    let replayed = deliver_decisions(
        &temp.path().join("deletion-repeated"),
        901_420,
        &base,
        &[accept.clone(), accept.clone(), reject.clone()],
    );
    assert_eq!(replayed.inspect().unwrap().document_state_sha256, hashes[0]);

    let conflict = suggestion_store(
        &temp.path().join("deletion-conflict"),
        NativeOfficeCollaborationMode::Edit,
        EDITOR_ID,
        901_421,
    );
    conflict
        .apply(suggestion_apply_request(
            "bootstrap-deletion-conflict",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            base,
        ))
        .unwrap();
    conflict
        .mutate(suggestion_request(
            "accept-hello-before-conflict",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            tracked_decision(
                "suggestion-offline-hello",
                NativeOfficeCollaborationDocumentSuggestionKind::Deletion,
                "Hello",
                NativeOfficeCollaborationDocumentSuggestionDecision::Accept,
            ),
        ))
        .unwrap();
    let before_conflict = conflict.inspect().unwrap();
    let rejected = conflict
        .mutate(suggestion_request(
            "reject-hello-after-accept",
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            tracked_decision(
                "suggestion-offline-hello",
                NativeOfficeCollaborationDocumentSuggestionKind::Deletion,
                "Hello",
                NativeOfficeCollaborationDocumentSuggestionDecision::Reject,
            ),
        ))
        .unwrap_err();
    assert_eq!(
        rejected.code,
        "office.collaboration.mutation_identity_conflict"
    );
    assert_eq!(
        conflict.inspect().unwrap().document_state_sha256,
        before_conflict.document_state_sha256
    );
}

fn deletion_suggestion(
    start_utf16: u32,
    end_utf16: u32,
    expected_text: &str,
    deletion_id: &str,
) -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::DocumentSuggestionCreate {
        paragraph_id: "00000001".to_owned(),
        expected_text_id: "00000002".to_owned(),
        start_utf16,
        end_utf16,
        expected_text: expected_text.to_owned(),
        replacement: String::new(),
        insertion_id: None,
        deletion_id: Some(deletion_id.to_owned()),
        author: SUGGESTER_NAME.to_owned(),
        created_at: CREATED_AT.to_owned(),
    }
}

fn tracked_decision(
    id: &str,
    kind: NativeOfficeCollaborationDocumentSuggestionKind,
    text: &str,
    decision: NativeOfficeCollaborationDocumentSuggestionDecision,
) -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::DocumentSuggestionDecide {
        suggestions: vec![suggestion_match(id, kind, text)],
        decision,
        decided_by: EDITOR_NAME.to_owned(),
        decided_at: DECIDED_AT.to_owned(),
    }
}

fn insertion_mutation(
    index_utf16: u32,
    insertion_id: &str,
    replacement: &str,
) -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::DocumentSuggestionCreate {
        paragraph_id: "00000001".to_owned(),
        expected_text_id: "00000002".to_owned(),
        start_utf16: index_utf16,
        end_utf16: index_utf16,
        expected_text: String::new(),
        replacement: replacement.to_owned(),
        insertion_id: Some(insertion_id.to_owned()),
        deletion_id: None,
        author: SUGGESTER_NAME.to_owned(),
        created_at: CREATED_AT.to_owned(),
    }
}

fn decision_for(
    id: &str,
    text: &str,
    decision: NativeOfficeCollaborationDocumentSuggestionDecision,
) -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::DocumentSuggestionDecide {
        suggestions: vec![suggestion_match(
            id,
            NativeOfficeCollaborationDocumentSuggestionKind::Insertion,
            text,
        )],
        decision,
        decided_by: EDITOR_NAME.to_owned(),
        decided_at: DECIDED_AT.to_owned(),
    }
}

fn decision_update(
    temp: &tempfile::TempDir,
    client_id: u64,
    operation_id: &str,
    base: &[u8],
    mutation: NativeOfficeCollaborationMutation,
) -> Vec<u8> {
    let store = suggestion_store(
        &temp.path().join(operation_id),
        NativeOfficeCollaborationMode::Edit,
        EDITOR_ID,
        client_id,
    );
    store
        .apply(suggestion_apply_request(
            &format!("base-{operation_id}"),
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            base.to_vec(),
        ))
        .unwrap();
    let vector = store.synchronize(None).unwrap().state_vector;
    store
        .mutate(suggestion_request(
            operation_id,
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            mutation,
        ))
        .unwrap();
    store.synchronize(Some(&vector)).unwrap().update
}

fn deliver_decisions(
    root: &std::path::Path,
    client_id: u64,
    base: &[u8],
    updates: &[Vec<u8>],
) -> NativeOfficeCollaborationStore {
    let store = suggestion_store(
        root,
        NativeOfficeCollaborationMode::Edit,
        EDITOR_ID,
        client_id,
    );
    store
        .apply(suggestion_apply_request(
            &format!("base-{client_id}"),
            NativeOfficeCollaborationMode::Edit,
            EDITOR_ID,
            base.to_vec(),
        ))
        .unwrap();
    for (index, update) in updates.iter().enumerate() {
        store
            .apply(suggestion_apply_request(
                &format!("deliver-{client_id}-{index}"),
                NativeOfficeCollaborationMode::Edit,
                EDITOR_ID,
                update.clone(),
            ))
            .unwrap();
    }
    store
}

#[allow(dead_code)]
fn apply_to_peer(update: &[u8]) {
    let peer = yrs::Doc::new();
    peer.transact_mut()
        .apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
}
