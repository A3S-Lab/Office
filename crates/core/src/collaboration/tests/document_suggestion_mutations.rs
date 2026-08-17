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
    assert_eq!(projection.version, 3);
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

#[allow(dead_code)]
fn apply_to_peer(update: &[u8]) {
    let peer = yrs::Doc::new();
    peer.transact_mut()
        .apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
}
