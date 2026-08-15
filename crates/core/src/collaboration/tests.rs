use std::fs;

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use yrs::encoding::read::Cursor;
use yrs::sync::{Message, SyncMessage};
use yrs::updates::decoder::{Decode, DecoderV1};
use yrs::{
    Any, Doc, GetString, Map, Out, ReadTxn, StateVector, Text, Transact, Update, Xml, XmlFragment,
    XmlOut,
};

use super::*;

mod document_structure;
mod pdf;
mod presentation;
mod spreadsheet;

const YJS_MARKDOWN_UPDATE_BASE64: &str = "AQey8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtbWFya2Rvd24oARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhtYXJrZG93bigBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjQyNDI0Mjpicm93c2VyLWZpeHR1cmUEARphM3Mub2ZmaWNlLm1hcmtkb3duLnNvdXJjZRUjIFNoYXJlZAoKWWpzIHRvIFlycy4A";
const YJS_DOCUMENT_UPDATE_BASE64: &str = "AQ+y8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtZG9jdW1lbnQoARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhkb2N1bWVudCgBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjQyNDI0Mjpicm93c2VyLWZpeHR1cmUHARthM3Mub2ZmaWNlLmRvY3VtZW50LmNvbnRlbnQDD2RvY3VtZW50U2VjdGlvbgcAsvIZBgMJcGFyYWdyYXBoBwCy8hkHBgQAsvIZCBBIZWxsbyDwn5iAIHdvcmxkKACy8hkHC3BhcmFncmFwaElkAXcIMDAwMDAwMDEoALLyGQcGdGV4dElkAXcIMDAwMDAwMDIoALLyGQYCaWQBdxJkb2N1bWVudC1zZWN0aW9uLTEoARthM3Mub2ZmaWNlLmRvY3VtZW50Lm9wdGlvbnMJcGFnZUNvbG9yAXcHI0Y4RkFGQygBG2Ezcy5vZmZpY2UuZG9jdW1lbnQub3B0aW9ucwx0cmFja0NoYW5nZXMBeAA=";
const YJS_MARKDOWN_STATE_VECTOR_BASE64: &str = "AbLyGRs=";
const YJS_REORDERED_FIRST_UPDATE_BASE64: &str = "AQfPuB8AKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtbWFya2Rvd24oARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhtYXJrZG93bigBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjUxNTE1MTpicm93c2VyLW9mZmxpbmUEARphM3Mub2ZmaWNlLm1hcmtkb3duLnNvdXJjZQFBAA==";
const YJS_REORDERED_SECOND_UPDATE_BASE64: &str = "AQHPuB8HhM+4HwYBQgA=";
const YJS_REORDERED_FINAL_STATE_VECTOR_BASE64: &str = "Ac+4Hwg=";

fn assert_send_sync<T: Send + Sync>() {}

#[test]
fn collaboration_store_is_send_and_sync() {
    assert_send_sync::<NativeOfficeCollaborationStore>();
    assert_send_sync::<NativeOfficeCollaborationTransportSession>();
}

fn create_request(root: &std::path::Path) -> NativeOfficeCollaborationCreateRequest {
    NativeOfficeCollaborationCreateRequest {
        store: root.to_path_buf(),
        artifact_id: "fixture-markdown".to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Markdown,
        actor_id: "agent-alpha".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Agent,
        mode: NativeOfficeCollaborationMode::Edit,
        operation_id: "create-1".to_owned(),
        namespace: None,
        client_id: Some(900_001),
        initial_update: None,
    }
}

fn apply_request(operation_id: &str, update: Vec<u8>) -> NativeOfficeCollaborationApplyRequest {
    NativeOfficeCollaborationApplyRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-alpha".to_owned(),
        mode: NativeOfficeCollaborationMode::Edit,
        expected_artifact_id: "fixture-markdown".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Markdown,
        update,
        if_state_vector: None,
        origin: None,
    }
}

fn checkpoint_request(operation_id: &str) -> NativeOfficeCollaborationCheckpointRequest {
    NativeOfficeCollaborationCheckpointRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-alpha".to_owned(),
        mode: NativeOfficeCollaborationMode::Edit,
        expected_artifact_id: "fixture-markdown".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Markdown,
        if_state_vector: None,
    }
}

fn mutation_request(
    operation_id: &str,
    mutation: NativeOfficeCollaborationMutation,
) -> NativeOfficeCollaborationMutationRequest {
    NativeOfficeCollaborationMutationRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-alpha".to_owned(),
        mode: NativeOfficeCollaborationMode::Edit,
        expected_artifact_id: "fixture-markdown".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Markdown,
        mutation,
        if_state_vector: None,
    }
}

fn document_create_request(root: &std::path::Path) -> NativeOfficeCollaborationCreateRequest {
    NativeOfficeCollaborationCreateRequest {
        store: root.to_path_buf(),
        artifact_id: "fixture-document".to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Document,
        actor_id: "agent-alpha".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Agent,
        mode: NativeOfficeCollaborationMode::Edit,
        operation_id: "create-document-1".to_owned(),
        namespace: None,
        client_id: Some(900_002),
        initial_update: None,
    }
}

fn document_apply_request(
    operation_id: &str,
    update: Vec<u8>,
) -> NativeOfficeCollaborationApplyRequest {
    NativeOfficeCollaborationApplyRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-alpha".to_owned(),
        mode: NativeOfficeCollaborationMode::Edit,
        expected_artifact_id: "fixture-document".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Document,
        update,
        if_state_vector: None,
        origin: None,
    }
}

fn document_mutation_request(
    operation_id: &str,
    mutation: NativeOfficeCollaborationMutation,
) -> NativeOfficeCollaborationMutationRequest {
    NativeOfficeCollaborationMutationRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-alpha".to_owned(),
        mode: NativeOfficeCollaborationMode::Edit,
        expected_artifact_id: "fixture-document".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Document,
        mutation,
        if_state_vector: None,
    }
}

fn markdown_source(store: &NativeOfficeCollaborationStore) -> String {
    let exported = store.synchronize(None).unwrap();
    let peer = Doc::with_client_id(818_181);
    peer.transact_mut()
        .apply_update(Update::decode_v1(&exported.update).unwrap())
        .unwrap();
    let transaction = peer.transact();
    transaction
        .get_text("a3s.office.markdown.source")
        .unwrap()
        .get_string(&transaction)
}

#[derive(Debug, PartialEq, Eq)]
struct DocumentParagraphState {
    paragraph_id: Option<String>,
    text_id: Option<String>,
    text: String,
}

#[derive(Debug, PartialEq, Eq)]
struct DocumentState {
    paragraphs: Vec<DocumentParagraphState>,
    page_color: Option<String>,
    track_changes: Option<bool>,
}

fn document_state(store: &NativeOfficeCollaborationStore) -> DocumentState {
    let exported = store.synchronize(None).unwrap();
    let peer = Doc::with_client_id(818_182);
    peer.transact_mut()
        .apply_update(Update::decode_v1(&exported.update).unwrap())
        .unwrap();
    let fragment = peer.get_or_insert_xml_fragment("a3s.office.document.content");
    let options = peer.get_or_insert_map("a3s.office.document.options");
    let transaction = peer.transact();
    let paragraphs = fragment
        .successors(&transaction)
        .filter_map(|node| match node {
            XmlOut::Element(element) if element.tag().as_ref() == "paragraph" => {
                let text = element
                    .children(&transaction)
                    .filter_map(|child| match child {
                        XmlOut::Text(text) => Some(text.get_string(&transaction)),
                        _ => None,
                    })
                    .collect::<String>();
                Some(DocumentParagraphState {
                    paragraph_id: xml_string_attribute(&element, &transaction, "paragraphId"),
                    text_id: xml_string_attribute(&element, &transaction, "textId"),
                    text,
                })
            }
            _ => None,
        })
        .collect::<Vec<_>>();
    let page_color = match options.get(&transaction, "pageColor") {
        Some(Out::Any(Any::String(value))) => Some(value.to_string()),
        None => None,
        value => panic!("unexpected page color: {value:?}"),
    };
    let track_changes = match options.get(&transaction, "trackChanges") {
        Some(Out::Any(Any::Bool(value))) => Some(value),
        None => None,
        value => panic!("unexpected track-changes value: {value:?}"),
    };
    DocumentState {
        paragraphs,
        page_color,
        track_changes,
    }
}

fn xml_string_attribute<T: ReadTxn>(
    element: &impl Xml,
    transaction: &T,
    name: &str,
) -> Option<String> {
    match element.get_attribute(transaction, name) {
        Some(Out::Any(Any::String(value))) => Some(value.to_string()),
        None => None,
        value => panic!("unexpected XML attribute '{name}': {value:?}"),
    }
}

fn document_paragraph(paragraph_id: &str, text_id: &str, text: &str) -> DocumentParagraphState {
    DocumentParagraphState {
        paragraph_id: Some(paragraph_id.to_owned()),
        text_id: Some(text_id.to_owned()),
        text: text.to_owned(),
    }
}

fn agent_notes_update(client_id: u64, value: &str) -> Vec<u8> {
    let peer = Doc::with_client_id(client_id);
    let text = peer.get_or_insert_text("agent.notes");
    text.insert(&mut peer.transact_mut(), 0, value);
    let update = peer
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    update
}

fn browser_transport_message(
    message_type: NativeOfficeCollaborationTransportMessageType,
    payload: Vec<u8>,
) -> NativeOfficeCollaborationTransportMessage {
    NativeOfficeCollaborationTransportMessage {
        protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
        version: NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
        artifact_id: "fixture-markdown".to_owned(),
        artifact_kind: NativeOfficeCollaborationArtifactKind::Markdown,
        namespace: NATIVE_OFFICE_COLLABORATION_NAMESPACE.to_owned(),
        sender_client_id: 424_242,
        message_type,
        payload,
        origin: None,
    }
}

#[test]
fn yjs_fixture_round_trips_through_yrs_state_vectors() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    let update = STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap();
    let applied = store
        .apply(apply_request("sync-browser-1", update))
        .unwrap();
    assert!(applied.state_changed);
    assert!(!applied.duplicate);

    let inspection = store.inspect().unwrap();
    assert_eq!(
        inspection.metadata,
        Some(NativeOfficeCollaborationMetadata {
            protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
            version: 1,
            artifact_id: "fixture-markdown".to_owned(),
            kind: NativeOfficeCollaborationArtifactKind::Markdown,
            initialized: true,
        })
    );
    assert_eq!(inspection.bootstrap_initializer_count, Some(1));
    assert_eq!(inspection.bootstrap_valid, Some(true));
    assert_eq!(
        inspection.state_vector,
        STANDARD.decode(YJS_MARKDOWN_STATE_VECTOR_BASE64).unwrap()
    );

    let exported = store.synchronize(None).unwrap();
    let browser_peer = Doc::with_client_id(123_456);
    browser_peer
        .transact_mut()
        .apply_update(Update::decode_v1(&exported.update).unwrap())
        .unwrap();
    let source = browser_peer
        .transact()
        .get_text("a3s.office.markdown.source")
        .unwrap();
    assert_eq!(
        source.get_string(&browser_peer.transact()),
        "# Shared\n\nYjs to Yrs."
    );
}

#[test]
fn typed_markdown_mutations_are_utf16_safe_durable_and_idempotent() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    store
        .apply(apply_request(
            "bootstrap-browser",
            STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();

    let replaced = store
        .mutate(mutation_request(
            "typed-replace-1",
            NativeOfficeCollaborationMutation::MarkdownReplace {
                markdown: "A😀B".to_owned(),
            },
        ))
        .unwrap();
    assert!(replaced.state_changed);
    assert_eq!(replaced.sequence, Some(2));
    assert_eq!(markdown_source(&store), "A😀B");

    let spliced_request = mutation_request(
        "typed-splice-1",
        NativeOfficeCollaborationMutation::MarkdownSplice {
            index_utf16: 1,
            delete_utf16: 2,
            insert: "🦀".to_owned(),
        },
    );
    let spliced = store.mutate(spliced_request.clone()).unwrap();
    assert!(spliced.state_changed);
    assert_eq!(spliced.sequence, Some(3));
    assert_eq!(markdown_source(&store), "A🦀B");

    let replay = store.mutate(spliced_request).unwrap();
    assert!(replay.duplicate);
    assert_eq!(replay.sequence, Some(3));

    let before_invalid = store.inspect().unwrap();
    let invalid = store
        .mutate(mutation_request(
            "typed-splice-invalid",
            NativeOfficeCollaborationMutation::MarkdownSplice {
                index_utf16: 2,
                delete_utf16: 0,
                insert: "!".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(invalid.code, "office.collaboration.mutation_range_invalid");
    let after_invalid = store.inspect().unwrap();
    assert_eq!(
        after_invalid.current_sequence,
        before_invalid.current_sequence
    );
    assert_eq!(
        after_invalid.document_state_sha256,
        before_invalid.document_state_sha256
    );

    drop(store);
    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    assert_eq!(markdown_source(&reopened), "A🦀B");
    let events = reopened
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(1),
            limit: 2,
        })
        .unwrap();
    assert_eq!(events.updates.len(), 2);
    assert!(events
        .updates
        .iter()
        .all(|event| event.operation_kind == NativeOfficeCollaborationOperationKind::Mutate));
}

#[test]
fn typed_document_mutations_converge_with_browser_xml_and_sidecars() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("document-replica");
    let store = NativeOfficeCollaborationStore::create(document_create_request(&root)).unwrap();
    store
        .apply(document_apply_request(
            "bootstrap-browser-document",
            STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    assert_eq!(
        document_state(&store),
        DocumentState {
            paragraphs: vec![document_paragraph("00000001", "00000002", "Hello 😀 world",)],
            page_color: Some("#F8FAFC".to_owned()),
            track_changes: Some(true),
        }
    );

    let replace_request = document_mutation_request(
        "document-replace-1",
        NativeOfficeCollaborationMutation::DocumentReplaceText {
            search: "😀".to_owned(),
            replacement: "🦀".to_owned(),
            expected_matches: 1,
        },
    );
    let replaced = store.mutate(replace_request.clone()).unwrap();
    assert!(replaced.state_changed);
    assert_eq!(replaced.sequence, Some(2));

    let page_color = store
        .mutate(document_mutation_request(
            "document-page-color-1",
            NativeOfficeCollaborationMutation::DocumentSetPageColor {
                page_color: "#101828".to_owned(),
            },
        ))
        .unwrap();
    assert!(page_color.state_changed);
    assert_eq!(page_color.sequence, Some(3));

    let track_changes = store
        .mutate(document_mutation_request(
            "document-track-changes-1",
            NativeOfficeCollaborationMutation::DocumentClearTrackChanges {},
        ))
        .unwrap();
    assert!(track_changes.state_changed);
    assert_eq!(track_changes.sequence, Some(4));

    let temporary = store
        .mutate(document_mutation_request(
            "document-insert-temporary-1",
            NativeOfficeCollaborationMutation::DocumentInsertParagraph {
                anchor_paragraph_id: "00000001".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                paragraph_id: "00000010".to_owned(),
                text_id: "00000011".to_owned(),
                text: "Temporary".to_owned(),
            },
        ))
        .unwrap();
    assert!(temporary.state_changed);
    assert_eq!(temporary.sequence, Some(5));

    let inserted = store
        .mutate(document_mutation_request(
            "document-insert-final-1",
            NativeOfficeCollaborationMutation::DocumentInsertParagraph {
                anchor_paragraph_id: "00000010".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                paragraph_id: "00000012".to_owned(),
                text_id: "00000013".to_owned(),
                text: "Native paragraph".to_owned(),
            },
        ))
        .unwrap();
    assert!(inserted.state_changed);
    assert_eq!(inserted.sequence, Some(6));

    let deleted = store
        .mutate(document_mutation_request(
            "document-delete-temporary-1",
            NativeOfficeCollaborationMutation::DocumentDeleteParagraph {
                paragraph_id: "00000010".to_owned(),
                expected_text_id: "00000011".to_owned(),
                expected_text: "Temporary".to_owned(),
            },
        ))
        .unwrap();
    assert!(deleted.state_changed);
    assert_eq!(deleted.sequence, Some(7));
    assert_eq!(
        document_state(&store),
        DocumentState {
            paragraphs: vec![
                document_paragraph("00000001", "00000003", "Hello 🦀 world"),
                document_paragraph("00000012", "00000013", "Native paragraph"),
            ],
            page_color: Some("#101828".to_owned()),
            track_changes: None,
        }
    );

    let replay = store.mutate(replace_request).unwrap();
    assert!(replay.duplicate);
    assert_eq!(replay.sequence, Some(2));

    let before_conflict = store.inspect().unwrap();
    let delete_conflict = store
        .mutate(document_mutation_request(
            "document-delete-conflict",
            NativeOfficeCollaborationMutation::DocumentDeleteParagraph {
                paragraph_id: "00000012".to_owned(),
                expected_text_id: "00000014".to_owned(),
                expected_text: "Native paragraph".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        delete_conflict.code,
        "office.collaboration.mutation_match_conflict"
    );
    let conflict = store
        .mutate(document_mutation_request(
            "document-replace-conflict",
            NativeOfficeCollaborationMutation::DocumentReplaceText {
                search: "world".to_owned(),
                replacement: "document".to_owned(),
                expected_matches: 2,
            },
        ))
        .unwrap_err();
    assert_eq!(
        conflict.code,
        "office.collaboration.mutation_match_conflict"
    );
    let after_conflict = store.inspect().unwrap();
    assert_eq!(
        after_conflict.current_sequence,
        before_conflict.current_sequence
    );
    assert_eq!(
        after_conflict.document_state_sha256,
        before_conflict.document_state_sha256
    );

    drop(store);
    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    assert_eq!(
        document_state(&reopened),
        DocumentState {
            paragraphs: vec![
                document_paragraph("00000001", "00000003", "Hello 🦀 world"),
                document_paragraph("00000012", "00000013", "Native paragraph"),
            ],
            page_color: Some("#101828".to_owned()),
            track_changes: None,
        }
    );
    let events = reopened
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(1),
            limit: 6,
        })
        .unwrap();
    assert_eq!(events.updates.len(), 6);
    assert!(events
        .updates
        .iter()
        .all(|event| event.operation_kind == NativeOfficeCollaborationOperationKind::Mutate));
}

#[test]
fn typed_mutations_require_initialization_and_edit_mode_but_raw_sync_does_not() {
    let temp = tempfile::tempdir().unwrap();
    let uninitialized_root = temp.path().join("uninitialized");
    let uninitialized =
        NativeOfficeCollaborationStore::create(create_request(&uninitialized_root)).unwrap();
    let error = uninitialized
        .mutate(mutation_request(
            "typed-before-bootstrap",
            NativeOfficeCollaborationMutation::MarkdownReplace {
                markdown: "blocked".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.mutation_uninitialized");

    for (index, mode) in [
        NativeOfficeCollaborationMode::View,
        NativeOfficeCollaborationMode::Comment,
        NativeOfficeCollaborationMode::Suggest,
    ]
    .into_iter()
    .enumerate()
    {
        let root = temp.path().join(format!("read-only-{index}"));
        let mut create = create_request(&root);
        create.mode = mode;
        create.client_id = Some(910_000 + index as u64);
        create.operation_id = format!("create-read-only-{index}");
        let store = NativeOfficeCollaborationStore::create(create).unwrap();

        let mut raw = apply_request(
            &format!("bootstrap-read-only-{index}"),
            STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap(),
        );
        raw.mode = mode;
        store.apply(raw).unwrap();
        assert_eq!(markdown_source(&store), "# Shared\n\nYjs to Yrs.");

        let mut typed = mutation_request(
            &format!("typed-read-only-{index}"),
            NativeOfficeCollaborationMutation::MarkdownReplace {
                markdown: "blocked".to_owned(),
            },
        );
        typed.mode = mode;
        let before = store.inspect().unwrap();
        let error = store.mutate(typed).unwrap_err();
        assert_eq!(error.code, "office.collaboration.mutation_forbidden");
        let after = store.inspect().unwrap();
        assert_eq!(after.current_sequence, before.current_sequence);
        assert_eq!(after.document_state_sha256, before.document_state_sha256);
    }
}

#[test]
fn standard_y_sync_messages_carry_the_replica_vector_and_update() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    let update = STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap();
    store
        .apply(apply_request("sync-browser-1", update.clone()))
        .unwrap();

    let step1 = store.sync_step1().unwrap();
    let message = Message::decode(&mut DecoderV1::new(Cursor::new(&step1.message))).unwrap();
    let Message::Sync(SyncMessage::SyncStep1(vector)) = message else {
        panic!("expected SyncStep1")
    };
    assert_eq!(canonical_state_vector(&vector), step1.state_vector);

    let encoded_update = NativeOfficeCollaborationStore::encode_sync_update(&update).unwrap();
    let message = Message::decode(&mut DecoderV1::new(Cursor::new(&encoded_update))).unwrap();
    assert_eq!(message, Message::Sync(SyncMessage::Update(update)));

    let remote = StateVector::default();
    let mut encoder = yrs::updates::encoder::EncoderV1::new();
    yrs::updates::encoder::Encode::encode(
        &Message::Sync(SyncMessage::SyncStep1(remote)),
        &mut encoder,
    );
    let request = yrs::updates::encoder::Encoder::to_vec(encoder);
    let handled = store.handle_sync_message(&request, None).unwrap();
    assert_eq!(
        handled.kind,
        NativeOfficeCollaborationSyncMessageKind::SyncStep1
    );
    let response = Message::decode(&mut DecoderV1::new(Cursor::new(
        handled.response.as_deref().unwrap(),
    )))
    .unwrap();
    assert!(matches!(response, Message::Sync(SyncMessage::SyncStep2(_))));

    let mut packet = request.clone();
    packet.extend_from_slice(&request);
    let error = store.handle_sync_message(&packet, None).unwrap_err();
    assert_eq!(error.code, "office.collaboration.sync_message_invalid");
}

#[test]
fn y_sync_update_handler_requires_identity_and_persists_idempotently() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    let update = STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap();
    let message = NativeOfficeCollaborationStore::encode_sync_update(&update).unwrap();

    let missing = store.handle_sync_message(&message, None).unwrap_err();
    assert_eq!(missing.code, "office.collaboration.sync_identity_required");
    let mut mutation = apply_request("sync-message-1", Vec::new());
    mutation.if_state_vector = Some(canonical_state_vector(&StateVector::default()));
    let first = store
        .handle_sync_message(&message, Some(mutation.clone()))
        .unwrap();
    assert!(first.apply.as_ref().unwrap().state_changed);
    let replay = store.handle_sync_message(&message, Some(mutation)).unwrap();
    assert!(replay.apply.unwrap().duplicate);
}

#[test]
fn live_transport_session_bridges_browser_envelopes_and_suppresses_echoes() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    let mut session = NativeOfficeCollaborationTransportSession::attach(store.clone()).unwrap();
    let mut observer = NativeOfficeCollaborationTransportSession::attach(store.clone()).unwrap();

    let initial = session.synchronize().unwrap();
    assert_eq!(
        initial.message_type,
        NativeOfficeCollaborationTransportMessageType::SyncStep1
    );
    assert_eq!(initial.sender_client_id, 900_001);

    let browser_update = STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap();
    let mut incoming = browser_transport_message(
        NativeOfficeCollaborationTransportMessageType::Update,
        browser_update.clone(),
    );
    incoming.origin = Some(NativeOfficeCollaborationOrigin {
        protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
        kind: NativeOfficeCollaborationOriginKind::Editor,
        actor_id: Some("browser-user".to_owned()),
        operation_id: Some("browser-edit-1".to_owned()),
    });
    let received = session
        .receive(NativeOfficeCollaborationTransportReceiveRequest {
            message: incoming.clone(),
            operation_id: Some("delivery-browser-edit-1".to_owned()),
            if_state_vector: None,
        })
        .unwrap();
    assert!(received.apply.as_ref().unwrap().state_changed);
    assert_eq!(received.apply.as_ref().unwrap().sequence, Some(1));
    let durable = store
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(0),
            limit: 1,
        })
        .unwrap();
    assert_eq!(durable.updates[0].origin, incoming.origin);

    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    let restarted = reopened
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(0),
            limit: 1,
        })
        .unwrap();
    assert_eq!(restarted.updates[0].origin, incoming.origin);
    let forwarded = observer
        .poll(MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH)
        .unwrap();
    assert_eq!(forwarded.messages.len(), 1);
    assert_eq!(forwarded.messages[0].origin, incoming.origin);
    assert_eq!(forwarded.messages[0].payload, browser_update);

    let replay = session
        .receive(NativeOfficeCollaborationTransportReceiveRequest {
            message: incoming.clone(),
            operation_id: Some("delivery-browser-edit-1".to_owned()),
            if_state_vector: None,
        })
        .unwrap();
    assert!(replay.apply.unwrap().duplicate);

    let mut conflicting_origin = incoming.clone();
    conflicting_origin.origin.as_mut().unwrap().operation_id =
        Some("browser-edit-conflict".to_owned());
    let conflict = session
        .receive(NativeOfficeCollaborationTransportReceiveRequest {
            message: conflicting_origin,
            operation_id: Some("delivery-browser-edit-1".to_owned()),
            if_state_vector: None,
        })
        .unwrap_err();
    assert_eq!(conflict.code, "office.collaboration.operation_conflict");
    let echoed = session
        .poll(MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH)
        .unwrap();
    assert!(echoed.messages.is_empty());
    assert_eq!(echoed.cursor_sequence, 1);

    let local_update = agent_notes_update(700_007, "live native edit");
    store
        .apply(apply_request("native-live-1", local_update.clone()))
        .unwrap();
    let outbound = session
        .poll(MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH)
        .unwrap();
    assert_eq!(outbound.messages.len(), 1);
    let message = &outbound.messages[0];
    assert_eq!(
        message.message_type,
        NativeOfficeCollaborationTransportMessageType::Update
    );
    assert_eq!(message.payload, local_update);
    assert_eq!(
        message.origin,
        Some(NativeOfficeCollaborationOrigin {
            protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
            kind: NativeOfficeCollaborationOriginKind::Agent,
            actor_id: Some("agent-alpha".to_owned()),
            operation_id: Some("native-live-1".to_owned()),
        })
    );

    let browser = Doc::with_client_id(424_242);
    browser
        .transact_mut()
        .apply_update(Update::decode_v1(&browser_update).unwrap())
        .unwrap();
    let browser_vector = canonical_state_vector(&browser.transact().state_vector());
    let response = session
        .receive(NativeOfficeCollaborationTransportReceiveRequest {
            message: browser_transport_message(
                NativeOfficeCollaborationTransportMessageType::SyncStep1,
                browser_vector,
            ),
            operation_id: None,
            if_state_vector: None,
        })
        .unwrap()
        .response
        .unwrap();
    assert_eq!(
        response.message_type,
        NativeOfficeCollaborationTransportMessageType::SyncStep2
    );
    browser
        .transact_mut()
        .apply_update(Update::decode_v1(&response.payload).unwrap())
        .unwrap();
    let transaction = browser.transact();
    assert_eq!(
        transaction
            .get_text("agent.notes")
            .unwrap()
            .get_string(&transaction),
        "live native edit"
    );
}

#[test]
fn live_transport_session_repairs_compacted_history_bidirectionally() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    let mut session = NativeOfficeCollaborationTransportSession::attach(store.clone()).unwrap();
    let update = agent_notes_update(700_008, "compacted native edit");
    store
        .apply(apply_request("native-before-compaction", update))
        .unwrap();
    store
        .checkpoint(checkpoint_request("compact-live-history"))
        .unwrap();

    let recovery = session
        .poll(MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH)
        .unwrap();
    assert!(recovery.resynchronized);
    assert_eq!(recovery.cursor_sequence, 1);
    assert_eq!(recovery.messages.len(), 2);

    let full_state = &recovery.messages[0];
    assert_eq!(
        full_state.message_type,
        NativeOfficeCollaborationTransportMessageType::Update
    );
    assert_eq!(
        full_state.origin,
        Some(NativeOfficeCollaborationOrigin {
            protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
            kind: NativeOfficeCollaborationOriginKind::System,
            actor_id: Some("agent-alpha".to_owned()),
            operation_id: None,
        })
    );
    let browser = Doc::with_client_id(424_243);
    browser
        .transact_mut()
        .apply_update(Update::decode_v1(&full_state.payload).unwrap())
        .unwrap();
    let transaction = browser.transact();
    assert_eq!(
        transaction
            .get_text("agent.notes")
            .unwrap()
            .get_string(&transaction),
        "compacted native edit"
    );
    let browser_vector = canonical_state_vector(&transaction.state_vector());
    drop(transaction);

    let handshake = &recovery.messages[1];
    assert_eq!(
        handshake.message_type,
        NativeOfficeCollaborationTransportMessageType::SyncStep1
    );
    assert_eq!(handshake.payload, browser_vector);
    assert!(handshake.origin.is_none());
}

#[test]
fn live_transport_session_fails_closed_across_identity_and_message_boundaries() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    let mut session = NativeOfficeCollaborationTransportSession::attach(store).unwrap();

    let mut wrong_artifact = browser_transport_message(
        NativeOfficeCollaborationTransportMessageType::SyncStep1,
        canonical_state_vector(&StateVector::default()),
    );
    wrong_artifact.artifact_id = "another-artifact".to_owned();
    let error = session
        .receive(NativeOfficeCollaborationTransportReceiveRequest {
            message: wrong_artifact,
            operation_id: None,
            if_state_vector: None,
        })
        .unwrap_err();
    assert_eq!(
        error.code,
        "office.collaboration.transport_identity_mismatch"
    );

    let update = agent_notes_update(777_777, "remote");
    let missing_operation = session
        .receive(NativeOfficeCollaborationTransportReceiveRequest {
            message: browser_transport_message(
                NativeOfficeCollaborationTransportMessageType::Update,
                update.clone(),
            ),
            operation_id: None,
            if_state_vector: None,
        })
        .unwrap_err();
    assert_eq!(
        missing_operation.code,
        "office.collaboration.sync_identity_required"
    );

    let mut own_message = browser_transport_message(
        NativeOfficeCollaborationTransportMessageType::Update,
        update,
    );
    own_message.sender_client_id = 900_001;
    let ignored = session
        .receive(NativeOfficeCollaborationTransportReceiveRequest {
            message: own_message,
            operation_id: Some("ignored-own-message".to_owned()),
            if_state_vector: None,
        })
        .unwrap();
    assert!(ignored.ignored);
    assert!(ignored.apply.is_none());

    let mut invalid_origin = browser_transport_message(
        NativeOfficeCollaborationTransportMessageType::SyncStep1,
        canonical_state_vector(&StateVector::default()),
    );
    invalid_origin.origin = Some(NativeOfficeCollaborationOrigin {
        protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
        kind: NativeOfficeCollaborationOriginKind::System,
        actor_id: None,
        operation_id: None,
    });
    let error = session
        .receive(NativeOfficeCollaborationTransportReceiveRequest {
            message: invalid_origin,
            operation_id: None,
            if_state_vector: None,
        })
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.transport_message_invalid");
}

#[test]
fn reordered_yjs_updates_survive_restart_and_converge() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    let second = STANDARD.decode(YJS_REORDERED_SECOND_UPDATE_BASE64).unwrap();
    let first = STANDARD.decode(YJS_REORDERED_FIRST_UPDATE_BASE64).unwrap();

    let pending = store.apply(apply_request("offline-2", second)).unwrap();
    assert!(pending.state_changed);
    assert!(store.inspect().unwrap().pending_updates);
    drop(store);

    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    assert!(reopened.inspect().unwrap().pending_updates);
    reopened.apply(apply_request("offline-1", first)).unwrap();
    let inspection = reopened.inspect().unwrap();
    assert!(!inspection.pending_updates);
    assert_eq!(
        inspection.state_vector,
        STANDARD
            .decode(YJS_REORDERED_FINAL_STATE_VECTOR_BASE64)
            .unwrap()
    );

    let exported = reopened.synchronize(None).unwrap();
    let peer = Doc::with_client_id(81);
    peer.transact_mut()
        .apply_update(Update::decode_v1(&exported.update).unwrap())
        .unwrap();
    let transaction = peer.transact();
    let source = transaction.get_text("a3s.office.markdown.source").unwrap();
    assert_eq!(source.get_string(&transaction), "AB");
}

#[test]
fn operation_replays_are_idempotent_and_conflicts_fail_closed() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    let update = STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap();
    let first = store
        .apply(apply_request("stable-operation", update.clone()))
        .unwrap();
    let replay = store
        .apply(apply_request("stable-operation", update))
        .unwrap();
    assert!(!first.duplicate);
    assert!(replay.duplicate);
    assert_eq!(first.sequence, replay.sequence);

    let conflict = store
        .apply(apply_request("stable-operation", vec![0, 0]))
        .unwrap_err();
    assert_eq!(conflict.code, "office.collaboration.operation_conflict");
}

#[test]
fn stale_preconditions_and_identity_mismatches_do_not_append_updates() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    let empty_vector = canonical_state_vector(&StateVector::default());
    store
        .apply(apply_request(
            "sync-browser-1",
            STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();

    let mut request = apply_request("sync-stale", vec![0, 0]);
    request.if_state_vector = Some(empty_vector);
    let stale = store.apply(request).unwrap_err();
    assert_eq!(stale.code, "office.collaboration.stale_state");

    let mut mismatch = apply_request("sync-wrong-artifact", vec![0, 0]);
    mismatch.expected_artifact_id = "another-artifact".to_owned();
    let mismatch = store.apply(mismatch).unwrap_err();
    assert_eq!(mismatch.code, "office.collaboration.artifact_mismatch");
    assert_eq!(store.inspect().unwrap().update_count, 1);
}

#[test]
fn checkpoint_compacts_updates_but_preserves_operation_receipts() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    store
        .apply(apply_request(
            "sync-browser-1",
            STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    assert_eq!(store.inspect().unwrap().update_count, 1);

    let checkpoint = store
        .checkpoint(checkpoint_request("checkpoint-1"))
        .unwrap();
    assert_eq!(checkpoint.compacted_updates, 1);
    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    let inspection = reopened.inspect().unwrap();
    assert_eq!(inspection.update_count, 0);
    assert_eq!(inspection.checkpoint_sequence, 1);
    assert_eq!(inspection.current_sequence, 1);
    assert_eq!(inspection.operation_count, 3);
    let replay = reopened
        .apply(apply_request(
            "sync-browser-1",
            STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    assert!(replay.duplicate);
}

#[test]
fn durable_events_are_bounded_resumable_and_auditable() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();

    let initial = store
        .events(NativeOfficeCollaborationEventsRequest::default())
        .unwrap();
    assert_eq!(initial.starting_sequence, 0);
    assert_eq!(initial.cursor_sequence, 0);
    assert!(initial.updates.is_empty());

    let browser_update = STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap();
    store
        .apply(apply_request("sync-browser-1", browser_update.clone()))
        .unwrap();
    let agent_update = agent_notes_update(777_777, "next");
    store
        .apply(apply_request("sync-agent-2", agent_update.clone()))
        .unwrap();

    let first = store
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(0),
            limit: 1,
        })
        .unwrap();
    assert_eq!(first.current_sequence, 2);
    assert_eq!(first.cursor_sequence, 1);
    assert!(first.has_more);
    assert!(first.reset.is_none());
    assert_eq!(first.updates.len(), 1);
    let event = &first.updates[0];
    assert_eq!(event.sequence, 1);
    assert_eq!(event.operation_id, "sync-browser-1");
    assert_eq!(
        event.operation_kind,
        NativeOfficeCollaborationOperationKind::Synchronize
    );
    assert_eq!(event.actor_id, "agent-alpha");
    assert_eq!(event.actor_kind, NativeOfficeCollaborationActorKind::Agent);
    assert_eq!(event.update, browser_update);
    assert_eq!(event.update_bytes, event.update.len() as u64);
    assert_eq!(event.update_sha256, sha256_hex(&event.update));

    let second = store
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(first.cursor_sequence),
            limit: 1,
        })
        .unwrap();
    assert_eq!(second.cursor_sequence, 2);
    assert!(!second.has_more);
    assert_eq!(second.updates.len(), 1);
    assert_eq!(second.updates[0].operation_id, "sync-agent-2");
    assert_eq!(second.updates[0].update, agent_update);

    let invalid_limit = store
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(2),
            limit: 0,
        })
        .unwrap_err();
    assert_eq!(
        invalid_limit.code,
        "office.collaboration.event_limit_invalid"
    );
    let ahead = store
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(3),
            limit: 1,
        })
        .unwrap_err();
    assert_eq!(ahead.code, "office.collaboration.sequence_ahead");
}

#[test]
fn compacted_event_history_returns_a_complete_reset_update() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    store
        .apply(apply_request(
            "sync-browser-1",
            STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    store
        .checkpoint(checkpoint_request("checkpoint-1"))
        .unwrap();

    let batch = store
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(0),
            limit: 1,
        })
        .unwrap();
    assert_eq!(batch.starting_sequence, 0);
    assert_eq!(batch.checkpoint_sequence, 1);
    assert_eq!(batch.current_sequence, 1);
    assert_eq!(batch.cursor_sequence, 1);
    assert!(!batch.has_more);
    assert!(batch.updates.is_empty());
    let reset = batch.reset.unwrap();
    assert_eq!(reset.sequence, 1);
    assert_eq!(reset.update_bytes, reset.update.len() as u64);
    assert_eq!(reset.update_sha256, sha256_hex(&reset.update));

    let peer = Doc::with_client_id(81);
    peer.transact_mut()
        .apply_update(Update::decode_v1(&reset.update).unwrap())
        .unwrap();
    let transaction = peer.transact();
    let source = transaction.get_text("a3s.office.markdown.source").unwrap();
    assert_eq!(source.get_string(&transaction), "# Shared\n\nYjs to Yrs.");
}

#[test]
fn interrupted_checkpoint_files_are_ignored_until_the_commit_marker_exists() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    let before = store.inspect().unwrap();
    let interrupted =
        root.join("checkpoints")
            .join(format!("{:020}-{}.update", 99, "0".repeat(64)));
    fs::write(interrupted, b"interrupted checkpoint bytes").unwrap();

    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    let after = reopened.inspect().unwrap();
    assert_eq!(after.checkpoint_sequence, before.checkpoint_sequence);
    assert_eq!(after.document_state_sha256, before.document_state_sha256);
}

#[test]
fn incomplete_log_is_reported_instead_of_silently_skipped() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    store
        .apply(apply_request(
            "sync-browser-1",
            STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();

    let second_peer = Doc::with_client_id(777_777);
    let text = second_peer.get_or_insert_text("agent.notes");
    text.insert(&mut second_peer.transact_mut(), 0, "next");
    let second_update = second_peer
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    store
        .apply(apply_request("sync-agent-2", second_update))
        .unwrap();
    let first_entry = fs::read_dir(root.join("updates"))
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .find(|path| {
            path.file_name()
                .and_then(|value| value.to_str())
                .is_some_and(|value| value.starts_with("00000000000000000001-"))
        })
        .unwrap();
    fs::remove_file(first_entry).unwrap();

    let error = store.inspect().unwrap_err();
    assert_eq!(error.code, "office.collaboration.log_incomplete");
}

#[test]
fn initialized_metadata_with_no_initializer_is_rejected() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    let peer = Doc::with_client_id(51);
    let metadata = peer.get_or_insert_map("a3s.office.metadata");
    let mut transaction = peer.transact_mut();
    metadata.insert(
        &mut transaction,
        "protocol",
        NATIVE_OFFICE_COLLABORATION_PROTOCOL,
    );
    metadata.insert(&mut transaction, "version", 1_i64);
    metadata.insert(&mut transaction, "artifactId", "fixture-markdown");
    metadata.insert(&mut transaction, "kind", "markdown");
    metadata.insert(&mut transaction, "initialized", true);
    drop(transaction);
    let update = peer
        .transact()
        .encode_state_as_update_v1(&StateVector::default());

    let error = store
        .apply(apply_request("invalid-bootstrap", update))
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.bootstrap_invalid");
    assert_eq!(store.inspect().unwrap().update_count, 0);
}

#[test]
fn create_refuses_to_replace_an_existing_directory() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    fs::create_dir(&root).unwrap();
    let error = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap_err();
    assert_eq!(error.code, "office.collaboration.store_exists");
}

#[test]
fn create_replay_is_idempotent_but_an_unrelated_existing_store_is_rejected() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("replica");
    let request = create_request(&root);
    let first = NativeOfficeCollaborationStore::create(request.clone()).unwrap();
    let replay = NativeOfficeCollaborationStore::create(request).unwrap();
    assert_eq!(first.path(), replay.path());

    let mut conflict = create_request(&root);
    conflict.operation_id = "create-2".to_owned();
    let error = NativeOfficeCollaborationStore::create(conflict).unwrap_err();
    assert_eq!(error.code, "office.collaboration.store_exists");
}
