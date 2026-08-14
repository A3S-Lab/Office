use std::fs;

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use yrs::encoding::read::Cursor;
use yrs::sync::{Message, SyncMessage};
use yrs::updates::decoder::{Decode, DecoderV1};
use yrs::{Doc, GetString, Map, ReadTxn, StateVector, Text, Transact, Update};

use super::*;

const YJS_MARKDOWN_UPDATE_BASE64: &str = "AQey8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtbWFya2Rvd24oARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhtYXJrZG93bigBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjQyNDI0Mjpicm93c2VyLWZpeHR1cmUEARphM3Mub2ZmaWNlLm1hcmtkb3duLnNvdXJjZRUjIFNoYXJlZAoKWWpzIHRvIFlycy4A";
const YJS_MARKDOWN_STATE_VECTOR_BASE64: &str = "AbLyGRs=";
const YJS_REORDERED_FIRST_UPDATE_BASE64: &str = "AQfPuB8AKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtbWFya2Rvd24oARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhtYXJrZG93bigBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjUxNTE1MTpicm93c2VyLW9mZmxpbmUEARphM3Mub2ZmaWNlLm1hcmtkb3duLnNvdXJjZQFBAA==";
const YJS_REORDERED_SECOND_UPDATE_BASE64: &str = "AQHPuB8HhM+4HwYBQgA=";
const YJS_REORDERED_FINAL_STATE_VECTOR_BASE64: &str = "Ac+4Hwg=";

fn assert_send_sync<T: Send + Sync>() {}

#[test]
fn collaboration_store_is_send_and_sync() {
    assert_send_sync::<NativeOfficeCollaborationStore>();
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

fn agent_notes_update(client_id: u64, value: &str) -> Vec<u8> {
    let peer = Doc::with_client_id(client_id);
    let text = peer.get_or_insert_text("agent.notes");
    text.insert(&mut peer.transact_mut(), 0, value);
    let update = peer
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    update
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
