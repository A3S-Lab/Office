use serde_json::json;
use yrs::sync::{Awareness, AwarenessUpdate};
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{ClientID, Doc};

use super::*;

fn manifest(kind: NativeOfficeCollaborationArtifactKind) -> NativeOfficeCollaborationManifest {
    NativeOfficeCollaborationManifest {
        format: NATIVE_OFFICE_COLLABORATION_STORE_FORMAT.to_owned(),
        schema_version: NATIVE_OFFICE_COLLABORATION_STORE_SCHEMA_VERSION,
        protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
        protocol_version: NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
        namespace: NATIVE_OFFICE_COLLABORATION_NAMESPACE.to_owned(),
        artifact_id: format!("presence-{}", kind.as_str()),
        kind,
        actor_id: "native-agent".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Agent,
        mode: NativeOfficeCollaborationMode::Edit,
        client_id: 900_101,
    }
}

fn profile() -> NativeOfficeCollaborationPresenceProfile {
    NativeOfficeCollaborationPresenceProfile {
        name: "A3S Agent".to_owned(),
        color: Some("#2563eb".to_owned()),
        avatar_url: Some("https://example.test/agent.png".to_owned()),
    }
}

fn remote_message(
    manifest: &NativeOfficeCollaborationManifest,
    sender_client_id: u64,
    clock: u32,
    state: serde_json::Value,
) -> NativeOfficeCollaborationAwarenessMessage {
    let mut awareness = Awareness::new(Doc::with_client_id(sender_client_id));
    for _ in 0..clock {
        awareness.set_local_state_raw(state.to_string());
    }
    NativeOfficeCollaborationAwarenessMessage {
        protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
        version: NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
        artifact_id: manifest.artifact_id.clone(),
        artifact_kind: manifest.kind,
        namespace: manifest.namespace.clone(),
        sender_client_id,
        payload: awareness.update().unwrap().encode_v1(),
    }
}

fn remote_state(
    manifest: &NativeOfficeCollaborationManifest,
    activity: &str,
    location: serde_json::Value,
) -> serde_json::Value {
    json!({
        "a3sOffice": {
            "protocol": NATIVE_OFFICE_COLLABORATION_PROTOCOL,
            "version": NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
            "artifactId": manifest.artifact_id,
            "artifactKind": manifest.kind,
            "namespace": manifest.namespace,
            "presenceId": "remote-browser:1",
            "actor": {
                "id": "browser-user",
                "name": "Browser User",
                "kind": "human"
            },
            "mode": "comment",
            "activity": activity,
            "location": location
        }
    })
}

#[test]
fn native_presence_publishes_browser_compatible_awareness_without_persistence() {
    let temp = tempfile::tempdir().unwrap();
    let store = NativeOfficeCollaborationStore::create(NativeOfficeCollaborationCreateRequest {
        store: temp.path().join("agent.replica"),
        artifact_id: "presence-markdown".to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Markdown,
        actor_id: "native-agent".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Agent,
        mode: NativeOfficeCollaborationMode::Edit,
        operation_id: "create-presence-fixture".to_owned(),
        namespace: None,
        client_id: Some(900_101),
        initial_update: None,
    })
    .unwrap();
    let before = store.inspect().unwrap();
    let presence_client_id = 900_102;
    let mut presence = NativeOfficeCollaborationPresenceSession::new_with_sender_client_id(
        before.manifest.clone(),
        presence_client_id,
        profile(),
    )
    .unwrap();

    let initial = presence.local_message().unwrap();
    assert_eq!(initial.sender_client_id, presence_client_id);
    assert_ne!(initial.sender_client_id, before.manifest.client_id);
    let decoded = AwarenessUpdate::decode_v1(&initial.payload).unwrap();
    assert_eq!(decoded.clients.len(), 1);
    let entry = decoded
        .clients
        .get(&ClientID::new(presence_client_id))
        .unwrap();
    let value: serde_json::Value = serde_json::from_str(&entry.json).unwrap();
    assert_eq!(value["a3sOffice"]["actor"]["name"], "A3S Agent");
    assert_eq!(value["a3sOffice"]["activity"], "active");
    assert!(value["a3sOffice"].get("location").is_none());

    let updated = presence
        .update(NativeOfficeCollaborationPresenceUpdate {
            activity: NativeOfficeCollaborationPresenceActivity::Idle,
            location: Some(NativeOfficeCollaborationPresenceLocation::Markdown {
                anchor: 2,
                head: 8,
                surface: Some(NativeOfficeCollaborationMarkdownSurface::Source),
            }),
        })
        .unwrap();
    assert!(!updated.payload.is_empty());
    let snapshot = presence.snapshot().unwrap();
    assert_eq!(snapshot.participants.len(), 1);
    assert!(snapshot.participants[0].local);
    assert_eq!(
        snapshot.participants[0].state.activity,
        NativeOfficeCollaborationPresenceActivity::Idle
    );

    let removal = presence.disconnect().unwrap();
    let decoded = AwarenessUpdate::decode_v1(&removal.payload).unwrap();
    assert_eq!(
        decoded
            .clients
            .get(&ClientID::new(presence_client_id))
            .unwrap()
            .json
            .as_ref(),
        "null"
    );

    let after = store.inspect().unwrap();
    assert_eq!(after.current_sequence, before.current_sequence);
    assert_eq!(after.operation_count, before.operation_count);
    assert_eq!(after.document_state_sha256, before.document_state_sha256);
}

#[test]
fn native_presence_accepts_every_editor_location_schema() {
    let locations = [
        (
            NativeOfficeCollaborationArtifactKind::Document,
            NativeOfficeCollaborationPresenceLocation::Document { anchor: 3, head: 9 },
        ),
        (
            NativeOfficeCollaborationArtifactKind::Markdown,
            NativeOfficeCollaborationPresenceLocation::Markdown {
                anchor: 4,
                head: 10,
                surface: Some(NativeOfficeCollaborationMarkdownSurface::Visual),
            },
        ),
        (
            NativeOfficeCollaborationArtifactKind::Spreadsheet,
            NativeOfficeCollaborationPresenceLocation::Spreadsheet {
                sheet_id: "sheet-1".to_owned(),
                ranges: vec![NativeOfficeCollaborationSpreadsheetPresenceRange {
                    start_row: 1,
                    start_column: 2,
                    end_row: 3,
                    end_column: 4,
                }],
                active_cell: Some(NativeOfficeCollaborationSpreadsheetPresenceCell {
                    row: 1,
                    column: 2,
                }),
            },
        ),
        (
            NativeOfficeCollaborationArtifactKind::Presentation,
            NativeOfficeCollaborationPresenceLocation::Presentation {
                slide_id: "slide-1".to_owned(),
                element_ids: vec!["shape-1".to_owned(), "shape-2".to_owned()],
            },
        ),
        (
            NativeOfficeCollaborationArtifactKind::Pdf,
            NativeOfficeCollaborationPresenceLocation::Pdf {
                page_index: 2,
                annotation_id: Some("annotation-1".to_owned()),
            },
        ),
    ];

    for (kind, location) in locations {
        let mut presence =
            NativeOfficeCollaborationPresenceSession::new(manifest(kind), profile()).unwrap();
        presence
            .update(NativeOfficeCollaborationPresenceUpdate {
                activity: NativeOfficeCollaborationPresenceActivity::Active,
                location: Some(location.clone()),
            })
            .unwrap();
        assert_eq!(
            presence.snapshot().unwrap().participants[0].state.location,
            Some(location)
        );
    }
}

#[test]
fn native_presence_projects_remote_browser_state_and_honors_awareness_clocks() {
    let manifest = manifest(NativeOfficeCollaborationArtifactKind::Document);
    let mut presence =
        NativeOfficeCollaborationPresenceSession::new(manifest.clone(), profile()).unwrap();
    let sender = 424_242;
    let newer = remote_message(
        &manifest,
        sender,
        2,
        remote_state(
            &manifest,
            "away",
            json!({ "kind": "document", "anchor": 4, "head": 12 }),
        ),
    );
    let received = presence.receive(newer).unwrap();
    assert!(received.changed);
    assert_eq!(received.snapshot.participants.len(), 2);
    let remote = received
        .snapshot
        .participants
        .iter()
        .find(|participant| !participant.local)
        .unwrap();
    assert_eq!(remote.client_id, sender);
    assert_eq!(remote.state.actor.name, "Browser User");
    assert_eq!(
        remote.state.location,
        Some(NativeOfficeCollaborationPresenceLocation::Document {
            anchor: 4,
            head: 12,
        })
    );

    let stale = remote_message(
        &manifest,
        sender,
        1,
        remote_state(
            &manifest,
            "active",
            json!({ "kind": "document", "anchor": 0, "head": 0 }),
        ),
    );
    let received = presence.receive(stale).unwrap();
    assert!(!received.changed);
    let remote = received
        .snapshot
        .participants
        .iter()
        .find(|participant| !participant.local)
        .unwrap();
    assert_eq!(
        remote.state.activity,
        NativeOfficeCollaborationPresenceActivity::Away
    );

    let removed = presence.peer_left(sender).unwrap();
    assert!(removed.changed);
    assert_eq!(removed.snapshot.participants.len(), 1);
}

#[test]
fn native_presence_rejects_forged_or_kind_mismatched_state_atomically() {
    let manifest = manifest(NativeOfficeCollaborationArtifactKind::Spreadsheet);
    let mut presence =
        NativeOfficeCollaborationPresenceSession::new(manifest.clone(), profile()).unwrap();
    let sender = 424_243;
    let forged = remote_message(
        &manifest,
        sender,
        1,
        json!({
            "a3sOffice": {
                "protocol": NATIVE_OFFICE_COLLABORATION_PROTOCOL,
                "version": NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
                "artifactId": "another-artifact",
                "artifactKind": "spreadsheet",
                "namespace": manifest.namespace,
                "presenceId": "forged:1",
                "actor": { "id": "attacker", "name": "Attacker", "kind": "human" },
                "mode": "edit",
                "activity": "active",
                "location": {
                    "kind": "spreadsheet",
                    "sheetId": "sheet-1",
                    "ranges": [{
                        "startRow": 0,
                        "startColumn": 0,
                        "endRow": 0,
                        "endColumn": 0
                    }]
                }
            }
        }),
    );
    let error = presence.receive(forged).unwrap_err();
    assert_eq!(
        error.code,
        "office.collaboration.presence_identity_mismatch"
    );
    assert_eq!(presence.snapshot().unwrap().participants.len(), 1);

    let error = presence
        .update(NativeOfficeCollaborationPresenceUpdate {
            activity: NativeOfficeCollaborationPresenceActivity::Active,
            location: Some(NativeOfficeCollaborationPresenceLocation::Pdf {
                page_index: 0,
                annotation_id: None,
            }),
        })
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.presence_invalid");
    assert_eq!(presence.snapshot().unwrap().participants.len(), 1);
}

#[test]
fn native_presence_is_send_and_sync() {
    assert_send_sync::<NativeOfficeCollaborationPresenceSession>();
}
