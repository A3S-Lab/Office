use std::collections::BTreeSet;
use std::future::{ready, Ready};
use std::sync::{Arc, Mutex};

use a3s_boot::{BootRequest, HttpMethod, Result, WebSocketMessage};
use a3s_office::{
    NativeOfficeCollaborationActorKind, NativeOfficeCollaborationArtifactKind,
    NativeOfficeCollaborationCreateRequest, NativeOfficeCollaborationEventsRequest,
    NativeOfficeCollaborationMode, NativeOfficeCollaborationMutation,
    NativeOfficeCollaborationMutationRequest, NativeOfficeCollaborationProjectedContent,
    NativeOfficeCollaborationStore, MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH,
    NATIVE_OFFICE_COLLABORATION_PROTOCOL, NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
};
use a3s_office_collaboration_server::{build_application, CollaborationConfig};
use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use serde_json::{json, Value};
use yrs::sync::Awareness;
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{
    Array, Doc, GetString, Map, ReadTxn, StateVector, Text, Transact, Update, Xml,
    XmlElementPrelim, XmlFragment, XmlOut, XmlTextPrelim,
};

const ORIGIN: &str = "http://localhost:4175";
const EMPTY_STATE_VECTOR_BASE64: &str = "AA==";
const MARKDOWN_UPDATE_BASE64: &str = "AQey8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtbWFya2Rvd24oARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhtYXJrZG93bigBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjQyNDI0Mjpicm93c2VyLWZpeHR1cmUEARphM3Mub2ZmaWNlLm1hcmtkb3duLnNvdXJjZRUjIFNoYXJlZAoKWWpzIHRvIFlycy4A";

#[tokio::test]
async fn boot_gateway_authenticates_persists_broadcasts_and_enforces_mode() {
    let temp = tempfile::tempdir().unwrap();
    let config = test_config(temp.path().to_path_buf());
    let app = build_application(config.clone()).unwrap();

    let editor_ticket = issue_ticket(&app, "editor-1", "edit").await;
    let viewer_ticket = issue_ticket(&app, "viewer-1", "view").await;
    let path = "/collaboration/markdown/fixture-markdown";
    let gateway = app.gateway_for(path).unwrap().clone();

    let forbidden_origin = gateway
        .connect_async_with_outbound(
            BootRequest::new(HttpMethod::Get, format!("{path}?ticket={editor_ticket}"))
                .with_header("origin", "https://attacker.example"),
            capture(Arc::new(Mutex::new(Vec::new()))),
        )
        .await;
    assert!(forbidden_origin.is_err());

    let editor_outbound = Arc::new(Mutex::new(Vec::new()));
    let editor = gateway
        .connect_async_with_outbound(
            authenticated_request(path, &editor_ticket),
            capture(Arc::clone(&editor_outbound)),
        )
        .await
        .unwrap();
    let viewer_outbound = Arc::new(Mutex::new(Vec::new()));
    let viewer = gateway
        .connect_async_with_outbound(
            authenticated_request(path, &viewer_ticket),
            capture(Arc::clone(&viewer_outbound)),
        )
        .await
        .unwrap();

    editor.dispatch(hello_message(424_242)).await.unwrap();
    viewer.dispatch(hello_message(515_151)).await.unwrap();
    let editor_ready = editor_outbound
        .lock()
        .unwrap()
        .iter()
        .find(|message| message.event == "collaboration.ready")
        .cloned()
        .unwrap();
    assert_eq!(editor_ready.data["actorId"], "editor-1");
    assert_eq!(editor_ready.data["actorName"], "Integration Test");
    assert_eq!(editor_ready.data["actorKind"], "human");
    assert_eq!(editor_ready.data["mode"], "edit");
    assert!(viewer_outbound
        .lock()
        .unwrap()
        .iter()
        .any(|message| message.event == "collaboration.ready"));
    editor_outbound.lock().unwrap().clear();
    viewer_outbound.lock().unwrap().clear();

    viewer
        .dispatch(document_message_with_payload(
            515_151,
            "sync-step-1",
            EMPTY_STATE_VECTOR_BASE64,
        ))
        .await
        .unwrap();
    assert!(viewer_outbound.lock().unwrap().iter().any(|message| {
        message.event == "collaboration.document" && message.data["type"] == "sync-step-2"
    }));
    viewer_outbound.lock().unwrap().clear();

    editor
        .dispatch(document_message(424_242, "sync-step-2"))
        .await
        .unwrap();
    assert!(editor_outbound
        .lock()
        .unwrap()
        .iter()
        .any(|message| message.event == "collaboration.ack"));
    assert!(viewer_outbound
        .lock()
        .unwrap()
        .iter()
        .any(|message| message.event == "collaboration.document"));

    let rejected = viewer
        .dispatch(document_message(515_151, "sync-step-2"))
        .await
        .unwrap()
        .unwrap();
    assert_eq!(rejected.event, "collaboration.error");
    assert_eq!(rejected.data["code"], "FORBIDDEN");

    viewer_outbound.lock().unwrap().clear();
    editor
        .dispatch(awareness_message(424_242, "editor-1", "edit"))
        .await
        .unwrap();
    assert!(viewer_outbound
        .lock()
        .unwrap()
        .iter()
        .any(|message| message.event == "collaboration.awareness"));

    let store_path = std::fs::read_dir(&config.data_dir)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .find(|path| path.is_dir())
        .unwrap();
    let store = NativeOfficeCollaborationStore::open(store_path).unwrap();
    let snapshot = store.synchronize(None).unwrap();
    let document = Doc::with_client_id(818_181);
    document
        .transact_mut()
        .apply_update(Update::decode_v1(&snapshot.update).unwrap())
        .unwrap();
    let transaction = document.transact();
    let markdown = transaction
        .get_text("a3s.office.markdown.source")
        .unwrap()
        .get_string(&transaction);
    assert_eq!(markdown, "# Shared\n\nYjs to Yrs.");

    viewer_outbound.lock().unwrap().clear();
    editor.close().await.unwrap();
    assert!(viewer_outbound.lock().unwrap().iter().any(|message| {
        message.event == "collaboration.peer-left" && message.data["senderClientId"] == 424_242
    }));
    viewer.close().await.unwrap();
}

#[tokio::test]
async fn ticket_api_uses_the_repository_response_contract() {
    let temp = tempfile::tempdir().unwrap();
    let app = build_application(test_config(temp.path().to_path_buf())).unwrap();
    let unauthorized = app
        .handle(
            BootRequest::new(HttpMethod::Post, "/api/collaboration/tickets")
                .with_json(&ticket_body("editor-1", "edit"))
                .unwrap(),
        )
        .await;
    assert_eq!(unauthorized.status, 401);
    let body: Value = serde_json::from_slice(&unauthorized.body).unwrap();
    assert_eq!(body["code"], 401);
    assert_eq!(body["statusCode"], "UNAUTHORIZED");
    assert!(body["requestId"].is_string());
    assert!(body["timestamp"].is_string());

    let health = app
        .handle(BootRequest::new(
            HttpMethod::Get,
            "/api/collaboration/healthz",
        ))
        .await;
    assert_eq!(health.status, 200);
    let body: Value = serde_json::from_slice(&health.body).unwrap();
    assert_eq!(body["data"]["status"], "ok");
    assert_eq!(body["data"]["protocol"], "a3s.office.collaboration");
}

#[tokio::test]
async fn comment_ticket_persists_selection_review_but_cannot_edit_document_text() {
    let temp = tempfile::tempdir().unwrap();
    let config = test_config(temp.path().to_path_buf());
    let app = build_application(config.clone()).unwrap();
    let path = "/collaboration/document/fixture-document";
    let editor_ticket = issue_ticket_for(
        &app,
        "fixture-document",
        "document",
        "editor-2",
        "Evan Editor",
        "edit",
    )
    .await;
    let commenter_ticket = issue_ticket_for(
        &app,
        "fixture-document",
        "document",
        "reviewer-2",
        "Ada Reviewer",
        "comment",
    )
    .await;
    let gateway = app.gateway_for(path).unwrap().clone();
    let editor_outbound = Arc::new(Mutex::new(Vec::new()));
    let editor = gateway
        .connect_async_with_outbound(
            authenticated_request(path, &editor_ticket),
            capture(Arc::clone(&editor_outbound)),
        )
        .await
        .unwrap();
    let commenter_outbound = Arc::new(Mutex::new(Vec::new()));
    let commenter = gateway
        .connect_async_with_outbound(
            authenticated_request(path, &commenter_ticket),
            capture(Arc::clone(&commenter_outbound)),
        )
        .await
        .unwrap();
    editor.dispatch(hello_message(616_161)).await.unwrap();
    commenter.dispatch(hello_message(626_262)).await.unwrap();
    editor_outbound.lock().unwrap().clear();
    commenter_outbound.lock().unwrap().clear();

    let bootstrap = document_bootstrap_update();
    editor
        .dispatch(document_message_for(
            "fixture-document",
            "document",
            616_161,
            "sync-step-2",
            &STANDARD.encode(&bootstrap),
        ))
        .await
        .unwrap();
    let producer_temp = tempfile::tempdir().unwrap();
    let producer = NativeOfficeCollaborationStore::create(NativeOfficeCollaborationCreateRequest {
        store: producer_temp.path().join("reviewer-replica"),
        artifact_id: "fixture-document".to_string(),
        kind: NativeOfficeCollaborationArtifactKind::Document,
        actor_id: "reviewer-2".to_string(),
        actor_kind: NativeOfficeCollaborationActorKind::Human,
        mode: NativeOfficeCollaborationMode::Comment,
        operation_id: "create-reviewer-replica".to_string(),
        namespace: None,
        client_id: Some(636_363),
        initial_update: Some(bootstrap),
    })
    .unwrap();
    producer
        .mutate(NativeOfficeCollaborationMutationRequest {
            operation_id: "create-selection-comment".to_string(),
            actor_id: "reviewer-2".to_string(),
            mode: NativeOfficeCollaborationMode::Comment,
            expected_artifact_id: "fixture-document".to_string(),
            expected_kind: NativeOfficeCollaborationArtifactKind::Document,
            mutation: NativeOfficeCollaborationMutation::DocumentCommentCreate {
                comment_id: "review-comment-1".to_string(),
                paragraph_id: "00000001".to_string(),
                expected_text_id: "00000002".to_string(),
                start_utf16: 6,
                end_utf16: 12,
                expected_text: "review".to_string(),
                author: "Ada Reviewer".to_string(),
                created_at: "2026-08-17T00:00:00.000Z".to_string(),
                text: "Clarify this review point.".to_string(),
            },
            if_state_vector: None,
        })
        .unwrap();
    let comment_update = producer
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(0),
            limit: MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH,
        })
        .unwrap()
        .updates
        .into_iter()
        .find(|event| event.operation_id == "create-selection-comment")
        .unwrap()
        .update;
    commenter
        .dispatch(document_message_for(
            "fixture-document",
            "document",
            626_262,
            "update",
            &STANDARD.encode(comment_update),
        ))
        .await
        .unwrap();
    assert!(commenter_outbound
        .lock()
        .unwrap()
        .iter()
        .any(|message| message.event == "collaboration.ack"));
    assert!(editor_outbound
        .lock()
        .unwrap()
        .iter()
        .any(|message| message.event == "collaboration.document"));

    let store_path = std::fs::read_dir(&config.data_dir)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .find(|path| path.is_dir())
        .unwrap();
    let store = NativeOfficeCollaborationStore::open(store_path).unwrap();
    let NativeOfficeCollaborationProjectedContent::Document { comments, .. } =
        store.project().unwrap().content
    else {
        panic!("expected Document projection");
    };
    assert_eq!(comments.len(), 1);
    assert_eq!(comments[0].actor_id.as_deref(), Some("reviewer-2"));
    assert_eq!(comments[0].anchors[0].text, "review");

    let attacker = Doc::with_client_id(646_464);
    attacker
        .transact_mut()
        .apply_update(Update::decode_v1(&store.synchronize(None).unwrap().update).unwrap())
        .unwrap();
    let before_attack = attacker.transact().state_vector();
    let fragment = attacker.get_or_insert_xml_fragment("a3s.office.document.content");
    let text = fragment
        .successors(&attacker.transact())
        .find_map(|node| match node {
            XmlOut::Text(text) => Some(text),
            _ => None,
        })
        .unwrap();
    text.insert(&mut attacker.transact_mut(), 0, "FORGED ");
    let forged = attacker
        .transact()
        .encode_state_as_update_v1(&before_attack);
    commenter_outbound.lock().unwrap().clear();
    let rejected = commenter
        .dispatch(document_message_for(
            "fixture-document",
            "document",
            626_262,
            "update",
            &STANDARD.encode(forged),
        ))
        .await
        .unwrap()
        .unwrap();
    assert_eq!(rejected.event, "collaboration.error");
    assert_eq!(rejected.data["code"], "FORBIDDEN");
    let NativeOfficeCollaborationProjectedContent::Document { plain_text, .. } =
        store.project().unwrap().content
    else {
        panic!("expected Document projection");
    };
    assert_eq!(plain_text, "Hello review world.");

    editor.close().await.unwrap();
    commenter.close().await.unwrap();
}

async fn issue_ticket(app: &a3s_boot::BootApplication, actor_id: &str, mode: &str) -> String {
    issue_ticket_for(
        app,
        "fixture-markdown",
        "markdown",
        actor_id,
        "Integration Test",
        mode,
    )
    .await
}

async fn issue_ticket_for(
    app: &a3s_boot::BootApplication,
    artifact_id: &str,
    artifact_kind: &str,
    actor_id: &str,
    actor_name: &str,
    mode: &str,
) -> String {
    let response = app
        .handle(
            BootRequest::new(HttpMethod::Post, "/api/collaboration/tickets")
                .with_header("authorization", "Bearer test-admin-token")
                .with_json(&json!({
                    "artifactId": artifact_id,
                    "artifactKind": artifact_kind,
                    "actorId": actor_id,
                    "actorName": actor_name,
                    "actorKind": "human",
                    "mode": mode,
                }))
                .unwrap(),
        )
        .await;
    assert_eq!(response.status, 201);
    let body: Value = serde_json::from_slice(&response.body).unwrap();
    body["data"]["ticket"].as_str().unwrap().to_string()
}

fn document_bootstrap_update() -> Vec<u8> {
    let document = Doc::with_client_id(606_060);
    let metadata = document.get_or_insert_map("a3s.office.metadata");
    let initializers = document.get_or_insert_array("a3s.office.bootstrap.initializers");
    let fragment = document.get_or_insert_xml_fragment("a3s.office.document.content");
    document.get_or_insert_map("a3s.office.document.options");
    document.get_or_insert_map("a3s.office.document.comments");
    document.get_or_insert_array("a3s.office.document.comment-order");
    document.get_or_insert_map("a3s.office.document.bibliography");
    document.get_or_insert_map("a3s.office.document.bibliography.sources");
    document.get_or_insert_array("a3s.office.document.bibliography.source-order");
    document.get_or_insert_array("a3s.office.document.record-claims");
    let mut transaction = document.transact_mut();
    metadata.insert(
        &mut transaction,
        "protocol",
        NATIVE_OFFICE_COLLABORATION_PROTOCOL,
    );
    metadata.insert(
        &mut transaction,
        "version",
        i64::from(NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION),
    );
    metadata.insert(&mut transaction, "artifactId", "fixture-document");
    metadata.insert(&mut transaction, "kind", "document");
    metadata.insert(&mut transaction, "initialized", true);
    initializers.push_back(&mut transaction, "606060:server-integration");
    let section = fragment.push_back(&mut transaction, XmlElementPrelim::empty("documentSection"));
    section.insert_attribute(&mut transaction, "id", "document-section-1");
    let paragraph = section.push_back(&mut transaction, XmlElementPrelim::empty("paragraph"));
    paragraph.insert_attribute(&mut transaction, "paragraphId", "00000001");
    paragraph.insert_attribute(&mut transaction, "textId", "00000002");
    paragraph.push_back(&mut transaction, XmlTextPrelim::new("Hello review world."));
    drop(transaction);
    let update = document
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    update
}

fn ticket_body(actor_id: &str, mode: &str) -> Value {
    json!({
        "artifactId": "fixture-markdown",
        "artifactKind": "markdown",
        "actorId": actor_id,
        "actorName": "Integration Test",
        "actorKind": "human",
        "mode": mode,
    })
}

fn authenticated_request(path: &str, ticket: &str) -> BootRequest {
    BootRequest::new(HttpMethod::Get, format!("{path}?ticket={ticket}"))
        .with_header("origin", ORIGIN)
}

fn hello_message(client_id: u64) -> WebSocketMessage {
    WebSocketMessage::new(
        "collaboration.hello",
        json!({
            "protocol": "a3s.office.collaboration",
            "version": 1,
            "senderClientId": client_id,
        }),
    )
}

fn document_message(client_id: u64, message_type: &str) -> WebSocketMessage {
    document_message_with_payload(client_id, message_type, MARKDOWN_UPDATE_BASE64)
}

fn document_message_with_payload(
    client_id: u64,
    message_type: &str,
    payload_base64: &str,
) -> WebSocketMessage {
    WebSocketMessage::new(
        "collaboration.document",
        json!({
            "protocol": "a3s.office.collaboration",
            "version": 1,
            "artifactId": "fixture-markdown",
            "artifactKind": "markdown",
            "namespace": "a3s.office",
            "senderClientId": client_id,
            "type": message_type,
            "payloadBase64": payload_base64,
        }),
    )
}

fn document_message_for(
    artifact_id: &str,
    artifact_kind: &str,
    client_id: u64,
    message_type: &str,
    payload_base64: &str,
) -> WebSocketMessage {
    WebSocketMessage::new(
        "collaboration.document",
        json!({
            "protocol": "a3s.office.collaboration",
            "version": 1,
            "artifactId": artifact_id,
            "artifactKind": artifact_kind,
            "namespace": "a3s.office",
            "senderClientId": client_id,
            "type": message_type,
            "payloadBase64": payload_base64,
        }),
    )
}

fn awareness_message(client_id: u64, actor_id: &str, mode: &str) -> WebSocketMessage {
    let mut awareness = Awareness::new(Doc::with_client_id(client_id));
    awareness
        .set_local_state(json!({
            "a3sOffice": {
                "protocol": "a3s.office.collaboration",
                "version": 1,
                "artifactId": "fixture-markdown",
                "artifactKind": "markdown",
                "namespace": "a3s.office",
                "mode": mode,
                "actor": {
                    "id": actor_id,
                    "kind": "human",
                    "name": "Integration Test",
                },
            },
        }))
        .unwrap();
    let payload = STANDARD.encode(awareness.update().unwrap().encode_v1());
    WebSocketMessage::new(
        "collaboration.awareness",
        json!({
            "protocol": "a3s.office.collaboration",
            "version": 1,
            "artifactId": "fixture-markdown",
            "artifactKind": "markdown",
            "namespace": "a3s.office",
            "senderClientId": client_id,
            "payloadBase64": payload,
        }),
    )
}

fn capture(
    messages: Arc<Mutex<Vec<WebSocketMessage>>>,
) -> impl Fn(WebSocketMessage) -> Ready<Result<()>> + Send + Sync + 'static {
    move |message| {
        messages.lock().unwrap().push(message);
        ready(Ok(()))
    }
}

fn test_config(data_dir: std::path::PathBuf) -> CollaborationConfig {
    CollaborationConfig {
        bind: ([127, 0, 0, 1], 0).into(),
        public_ws_url: "ws://127.0.0.1:8787/collaboration".to_string(),
        data_dir,
        namespace: "a3s.office".to_string(),
        allowed_origins: BTreeSet::from([ORIGIN.to_string()]),
        ticket_ttl_seconds: 300,
        poll_interval_milliseconds: 25,
        max_document_payload_bytes: 64 * 1024 * 1024,
        max_awareness_payload_bytes: 256 * 1024,
        ticket_secret: "test-ticket-secret-with-at-least-32-bytes".to_string(),
        admin_token: "test-admin-token".to_string(),
    }
}
