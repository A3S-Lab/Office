use std::collections::BTreeSet;
use std::future::{ready, Ready};
use std::sync::{Arc, Mutex};

use a3s_boot::{BootRequest, HttpMethod, Result, WebSocketMessage};
use a3s_office::NativeOfficeCollaborationStore;
use a3s_office_collaboration_server::{build_application, CollaborationConfig};
use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use serde_json::{json, Value};
use yrs::sync::Awareness;
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{Doc, GetString, ReadTxn, Transact, Update};

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
    assert!(editor_outbound
        .lock()
        .unwrap()
        .iter()
        .any(|message| message.event == "collaboration.ready"));
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

async fn issue_ticket(app: &a3s_boot::BootApplication, actor_id: &str, mode: &str) -> String {
    let response = app
        .handle(
            BootRequest::new(HttpMethod::Post, "/api/collaboration/tickets")
                .with_header("authorization", "Bearer test-admin-token")
                .with_json(&ticket_body(actor_id, mode))
                .unwrap(),
        )
        .await;
    assert_eq!(response.status, 201);
    let body: Value = serde_json::from_slice(&response.body).unwrap();
    body["data"]["ticket"].as_str().unwrap().to_string()
}

fn ticket_body(actor_id: &str, mode: &str) -> Value {
    json!({
        "artifactId": "fixture-markdown",
        "artifactKind": "markdown",
        "actorId": actor_id,
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
