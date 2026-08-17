use std::process::Stdio;
use std::time::Duration;

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};
use yrs::updates::decoder::Decode;
use yrs::{
    Any, Doc, GetString, Map, Out, ReadTxn, StateVector, Text, Transact, Update, Xml, XmlFragment,
    XmlOut,
};

const YJS_MARKDOWN_UPDATE_BASE64: &str = "AQey8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtbWFya2Rvd24oARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhtYXJrZG93bigBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjQyNDI0Mjpicm93c2VyLWZpeHR1cmUEARphM3Mub2ZmaWNlLm1hcmtkb3duLnNvdXJjZRUjIFNoYXJlZAoKWWpzIHRvIFlycy4A";
const YJS_DOCUMENT_UPDATE_BASE64: &str = "AQ+y8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtZG9jdW1lbnQoARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhkb2N1bWVudCgBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjQyNDI0Mjpicm93c2VyLWZpeHR1cmUHARthM3Mub2ZmaWNlLmRvY3VtZW50LmNvbnRlbnQDD2RvY3VtZW50U2VjdGlvbgcAsvIZBgMJcGFyYWdyYXBoBwCy8hkHBgQAsvIZCBBIZWxsbyDwn5iAIHdvcmxkKACy8hkHC3BhcmFncmFwaElkAXcIMDAwMDAwMDEoALLyGQcGdGV4dElkAXcIMDAwMDAwMDIoALLyGQYCaWQBdxJkb2N1bWVudC1zZWN0aW9uLTEoARthM3Mub2ZmaWNlLmRvY3VtZW50Lm9wdGlvbnMJcGFnZUNvbG9yAXcHI0Y4RkFGQygBG2Ezcy5vZmZpY2UuZG9jdW1lbnQub3B0aW9ucwx0cmFja0NoYW5nZXMBeAA=";

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

#[tokio::test]
async fn native_standard_mcp_runs_a_resumable_collaboration_event_loop() {
    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let replica = temp.path().join("coding-agent.replica");
    let replica_text = replica.to_str().unwrap();
    let document_replica = temp.path().join("document-coding-agent.replica");
    let document_replica_text = document_replica.to_str().unwrap();
    let comment_replica = temp.path().join("document-review-agent.replica");
    let comment_replica_text = comment_replica.to_str().unwrap();
    let mut child = tokio::process::Command::new(binary())
        .args(["mcp"])
        .env("A3S_OFFICECLI_EXECUTABLE", &provider)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .kill_on_drop(true)
        .spawn()
        .unwrap();
    let mut stdin = child.stdin.take().unwrap();
    let mut stdout = BufReader::new(child.stdout.take().unwrap());
    let mut stderr = child.stderr.take().unwrap();

    request(
        &mut stdin,
        &mut stdout,
        serde_json::json!({
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2025-06-18",
                "capabilities": {},
                "clientInfo": { "name": "office-collaboration-test", "version": "1" }
            }
        }),
        TIMEOUT,
    )
    .await;
    stdin
        .write_all(
            b"{\"jsonrpc\":\"2.0\",\"method\":\"notifications/initialized\",\"params\":{}}\n",
        )
        .await
        .unwrap();
    stdin.flush().await.unwrap();

    let tools = request(
        &mut stdin,
        &mut stdout,
        serde_json::json!({"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}),
        TIMEOUT,
    )
    .await;
    let tools = tools["result"]["tools"].as_array().unwrap();
    for name in [
        "office_collaboration_create",
        "office_collaboration_inspect",
        "office_collaboration_read",
        "office_collaboration_diff",
        "office_collaboration_events",
        "office_collaboration_apply",
        "office_collaboration_mutate",
        "office_collaboration_checkpoint",
    ] {
        assert!(
            tools.iter().any(|tool| tool["name"] == name),
            "missing {name}"
        );
    }
    let events_schema = tools
        .iter()
        .find(|tool| tool["name"] == "office_collaboration_events")
        .unwrap()["inputSchema"]
        .to_string();
    for expected in ["afterSequence", "limit", "includeUpdates"] {
        assert!(events_schema.contains(expected), "missing {expected}");
    }
    let mutation_schema = tools
        .iter()
        .find(|tool| tool["name"] == "office_collaboration_mutate")
        .unwrap()["inputSchema"]
        .to_string();
    for expected in [
        "markdown-splice",
        "indexUtf16",
        "deleteUtf16",
        "document-replace-text",
        "document-replace-paragraph",
        "expectedMatches",
        "document-set-page-color",
        "pageColor",
        "document-insert-paragraph",
        "anchorParagraphId",
        "paragraphId",
        "textId",
        "document-delete-paragraph",
        "expectedTextId",
        "document-comment-create",
        "document-comment-reply",
        "document-comment-set-resolved",
        "document-comment-delete",
        "document-suggestion-create",
        "document-suggestion-decide",
        "insertionId",
        "deletionId",
        "expectedActorId",
        "expectedCreatedAt",
        "decidedBy",
        "decidedAt",
        "commentId",
        "replyId",
        "startUtf16",
        "endUtf16",
        "spreadsheet-set-cell",
        "spreadsheet-delete-cell",
        "spreadsheet-batch-cells",
        "changes",
        "expectedCell",
        "nextCell",
    ] {
        assert!(mutation_schema.contains(expected), "missing {expected}");
    }

    let created = call(
        &mut stdin,
        &mut stdout,
        3,
        "office_collaboration_create",
        serde_json::json!({
            "store": replica_text,
            "artifactId": "fixture-markdown",
            "kind": "markdown",
            "actorId": "coding-agent-9",
            "actorKind": "agent",
            "mode": "edit",
            "operationId": "join-browser-1",
            "clientId": 900009,
            "initialUpdateBase64": YJS_MARKDOWN_UPDATE_BASE64
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(created["result"]["isError"], true, "{created}");
    let replica_value = &created["result"]["structuredContent"]["replica"];
    assert_eq!(created["result"]["structuredContent"]["action"], "joined");
    assert_eq!(replica_value["manifest"]["actorId"], "coding-agent-9");
    assert_eq!(replica_value["currentSequence"], 0);
    let initial_state_vector = replica_value["stateVectorBase64"]
        .as_str()
        .unwrap()
        .to_owned();

    let inspected = call(
        &mut stdin,
        &mut stdout,
        30,
        "office_collaboration_inspect",
        serde_json::json!({"store": replica_text}),
        TIMEOUT,
    )
    .await;
    assert_ne!(inspected["result"]["isError"], true, "{inspected}");
    assert_eq!(
        inspected["result"]["structuredContent"]["manifest"]["artifactId"],
        "fixture-markdown"
    );
    assert!(inspected["result"]["structuredContent"]["stateVectorBase64"].is_string());

    let projected = call(
        &mut stdin,
        &mut stdout,
        31,
        "office_collaboration_read",
        serde_json::json!({"store": replica_text}),
        TIMEOUT,
    )
    .await;
    assert_ne!(projected["result"]["isError"], true, "{projected}");
    assert_eq!(
        projected["result"]["structuredContent"]["content"]["kind"],
        "markdown"
    );
    assert_eq!(
        projected["result"]["structuredContent"]["content"]["source"],
        "# Shared\n\nYjs to Yrs."
    );
    assert!(projected["result"]["structuredContent"]["stateVectorBase64"].is_string());

    let ready = call(
        &mut stdin,
        &mut stdout,
        4,
        "office_collaboration_events",
        serde_json::json!({
            "store": replica_text,
            "limit": 1,
            "includeUpdates": true
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(ready["result"]["isError"], true, "{ready}");
    let ready = &ready["result"]["structuredContent"];
    assert_eq!(ready["startingSequence"], 0);
    assert_eq!(ready["cursorSequence"], 0);
    assert_eq!(ready["eventCount"], 0);

    let peer = Doc::with_client_id(800_009);
    let notes = peer.get_or_insert_text("agent.notes");
    notes.insert(&mut peer.transact_mut(), 0, "reviewed live");
    let live_update = peer
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    let apply_arguments = serde_json::json!({
        "store": replica_text,
        "operationId": "agent-update-1",
        "actorId": "coding-agent-9",
        "mode": "edit",
        "artifactId": "fixture-markdown",
        "kind": "markdown",
        "updateBase64": STANDARD.encode(&live_update)
    });
    let applied = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_collaboration_apply",
        apply_arguments.clone(),
        TIMEOUT,
    )
    .await;
    assert_ne!(applied["result"]["isError"], true, "{applied}");
    let applied = &applied["result"]["structuredContent"];
    assert_eq!(applied["sequence"], 1);
    assert_eq!(applied["duplicate"], false);
    assert!(applied["stateVectorBase64"].is_string());

    let replay = call(
        &mut stdin,
        &mut stdout,
        6,
        "office_collaboration_apply",
        apply_arguments,
        TIMEOUT,
    )
    .await;
    assert_eq!(replay["result"]["structuredContent"]["duplicate"], true);

    let metadata_events = call(
        &mut stdin,
        &mut stdout,
        7,
        "office_collaboration_events",
        serde_json::json!({
            "store": replica_text,
            "afterSequence": 0,
            "limit": 1
        }),
        TIMEOUT,
    )
    .await;
    let metadata_events = &metadata_events["result"]["structuredContent"];
    assert_eq!(metadata_events["cursorSequence"], 1);
    assert_eq!(metadata_events["eventCount"], 1);
    assert_eq!(
        metadata_events["updates"][0]["operationId"],
        "agent-update-1"
    );
    assert!(metadata_events["updates"][0].get("updateBase64").is_none());

    let payload_events = call(
        &mut stdin,
        &mut stdout,
        8,
        "office_collaboration_events",
        serde_json::json!({
            "store": replica_text,
            "afterSequence": 0,
            "limit": 1,
            "includeUpdates": true
        }),
        TIMEOUT,
    )
    .await;
    assert_eq!(
        payload_events["result"]["structuredContent"]["updates"][0]["updateBase64"],
        STANDARD.encode(&live_update)
    );

    let diff = call(
        &mut stdin,
        &mut stdout,
        9,
        "office_collaboration_diff",
        serde_json::json!({
            "store": replica_text,
            "stateVectorBase64": initial_state_vector
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(diff["result"]["isError"], true, "{diff}");
    assert!(
        diff["result"]["structuredContent"]["updateBytes"]
            .as_u64()
            .unwrap()
            > 2
    );

    let mutation_arguments = serde_json::json!({
        "store": replica_text,
        "operationId": "typed-markdown-1",
        "actorId": "coding-agent-9",
        "mode": "edit",
        "artifactId": "fixture-markdown",
        "kind": "markdown",
        "mutation": {
            "type": "markdown-replace",
            "markdown": "# MCP typed 😀"
        }
    });
    let mutated = call(
        &mut stdin,
        &mut stdout,
        90,
        "office_collaboration_mutate",
        mutation_arguments.clone(),
        TIMEOUT,
    )
    .await;
    assert_ne!(mutated["result"]["isError"], true, "{mutated}");
    assert_eq!(mutated["result"]["structuredContent"]["action"], "mutated");
    assert_eq!(mutated["result"]["structuredContent"]["sequence"], 2);
    let mutation_replay = call(
        &mut stdin,
        &mut stdout,
        91,
        "office_collaboration_mutate",
        mutation_arguments,
        TIMEOUT,
    )
    .await;
    assert_eq!(
        mutation_replay["result"]["structuredContent"]["duplicate"],
        true
    );
    let mutation_event = call(
        &mut stdin,
        &mut stdout,
        92,
        "office_collaboration_events",
        serde_json::json!({
            "store": replica_text,
            "afterSequence": 1,
            "limit": 1,
            "includeUpdates": true
        }),
        TIMEOUT,
    )
    .await;
    let mutation_event = &mutation_event["result"]["structuredContent"]["updates"][0];
    assert_eq!(mutation_event["operationId"], "typed-markdown-1");
    assert_eq!(mutation_event["operationKind"], "mutate");

    let document_created = call(
        &mut stdin,
        &mut stdout,
        93,
        "office_collaboration_create",
        serde_json::json!({
            "store": document_replica_text,
            "artifactId": "fixture-document",
            "kind": "document",
            "actorId": "coding-agent-10",
            "actorKind": "agent",
            "mode": "edit",
            "operationId": "join-browser-document-1",
            "clientId": 900010,
            "initialUpdateBase64": YJS_DOCUMENT_UPDATE_BASE64
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(
        document_created["result"]["isError"], true,
        "{document_created}"
    );
    let document_mutated = call(
        &mut stdin,
        &mut stdout,
        94,
        "office_collaboration_mutate",
        serde_json::json!({
            "store": document_replica_text,
            "operationId": "typed-document-text-1",
            "actorId": "coding-agent-10",
            "mode": "edit",
            "artifactId": "fixture-document",
            "kind": "document",
            "mutation": {
                "type": "document-replace-text",
                "search": "😀",
                "replacement": "🦀",
                "expectedMatches": 1
            }
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(
        document_mutated["result"]["isError"], true,
        "{document_mutated}"
    );
    assert_eq!(
        document_mutated["result"]["structuredContent"]["sequence"],
        1
    );
    let document_page_color = call(
        &mut stdin,
        &mut stdout,
        95,
        "office_collaboration_mutate",
        serde_json::json!({
            "store": document_replica_text,
            "operationId": "typed-document-page-color-1",
            "actorId": "coding-agent-10",
            "mode": "edit",
            "artifactId": "fixture-document",
            "kind": "document",
            "mutation": {
                "type": "document-set-page-color",
                "pageColor": "#101828"
            }
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(
        document_page_color["result"]["isError"], true,
        "{document_page_color}"
    );
    let document_inserted = call(
        &mut stdin,
        &mut stdout,
        96,
        "office_collaboration_mutate",
        serde_json::json!({
            "store": document_replica_text,
            "operationId": "typed-document-insert-1",
            "actorId": "coding-agent-10",
            "mode": "edit",
            "artifactId": "fixture-document",
            "kind": "document",
            "mutation": {
                "type": "document-insert-paragraph",
                "anchorParagraphId": "00000001",
                "position": "after",
                "paragraphId": "00000040",
                "textId": "00000041",
                "text": "MCP paragraph"
            }
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(
        document_inserted["result"]["isError"], true,
        "{document_inserted}"
    );
    assert_eq!(
        document_inserted["result"]["structuredContent"]["sequence"],
        3
    );
    let document_diff = call(
        &mut stdin,
        &mut stdout,
        97,
        "office_collaboration_diff",
        serde_json::json!({"store": document_replica_text}),
        TIMEOUT,
    )
    .await;
    assert_ne!(document_diff["result"]["isError"], true, "{document_diff}");
    let document_peer = Doc::with_client_id(700_010);
    document_peer
        .transact_mut()
        .apply_update(
            Update::decode_v1(
                &STANDARD
                    .decode(
                        document_diff["result"]["structuredContent"]["updateBase64"]
                            .as_str()
                            .unwrap(),
                    )
                    .unwrap(),
            )
            .unwrap(),
        )
        .unwrap();
    let document_fragment = document_peer.get_or_insert_xml_fragment("a3s.office.document.content");
    let document_options = document_peer.get_or_insert_map("a3s.office.document.options");
    let document_transaction = document_peer.transact();
    let document_paragraphs = document_fragment
        .successors(&document_transaction)
        .filter_map(|node| match node {
            XmlOut::Element(element) if element.tag().as_ref() == "paragraph" => Some((
                element
                    .get_attribute(&document_transaction, "paragraphId")
                    .and_then(|value| match value {
                        Out::Any(Any::String(value)) => Some(value.to_string()),
                        _ => None,
                    }),
                element
                    .children(&document_transaction)
                    .filter_map(|child| match child {
                        XmlOut::Text(text) => Some(text.get_string(&document_transaction)),
                        _ => None,
                    })
                    .collect::<String>(),
            )),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(
        document_paragraphs,
        vec![
            (Some("00000001".to_owned()), "Hello 🦀 world".to_owned()),
            (Some("00000040".to_owned()), "MCP paragraph".to_owned()),
        ]
    );
    assert!(matches!(
        document_options.get(&document_transaction, "pageColor"),
        Some(Out::Any(Any::String(value))) if value.as_ref() == "#101828"
    ));
    drop(document_transaction);

    let comment_created = call(
        &mut stdin,
        &mut stdout,
        98,
        "office_collaboration_create",
        serde_json::json!({
            "store": comment_replica_text,
            "artifactId": "fixture-document",
            "kind": "document",
            "actorId": "review-agent-10",
            "actorKind": "agent",
            "mode": "comment",
            "operationId": "join-review-document-1",
            "clientId": 900110,
            "initialUpdateBase64": YJS_DOCUMENT_UPDATE_BASE64
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(
        comment_created["result"]["isError"], true,
        "{comment_created}"
    );
    let comment_mutated = call(
        &mut stdin,
        &mut stdout,
        99,
        "office_collaboration_mutate",
        serde_json::json!({
            "store": comment_replica_text,
            "operationId": "mcp-comment-create-1",
            "actorId": "review-agent-10",
            "mode": "comment",
            "artifactId": "fixture-document",
            "kind": "document",
            "mutation": {
                "type": "document-comment-create",
                "commentId": "comment-mcp-1",
                "paragraphId": "00000001",
                "expectedTextId": "00000002",
                "startUtf16": 6,
                "endUtf16": 8,
                "expectedText": "😀",
                "author": "Review Agent",
                "createdAt": "2026-08-17T00:00:00.000Z",
                "text": "Review this symbol."
            }
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(
        comment_mutated["result"]["isError"], true,
        "{comment_mutated}"
    );
    assert_eq!(
        comment_mutated["result"]["structuredContent"]["sequence"],
        1
    );
    let comment_reply = call(
        &mut stdin,
        &mut stdout,
        100,
        "office_collaboration_mutate",
        serde_json::json!({
            "store": comment_replica_text,
            "operationId": "mcp-comment-reply-1",
            "actorId": "review-agent-10",
            "mode": "comment",
            "artifactId": "fixture-document",
            "kind": "document",
            "mutation": {
                "type": "document-comment-reply",
                "commentId": "comment-mcp-1",
                "replyId": "reply-mcp-1",
                "author": "Review Agent",
                "createdAt": "2026-08-17T00:01:00.000Z",
                "text": "Confirmed."
            }
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(comment_reply["result"]["isError"], true, "{comment_reply}");
    let comment_read = call(
        &mut stdin,
        &mut stdout,
        101,
        "office_collaboration_read",
        serde_json::json!({ "store": comment_replica_text }),
        TIMEOUT,
    )
    .await;
    assert_ne!(comment_read["result"]["isError"], true, "{comment_read}");
    let comment = &comment_read["result"]["structuredContent"]["content"]["comments"][0];
    assert_eq!(comment["id"], "comment-mcp-1");
    assert_eq!(comment["actorId"], "review-agent-10");
    assert_eq!(comment["replies"][0]["id"], "reply-mcp-1");
    assert_eq!(comment["anchors"][0]["startUtf16"], 6);
    assert_eq!(comment["anchors"][0]["endUtf16"], 8);
    assert_eq!(comment["detached"], false);

    let checkpoint = call(
        &mut stdin,
        &mut stdout,
        10,
        "office_collaboration_checkpoint",
        serde_json::json!({
            "store": replica_text,
            "operationId": "checkpoint-1",
            "actorId": "coding-agent-9",
            "mode": "edit",
            "artifactId": "fixture-markdown",
            "kind": "markdown"
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(checkpoint["result"]["isError"], true, "{checkpoint}");
    assert_eq!(checkpoint["result"]["structuredContent"]["sequence"], 2);
    assert_eq!(
        checkpoint["result"]["structuredContent"]["compactedUpdates"],
        2
    );

    let compacted = call(
        &mut stdin,
        &mut stdout,
        11,
        "office_collaboration_events",
        serde_json::json!({
            "store": replica_text,
            "afterSequence": 0,
            "limit": 1,
            "includeUpdates": true
        }),
        TIMEOUT,
    )
    .await;
    let compacted = &compacted["result"]["structuredContent"];
    assert_eq!(compacted["cursorSequence"], 2);
    assert_eq!(compacted["reset"]["reason"], "history-compacted");
    let reset_update = STANDARD
        .decode(compacted["reset"]["updateBase64"].as_str().unwrap())
        .unwrap();
    let reset_peer = Doc::with_client_id(700_007);
    reset_peer
        .transact_mut()
        .apply_update(Update::decode_v1(&reset_update).unwrap())
        .unwrap();
    let transaction = reset_peer.transact();
    assert_eq!(
        transaction
            .get_text("a3s.office.markdown.source")
            .unwrap()
            .get_string(&transaction),
        "# MCP typed 😀"
    );
    assert_eq!(
        transaction
            .get_text("agent.notes")
            .unwrap()
            .get_string(&transaction),
        "reviewed live"
    );
    drop(transaction);

    let left = call(
        &mut stdin,
        &mut stdout,
        31,
        "office_collaboration_checkpoint",
        serde_json::json!({
            "store": replica_text,
            "operationId": "leave-1",
            "actorId": "coding-agent-9",
            "mode": "edit",
            "artifactId": "fixture-markdown",
            "kind": "markdown",
            "leave": true
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(left["result"]["isError"], true, "{left}");
    assert_eq!(left["result"]["structuredContent"]["action"], "left");
    assert_eq!(left["result"]["structuredContent"]["sequence"], 2);

    let ahead = call(
        &mut stdin,
        &mut stdout,
        12,
        "office_collaboration_events",
        serde_json::json!({
            "store": replica_text,
            "afterSequence": 3,
            "limit": 1
        }),
        TIMEOUT,
    )
    .await;
    assert_eq!(ahead["result"]["isError"], true, "{ahead}");
    assert_eq!(
        ahead["result"]["structuredContent"]["code"],
        "office.collaboration.sequence_ahead"
    );

    drop(stdin);
    let status = tokio::time::timeout(TIMEOUT, child.wait())
        .await
        .unwrap()
        .unwrap();
    assert!(status.success());
    let mut diagnostics = Vec::new();
    stderr.read_to_end(&mut diagnostics).await.unwrap();
    assert!(
        diagnostics.is_empty(),
        "{}",
        String::from_utf8_lossy(&diagnostics)
    );
    assert!(replica.is_dir());
    assert!(!provider.exists());
}

async fn call(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut BufReader<tokio::process::ChildStdout>,
    id: u32,
    name: &str,
    arguments: serde_json::Value,
    timeout: Duration,
) -> serde_json::Value {
    request(
        stdin,
        stdout,
        serde_json::json!({
            "jsonrpc": "2.0",
            "id": id,
            "method": "tools/call",
            "params": { "name": name, "arguments": arguments }
        }),
        timeout,
    )
    .await
}

async fn request(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut BufReader<tokio::process::ChildStdout>,
    value: serde_json::Value,
    timeout: Duration,
) -> serde_json::Value {
    let mut encoded = serde_json::to_vec(&value).unwrap();
    encoded.push(b'\n');
    stdin.write_all(&encoded).await.unwrap();
    stdin.flush().await.unwrap();
    let mut line = String::new();
    let bytes = tokio::time::timeout(timeout, stdout.read_line(&mut line))
        .await
        .unwrap()
        .unwrap();
    assert!(bytes > 0, "native Office MCP closed before responding");
    serde_json::from_str(&line).unwrap()
}
