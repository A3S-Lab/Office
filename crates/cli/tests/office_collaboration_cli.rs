use std::fs;
use std::io::{BufRead, BufReader, Read, Write};
use std::path::Path;
use std::process::{Command, Output, Stdio};

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use yrs::sync::{Awareness, AwarenessUpdate};
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{
    Any, ClientID, Doc, GetString, Map, Out, ReadTxn, StateVector, Text, Transact, Update, Xml,
    XmlFragment, XmlOut,
};

const YJS_MARKDOWN_UPDATE_BASE64: &str = "AQey8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtbWFya2Rvd24oARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhtYXJrZG93bigBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjQyNDI0Mjpicm93c2VyLWZpeHR1cmUEARphM3Mub2ZmaWNlLm1hcmtkb3duLnNvdXJjZRUjIFNoYXJlZAoKWWpzIHRvIFlycy4A";
const YJS_DOCUMENT_UPDATE_BASE64: &str = "AQ+y8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtZG9jdW1lbnQoARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhkb2N1bWVudCgBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjQyNDI0Mjpicm93c2VyLWZpeHR1cmUHARthM3Mub2ZmaWNlLmRvY3VtZW50LmNvbnRlbnQDD2RvY3VtZW50U2VjdGlvbgcAsvIZBgMJcGFyYWdyYXBoBwCy8hkHBgQAsvIZCBBIZWxsbyDwn5iAIHdvcmxkKACy8hkHC3BhcmFncmFwaElkAXcIMDAwMDAwMDEoALLyGQcGdGV4dElkAXcIMDAwMDAwMDIoALLyGQYCaWQBdxJkb2N1bWVudC1zZWN0aW9uLTEoARthM3Mub2ZmaWNlLmRvY3VtZW50Lm9wdGlvbnMJcGFnZUNvbG9yAXcHI0Y4RkFGQygBG2Ezcy5vZmZpY2UuZG9jdW1lbnQub3B0aW9ucwx0cmFja0NoYW5nZXMBeAA=";

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

fn execute(args: &[&str]) -> Output {
    Command::new(binary()).args(args).output().unwrap()
}

fn run(args: &[&str]) -> serde_json::Value {
    let output = execute(args);
    assert!(output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

fn run_failure(args: &[&str]) -> serde_json::Value {
    let output = execute(args);
    assert!(!output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

fn join(replica: &Path, update: &Path) -> serde_json::Value {
    run(&[
        "collab",
        "join",
        replica.to_str().unwrap(),
        "--artifact-id",
        "fixture-markdown",
        "--kind",
        "markdown",
        "--actor-id",
        "coding-agent-7",
        "--actor-kind",
        "agent",
        "--mode",
        "edit",
        "--operation-id",
        "join-browser-1",
        "--input",
        update.to_str().unwrap(),
        "--client-id",
        "900007",
        "--json",
    ])
}

#[test]
fn cli_comment_mode_mutations_round_trip_through_the_native_projection() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("reviewer.replica");
    let input = temp.path().join("document.update");
    fs::write(&input, STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap()).unwrap();

    let joined = run(&[
        "collab",
        "join",
        replica.to_str().unwrap(),
        "--artifact-id",
        "fixture-document",
        "--kind",
        "document",
        "--actor-id",
        "review-agent",
        "--actor-kind",
        "agent",
        "--mode",
        "comment",
        "--operation-id",
        "join-comment-document-1",
        "--input",
        input.to_str().unwrap(),
        "--client-id",
        "900108",
        "--json",
    ]);
    assert_eq!(joined["data"]["replica"]["manifest"]["mode"], "comment");

    let created = run(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        r#"{"type":"document-comment-create","commentId":"comment-cli-1","paragraphId":"00000001","expectedTextId":"00000002","startUtf16":6,"endUtf16":8,"expectedText":"😀","author":"Review Agent","createdAt":"2026-08-17T00:00:00.000Z","text":"Check this symbol."}"#,
        "--actor-id",
        "review-agent",
        "--operation-id",
        "comment-cli-create-1",
        "--artifact-id",
        "fixture-document",
        "--kind",
        "document",
        "--mode",
        "comment",
        "--json",
    ]);
    assert_eq!(
        created["data"]["mutation"]["type"],
        "document-comment-create"
    );
    assert_eq!(created["data"]["sequence"], 1);

    run(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        r#"{"type":"document-comment-reply","commentId":"comment-cli-1","replyId":"reply-cli-1","author":"Review Agent","createdAt":"2026-08-17T00:01:00.000Z","text":"Confirmed."}"#,
        "--actor-id",
        "review-agent",
        "--operation-id",
        "comment-cli-reply-1",
        "--artifact-id",
        "fixture-document",
        "--kind",
        "document",
        "--mode",
        "comment",
        "--json",
    ]);
    run(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        r#"{"type":"document-comment-set-resolved","commentId":"comment-cli-1","resolved":true}"#,
        "--actor-id",
        "review-agent",
        "--operation-id",
        "comment-cli-resolve-1",
        "--artifact-id",
        "fixture-document",
        "--kind",
        "document",
        "--mode",
        "comment",
        "--json",
    ]);

    let projected = run(&["collab", "read", replica.to_str().unwrap(), "--json"]);
    let comment = &projected["data"]["content"]["comments"][0];
    assert_eq!(comment["id"], "comment-cli-1");
    assert_eq!(comment["actorId"], "review-agent");
    assert_eq!(comment["resolved"], true);
    assert_eq!(comment["replies"][0]["id"], "reply-cli-1");
    assert_eq!(comment["anchors"][0]["startUtf16"], 6);
    assert_eq!(comment["anchors"][0]["endUtf16"], 8);
    assert_eq!(comment["anchors"][0]["text"], "😀");

    let forbidden = run_failure(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        r#"{"type":"document-replace-text","search":"Hello","replacement":"Changed","expectedMatches":1}"#,
        "--actor-id",
        "review-agent",
        "--operation-id",
        "comment-cli-forbidden-1",
        "--artifact-id",
        "fixture-document",
        "--kind",
        "document",
        "--mode",
        "comment",
        "--json",
    ]);
    assert_eq!(
        forbidden["error"]["code"],
        "office.collaboration.mutation_forbidden"
    );
}

#[test]
fn cli_joins_yjs_replica_inspects_and_exports_state_vector_diff() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("agent.replica");
    let input = temp.path().join("browser.update");
    let output = temp.path().join("agent.diff");
    fs::write(&input, STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap()).unwrap();

    let joined = join(&replica, &input);
    assert_eq!(joined["data"]["action"], "joined");
    assert_eq!(
        joined["data"]["replica"]["metadata"]["artifactId"],
        "fixture-markdown"
    );
    assert_eq!(
        joined["data"]["replica"]["manifest"]["actorId"],
        "coding-agent-7"
    );
    assert_eq!(joined["data"]["replica"]["manifest"]["clientId"], 900_007);

    let inspected = run(&["collab", "inspect", replica.to_str().unwrap(), "--json"]);
    let state_vector = inspected["data"]["stateVector"]
        .as_str()
        .unwrap()
        .to_owned();
    assert_eq!(inspected["data"]["bootstrapValid"], true);
    assert_eq!(inspected["data"]["currentSequence"], 0);

    let projected = run(&["collab", "read", replica.to_str().unwrap(), "--json"]);
    assert_eq!(projected["data"]["content"]["kind"], "markdown");
    assert_eq!(
        projected["data"]["content"]["source"],
        "# Shared\n\nYjs to Yrs."
    );
    assert!(projected["data"]["stateVector"].is_string());

    let no_diff = run(&[
        "collab",
        "diff",
        replica.to_str().unwrap(),
        "--state-vector",
        &state_vector,
        "--json",
    ]);
    assert_eq!(no_diff["data"]["updateBase64"], "AAA=");

    let diff = run(&[
        "collab",
        "diff",
        replica.to_str().unwrap(),
        "--output",
        output.to_str().unwrap(),
        "--json",
    ]);
    assert_eq!(diff["data"]["output"], output.to_str().unwrap());
    assert!(fs::metadata(&output).unwrap().len() > 2);

    let step1 = run(&["collab", "sync-step1", replica.to_str().unwrap(), "--json"]);
    assert_eq!(step1["data"]["action"], "sync-step1");
    assert!(step1["data"]["messageBase64"].as_str().unwrap().len() > 4);

    let encoded = run(&[
        "collab",
        "encode-update",
        "--input",
        input.to_str().unwrap(),
        "--json",
    ]);
    assert_eq!(encoded["data"]["action"], "encode-update");
    assert!(encoded["data"]["messageBytes"].as_u64().unwrap() > 2);

    let step1_message = temp.path().join("step1.message");
    let response_message = temp.path().join("step2.message");
    run(&[
        "collab",
        "sync-step1",
        replica.to_str().unwrap(),
        "--output",
        step1_message.to_str().unwrap(),
        "--json",
    ]);
    let handled = run(&[
        "collab",
        "handle-message",
        replica.to_str().unwrap(),
        "--input",
        step1_message.to_str().unwrap(),
        "--output",
        response_message.to_str().unwrap(),
        "--json",
    ]);
    assert_eq!(handled["data"]["kind"], "sync-step1");
    assert!(fs::metadata(response_message).unwrap().len() > 2);
}

#[test]
fn cli_requires_explicit_mutation_identity_and_reports_stale_state() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("agent.replica");
    let input = temp.path().join("browser.update");
    let empty_update = temp.path().join("empty.update");
    fs::write(&input, STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap()).unwrap();
    fs::write(&empty_update, [0_u8, 0]).unwrap();
    join(&replica, &input);

    let missing_identity = run_failure(&[
        "collab",
        "apply",
        replica.to_str().unwrap(),
        "--input",
        empty_update.to_str().unwrap(),
        "--json",
    ]);
    assert_eq!(
        missing_identity["error"]["code"],
        "office.collaboration.usage"
    );

    let stale = run_failure(&[
        "collab",
        "apply",
        replica.to_str().unwrap(),
        "--input",
        empty_update.to_str().unwrap(),
        "--actor-id",
        "coding-agent-7",
        "--operation-id",
        "sync-stale-1",
        "--artifact-id",
        "fixture-markdown",
        "--kind",
        "markdown",
        "--mode",
        "edit",
        "--if-state-vector",
        "AA==",
        "--json",
    ]);
    assert_eq!(stale["error"]["code"], "office.collaboration.stale_state");
}

#[test]
fn cli_applies_typed_document_text_option_and_paragraph_mutations() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("document-agent.replica");
    let input = temp.path().join("browser-document.update");
    fs::write(&input, STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap()).unwrap();
    run(&[
        "collab",
        "join",
        replica.to_str().unwrap(),
        "--artifact-id",
        "fixture-document",
        "--kind",
        "document",
        "--actor-id",
        "coding-agent-8",
        "--actor-kind",
        "agent",
        "--mode",
        "edit",
        "--operation-id",
        "join-browser-document-1",
        "--input",
        input.to_str().unwrap(),
        "--client-id",
        "900008",
        "--json",
    ]);

    let replaced = run(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        r#"{"type":"document-replace-text","search":"😀","replacement":"🦀","expectedMatches":1}"#,
        "--actor-id",
        "coding-agent-8",
        "--operation-id",
        "document-text-1",
        "--artifact-id",
        "fixture-document",
        "--kind",
        "document",
        "--mode",
        "edit",
        "--json",
    ]);
    assert_eq!(replaced["data"]["action"], "mutated");
    assert_eq!(replaced["data"]["sequence"], 1);
    assert_eq!(
        replaced["data"]["mutation"]["type"],
        "document-replace-text"
    );

    let page_color = run(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        r##"{"type":"document-set-page-color","pageColor":"#101828"}"##,
        "--actor-id",
        "coding-agent-8",
        "--operation-id",
        "document-page-color-1",
        "--artifact-id",
        "fixture-document",
        "--kind",
        "document",
        "--mode",
        "edit",
        "--json",
    ]);
    assert_eq!(page_color["data"]["sequence"], 2);

    let inserted = run(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        r#"{"type":"document-insert-paragraph","anchorParagraphId":"00000001","position":"after","paragraphId":"00000030","textId":"00000031","text":"CLI paragraph"}"#,
        "--actor-id",
        "coding-agent-8",
        "--operation-id",
        "document-insert-paragraph-1",
        "--artifact-id",
        "fixture-document",
        "--kind",
        "document",
        "--mode",
        "edit",
        "--json",
    ]);
    assert_eq!(inserted["data"]["sequence"], 3);

    let exported = run(&["collab", "diff", replica.to_str().unwrap(), "--json"]);
    let peer = Doc::with_client_id(700_008);
    peer.transact_mut()
        .apply_update(
            Update::decode_v1(
                &STANDARD
                    .decode(exported["data"]["updateBase64"].as_str().unwrap())
                    .unwrap(),
            )
            .unwrap(),
        )
        .unwrap();
    let fragment = peer.get_or_insert_xml_fragment("a3s.office.document.content");
    let options = peer.get_or_insert_map("a3s.office.document.options");
    let transaction = peer.transact();
    let paragraphs = fragment
        .successors(&transaction)
        .filter_map(|node| match node {
            XmlOut::Element(element) if element.tag().as_ref() == "paragraph" => Some((
                element
                    .get_attribute(&transaction, "paragraphId")
                    .and_then(|value| match value {
                        Out::Any(Any::String(value)) => Some(value.to_string()),
                        _ => None,
                    }),
                element
                    .get_attribute(&transaction, "textId")
                    .and_then(|value| match value {
                        Out::Any(Any::String(value)) => Some(value.to_string()),
                        _ => None,
                    }),
                element
                    .children(&transaction)
                    .filter_map(|child| match child {
                        XmlOut::Text(text) => Some(text.get_string(&transaction)),
                        _ => None,
                    })
                    .collect::<String>(),
            )),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(
        paragraphs,
        vec![
            (
                Some("00000001".to_owned()),
                Some("00000003".to_owned()),
                "Hello 🦀 world".to_owned(),
            ),
            (
                Some("00000030".to_owned()),
                Some("00000031".to_owned()),
                "CLI paragraph".to_owned(),
            ),
        ]
    );
    assert!(matches!(
        options.get(&transaction, "pageColor"),
        Some(Out::Any(Any::String(value))) if value.as_ref() == "#101828"
    ));
}

#[test]
fn cli_operation_replay_survives_checkpoint_and_leave() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("agent.replica");
    let input = temp.path().join("browser.update");
    let empty_update = temp.path().join("empty.update");
    fs::write(&input, STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap()).unwrap();
    fs::write(&empty_update, [0_u8, 0]).unwrap();
    join(&replica, &input);

    let mutation = [
        "collab",
        "apply",
        replica.to_str().unwrap(),
        "--input",
        empty_update.to_str().unwrap(),
        "--actor-id",
        "coding-agent-7",
        "--operation-id",
        "sync-noop-1",
        "--artifact-id",
        "fixture-markdown",
        "--kind",
        "markdown",
        "--mode",
        "edit",
        "--json",
    ];
    let first = run(&mutation);
    let replay = run(&mutation);
    assert_eq!(first["data"]["duplicate"], false);
    assert_eq!(first["data"]["stateChanged"], false);
    assert_eq!(replay["data"]["duplicate"], true);

    let left = run(&[
        "collab",
        "leave",
        replica.to_str().unwrap(),
        "--actor-id",
        "coding-agent-7",
        "--operation-id",
        "leave-1",
        "--artifact-id",
        "fixture-markdown",
        "--kind",
        "markdown",
        "--mode",
        "edit",
        "--json",
    ]);
    assert_eq!(left["data"]["action"], "left");
    assert!(replica.is_dir());
    assert_eq!(
        run(&["collab", "inspect", replica.to_str().unwrap(), "--json"])["data"]["metadata"]
            ["initialized"],
        true
    );
}

#[test]
fn cli_watch_streams_resumable_jsonl_updates_between_agent_processes() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("agent.replica");
    let initial = temp.path().join("browser.update");
    let live = temp.path().join("live-agent.update");
    fs::write(
        &initial,
        STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap(),
    )
    .unwrap();
    join(&replica, &initial);

    let peer = Doc::with_client_id(800_008);
    let notes = peer.get_or_insert_text("agent.notes");
    notes.insert(&mut peer.transact_mut(), 0, "live change");
    let update = peer
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    fs::write(&live, &update).unwrap();

    let mut watcher = Command::new(binary())
        .args([
            "collab",
            "watch",
            replica.to_str().unwrap(),
            "--after-sequence",
            "0",
            "--poll-ms",
            "50",
            "--timeout-ms",
            "30000",
            "--max-events",
            "1",
            "--include-updates",
            "--json",
        ])
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    let mut stdout = BufReader::new(watcher.stdout.take().unwrap());
    let mut stderr = watcher.stderr.take().unwrap();

    let ready = read_jsonl(&mut stdout);
    assert_eq!(ready["type"], "ready");
    assert_eq!(ready["startingSequence"], 0);
    assert_eq!(ready["cursorSequence"], 0);
    assert_eq!(ready["artifactId"], "fixture-markdown");

    let applied = run(&[
        "collab",
        "apply",
        replica.to_str().unwrap(),
        "--input",
        live.to_str().unwrap(),
        "--actor-id",
        "coding-agent-7",
        "--operation-id",
        "live-agent-1",
        "--artifact-id",
        "fixture-markdown",
        "--kind",
        "markdown",
        "--mode",
        "edit",
        "--json",
    ]);
    assert_eq!(applied["data"]["sequence"], 1);

    let event = read_jsonl(&mut stdout);
    assert_eq!(event["type"], "update");
    assert_eq!(event["sequence"], 1);
    assert_eq!(event["cursorSequence"], 1);
    assert_eq!(event["operationId"], "live-agent-1");
    assert_eq!(event["actorId"], "coding-agent-7");
    assert_eq!(event["updateBase64"], STANDARD.encode(&update));

    let complete = read_jsonl(&mut stdout);
    assert_eq!(complete["type"], "complete");
    assert_eq!(complete["reason"], "max-events");
    assert_eq!(complete["cursorSequence"], 1);
    assert_eq!(complete["eventCount"], 1);

    let status = watcher.wait().unwrap();
    let mut error_text = String::new();
    stderr.read_to_string(&mut error_text).unwrap();
    assert!(status.success(), "watch failed: {error_text}");
}

#[test]
fn cli_session_bridges_live_browser_envelopes_and_external_agent_updates() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("agent.replica");
    let initial = temp.path().join("browser.update");
    let local = temp.path().join("local-agent-mutation.json");
    fs::write(
        &initial,
        STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap(),
    )
    .unwrap();
    join(&replica, &initial);

    let mut child = Command::new(binary())
        .args([
            "collab",
            "session",
            replica.to_str().unwrap(),
            "--poll-ms",
            "50",
            "--timeout-ms",
            "30000",
            "--json",
        ])
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    let mut stdin = child.stdin.take().unwrap();
    let mut stdout = BufReader::new(child.stdout.take().unwrap());
    let mut stderr = child.stderr.take().unwrap();

    let ready = read_jsonl(&mut stdout);
    assert_eq!(ready["type"], "ready");
    assert_eq!(ready["artifactId"], "fixture-markdown");
    assert_eq!(ready["clientId"], 900_007);
    let initial_sync = read_jsonl(&mut stdout);
    assert_eq!(initial_sync["type"], "outbound");
    assert_eq!(initial_sync["reason"], "initial-connect");
    assert_eq!(initial_sync["message"]["type"], "sync-step-1");
    assert!(initial_sync["message"].get("origin").is_none());

    fs::write(
        &local,
        serde_json::to_vec(&serde_json::json!({
            "type": "markdown-replace",
            "markdown": "# Agent live 😀"
        }))
        .unwrap(),
    )
    .unwrap();
    let applied = run(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation-input",
        local.to_str().unwrap(),
        "--actor-id",
        "coding-agent-7",
        "--operation-id",
        "agent-live-1",
        "--artifact-id",
        "fixture-markdown",
        "--kind",
        "markdown",
        "--mode",
        "edit",
        "--json",
    ]);
    assert_eq!(applied["data"]["sequence"], 1);

    let outbound = read_jsonl(&mut stdout);
    assert_eq!(outbound["type"], "outbound");
    assert_eq!(outbound["reason"], "durable-update");
    assert_eq!(outbound["message"]["type"], "update");
    assert_eq!(outbound["message"]["origin"]["operationId"], "agent-live-1");
    let typed_peer = Doc::with_client_id(700_007);
    for encoded in [
        YJS_MARKDOWN_UPDATE_BASE64,
        outbound["message"]["payloadBase64"].as_str().unwrap(),
    ] {
        typed_peer
            .transact_mut()
            .apply_update(Update::decode_v1(&STANDARD.decode(encoded).unwrap()).unwrap())
            .unwrap();
    }
    let transaction = typed_peer.transact();
    assert_eq!(
        transaction
            .get_text("a3s.office.markdown.source")
            .unwrap()
            .get_string(&transaction),
        "# Agent live 😀"
    );

    write_jsonl(&mut stdin, &serde_json::json!({ "type": "reconnect" }));
    let reconnect = read_jsonl(&mut stdout);
    assert_eq!(reconnect["type"], "outbound");
    assert_eq!(reconnect["reason"], "reconnect");
    assert_eq!(reconnect["message"]["type"], "sync-step-1");

    write_jsonl(
        &mut stdin,
        &serde_json::json!({
            "type": "receive",
            "message": {
                "protocol": "a3s.office.collaboration",
                "version": 1,
                "artifactId": "fixture-markdown",
                "artifactKind": "markdown",
                "namespace": "a3s.office",
                "senderClientId": 424242,
                "type": "sync-step-1",
                "payloadBase64": "AA=="
            }
        }),
    );
    let received_step1 = read_jsonl(&mut stdout);
    assert_eq!(received_step1["type"], "received");
    assert_eq!(received_step1["messageType"], "sync-step-1");
    let response = read_jsonl(&mut stdout);
    assert_eq!(response["type"], "outbound");
    assert_eq!(response["reason"], "peer-sync-step1");
    assert_eq!(response["message"]["type"], "sync-step-2");
    assert!(response["message"].get("origin").is_none());

    let browser_peer = Doc::with_client_id(600_006);
    browser_peer.get_or_insert_text("browser.live").insert(
        &mut browser_peer.transact_mut(),
        0,
        "browser update",
    );
    let browser_update = browser_peer
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    write_jsonl(
        &mut stdin,
        &serde_json::json!({
            "type": "receive",
            "operationId": "browser-delivery-1",
            "message": {
                "protocol": "a3s.office.collaboration",
                "version": 1,
                "artifactId": "fixture-markdown",
                "artifactKind": "markdown",
                "namespace": "a3s.office",
                "senderClientId": 600006,
                "type": "update",
                "payloadBase64": STANDARD.encode(&browser_update),
                "origin": {
                    "protocol": "a3s.office.collaboration",
                    "kind": "editor",
                    "actorId": "browser-user",
                    "operationId": "browser-edit-1"
                }
            }
        }),
    );
    let received_update = read_jsonl(&mut stdout);
    assert_eq!(
        received_update["type"], "received",
        "unexpected session response: {received_update}"
    );
    assert_eq!(received_update["messageType"], "update");
    assert_eq!(received_update["operationId"], "browser-delivery-1");
    assert_eq!(received_update["sequence"], 2);
    assert_eq!(received_update["duplicate"], false);

    write_jsonl(&mut stdin, &serde_json::json!({ "type": "close" }));
    let complete = read_jsonl(&mut stdout);
    assert_eq!(complete["type"], "complete");
    assert_eq!(complete["reason"], "host-close");
    assert_eq!(complete["cursorSequence"], 2);
    drop(stdin);

    let status = child.wait().unwrap();
    let mut error_text = String::new();
    stderr.read_to_string(&mut error_text).unwrap();
    assert!(status.success(), "session failed: {error_text}");

    let inspected = run(&["collab", "inspect", replica.to_str().unwrap(), "--json"]);
    assert_eq!(inspected["data"]["currentSequence"], 2);
}

#[test]
fn cli_session_bridges_ephemeral_native_presence_without_touching_the_replica() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("agent.replica");
    let initial = temp.path().join("browser.update");
    fs::write(
        &initial,
        STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap(),
    )
    .unwrap();
    join(&replica, &initial);
    let before = run(&["collab", "inspect", replica.to_str().unwrap(), "--json"]);

    let mut child = Command::new(binary())
        .args([
            "collab",
            "session",
            replica.to_str().unwrap(),
            "--poll-ms",
            "50",
            "--timeout-ms",
            "30000",
            "--actor-name",
            "A3S Agent",
            "--actor-color",
            "#2563eb",
            "--json",
        ])
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    let mut stdin = child.stdin.take().unwrap();
    let mut stdout = BufReader::new(child.stdout.take().unwrap());
    let mut stderr = child.stderr.take().unwrap();

    let ready = read_jsonl(&mut stdout);
    assert_eq!(ready["type"], "ready");
    assert_eq!(ready["presenceEnabled"], true);
    assert_eq!(ready["actorName"], "A3S Agent");
    assert_eq!(ready["replicaClientId"], 900_007);
    let presence_client_id = ready["clientId"].as_u64().unwrap();
    assert_ne!(presence_client_id, 900_007);
    assert_eq!(read_jsonl(&mut stdout)["type"], "outbound");
    let initial_presence = read_jsonl(&mut stdout);
    assert_eq!(initial_presence["type"], "outbound-awareness");
    assert_eq!(initial_presence["reason"], "initial-connect");
    let decoded = AwarenessUpdate::decode_v1(
        &STANDARD
            .decode(
                initial_presence["message"]["payloadBase64"]
                    .as_str()
                    .unwrap(),
            )
            .unwrap(),
    )
    .unwrap();
    let local = decoded
        .clients
        .get(&ClientID::new(presence_client_id))
        .unwrap();
    let local: serde_json::Value = serde_json::from_str(local.json.as_ref()).unwrap();
    assert_eq!(local["a3sOffice"]["actor"]["id"], "coding-agent-7");
    assert_eq!(local["a3sOffice"]["actor"]["name"], "A3S Agent");
    assert_eq!(local["a3sOffice"]["actor"]["kind"], "agent");

    write_jsonl(
        &mut stdin,
        &serde_json::json!({
            "type": "set-presence",
            "activity": "idle",
            "location": {
                "kind": "markdown",
                "anchor": 2,
                "head": 8,
                "surface": "source"
            }
        }),
    );
    let local_update = read_jsonl(&mut stdout);
    assert_eq!(local_update["type"], "outbound-awareness");
    assert_eq!(local_update["reason"], "local-presence");
    let local_snapshot = read_jsonl(&mut stdout);
    assert_eq!(local_snapshot["type"], "presence");
    assert_eq!(
        local_snapshot["snapshot"]["participants"][0]["activity"],
        "idle"
    );
    assert_eq!(
        local_snapshot["snapshot"]["participants"][0]["location"]["head"],
        8
    );

    let remote_client_id = 424_242;
    let mut remote = Awareness::new(Doc::with_client_id(remote_client_id));
    remote
        .set_local_state(serde_json::json!({
            "a3sOffice": {
                "protocol": "a3s.office.collaboration",
                "version": 1,
                "artifactId": "fixture-markdown",
                "artifactKind": "markdown",
                "namespace": "a3s.office",
                "presenceId": "browser-user:1",
                "actor": {
                    "id": "browser-user",
                    "name": "Browser User",
                    "kind": "human"
                },
                "mode": "comment",
                "activity": "active",
                "location": {
                    "kind": "markdown",
                    "anchor": 1,
                    "head": 4,
                    "surface": "visual"
                }
            }
        }))
        .unwrap();
    let remote_payload = remote.update().unwrap().encode_v1();
    write_jsonl(
        &mut stdin,
        &serde_json::json!({
            "type": "receive-awareness",
            "message": {
                "protocol": "a3s.office.collaboration",
                "version": 1,
                "artifactId": "fixture-markdown",
                "artifactKind": "markdown",
                "namespace": "a3s.office",
                "senderClientId": remote_client_id,
                "payloadBase64": STANDARD.encode(&remote_payload)
            }
        }),
    );
    let remote_snapshot = read_jsonl(&mut stdout);
    assert_eq!(
        remote_snapshot["type"], "presence",
        "unexpected remote Awareness response: {remote_snapshot}"
    );
    assert_eq!(remote_snapshot["changed"], true);
    assert_eq!(
        remote_snapshot["snapshot"]["participants"]
            .as_array()
            .unwrap()
            .len(),
        2
    );
    assert!(remote_snapshot["snapshot"]["participants"]
        .as_array()
        .unwrap()
        .iter()
        .any(|participant| participant["actor"]["name"] == "Browser User"));

    write_jsonl(
        &mut stdin,
        &serde_json::json!({
            "type": "peer-left",
            "senderClientId": remote_client_id
        }),
    );
    let peer_left = read_jsonl(&mut stdout);
    assert_eq!(peer_left["type"], "presence");
    assert_eq!(peer_left["reason"], "peer-left");
    assert_eq!(
        peer_left["snapshot"]["participants"]
            .as_array()
            .unwrap()
            .len(),
        1
    );

    write_jsonl(&mut stdin, &serde_json::json!({ "type": "reconnect" }));
    assert_eq!(read_jsonl(&mut stdout)["type"], "outbound");
    let reconnect_presence = read_jsonl(&mut stdout);
    assert_eq!(reconnect_presence["type"], "outbound-awareness");
    assert_eq!(reconnect_presence["reason"], "reconnect");

    write_jsonl(&mut stdin, &serde_json::json!({ "type": "close" }));
    let removal = read_jsonl(&mut stdout);
    assert_eq!(removal["type"], "outbound-awareness");
    assert_eq!(removal["reason"], "disconnect");
    let decoded = AwarenessUpdate::decode_v1(
        &STANDARD
            .decode(removal["message"]["payloadBase64"].as_str().unwrap())
            .unwrap(),
    )
    .unwrap();
    assert_eq!(
        decoded
            .clients
            .get(&ClientID::new(presence_client_id))
            .unwrap()
            .json
            .as_ref(),
        "null"
    );
    let complete = read_jsonl(&mut stdout);
    assert_eq!(complete["type"], "complete");
    assert_eq!(complete["reason"], "host-close");
    drop(stdin);

    let status = child.wait().unwrap();
    let mut error_text = String::new();
    stderr.read_to_string(&mut error_text).unwrap();
    assert!(status.success(), "presence session failed: {error_text}");

    let after = run(&["collab", "inspect", replica.to_str().unwrap(), "--json"]);
    assert_eq!(
        after["data"]["currentSequence"],
        before["data"]["currentSequence"]
    );
    assert_eq!(
        after["data"]["operationCount"],
        before["data"]["operationCount"]
    );
    assert_eq!(
        after["data"]["documentStateSha256"],
        before["data"]["documentStateSha256"]
    );
}

#[test]
fn cli_session_reports_runtime_failures_as_one_jsonl_record() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("agent.replica");
    let initial = temp.path().join("browser.update");
    fs::write(
        &initial,
        STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap(),
    )
    .unwrap();
    join(&replica, &initial);

    let mut child = Command::new(binary())
        .args([
            "collab",
            "session",
            replica.to_str().unwrap(),
            "--timeout-ms",
            "30000",
            "--json",
        ])
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    let mut stdin = child.stdin.take().unwrap();
    let mut stdout = BufReader::new(child.stdout.take().unwrap());
    let mut stderr = child.stderr.take().unwrap();

    assert_eq!(read_jsonl(&mut stdout)["type"], "ready");
    assert_eq!(read_jsonl(&mut stdout)["type"], "outbound");
    write_jsonl(&mut stdin, &serde_json::json!({ "type": "receive" }));
    let error = read_jsonl(&mut stdout);
    assert_eq!(error["type"], "error");
    assert_eq!(
        error["error"]["code"],
        "office.collaboration.session_input_invalid"
    );
    drop(stdin);

    let status = child.wait().unwrap();
    assert_eq!(status.code(), Some(1));
    let mut trailing_output = String::new();
    stdout.read_to_string(&mut trailing_output).unwrap();
    assert!(trailing_output.is_empty(), "{trailing_output}");
    let mut error_text = String::new();
    stderr.read_to_string(&mut error_text).unwrap();
    assert!(error_text.is_empty(), "{error_text}");
}

fn read_jsonl(reader: &mut impl BufRead) -> serde_json::Value {
    let mut line = String::new();
    assert!(
        reader.read_line(&mut line).unwrap() > 0,
        "watch ended early"
    );
    serde_json::from_str(&line).unwrap()
}

fn write_jsonl(writer: &mut impl Write, value: &serde_json::Value) {
    serde_json::to_writer(&mut *writer, value).unwrap();
    writer.write_all(b"\n").unwrap();
    writer.flush().unwrap();
}
