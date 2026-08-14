use std::fs;
use std::io::{BufRead, BufReader, Read};
use std::path::Path;
use std::process::{Command, Output, Stdio};

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use yrs::{Doc, ReadTxn, StateVector, Text, Transact};

const YJS_MARKDOWN_UPDATE_BASE64: &str = "AQey8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtbWFya2Rvd24oARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhtYXJrZG93bigBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjQyNDI0Mjpicm93c2VyLWZpeHR1cmUEARphM3Mub2ZmaWNlLm1hcmtkb3duLnNvdXJjZRUjIFNoYXJlZAoKWWpzIHRvIFlycy4A";

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
            "5000",
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

fn read_jsonl(reader: &mut impl BufRead) -> serde_json::Value {
    let mut line = String::new();
    assert!(
        reader.read_line(&mut line).unwrap() > 0,
        "watch ended early"
    );
    serde_json::from_str(&line).unwrap()
}
