use std::path::Path;
use std::process::{Command, Output};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

fn execute(provider: &Path, args: &[&str]) -> Output {
    Command::new(binary())
        .args(args)
        .env("A3S_OFFICECLI_EXECUTABLE", provider)
        .output()
        .unwrap()
}

fn run(provider: &Path, args: &[&str]) -> serde_json::Value {
    let output = execute(provider, args);
    assert!(output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

fn run_failure(provider: &Path, args: &[&str]) -> serde_json::Value {
    let output = execute(provider, args);
    assert!(!output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

fn create(provider: &Path, document: &Path) {
    run(provider, &["create", document.to_str().unwrap(), "--json"]);
}

#[test]
fn native_dump_artifacts_replay_through_atomic_batch_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let source = temp.path().join("source.docx");
    let target = temp.path().join("target.docx");
    let artifact = temp.path().join("source.replay.json");
    let target_artifact = temp.path().join("target.replay.json");

    create(&provider, &source);
    run(
        &provider,
        &[
            "set",
            source.to_str().unwrap(),
            "/body/p[1]",
            "--text",
            "Replay me",
            "--json",
        ],
    );
    run(
        &provider,
        &[
            "add",
            source.to_str().unwrap(),
            "/body",
            "--type",
            "paragraph",
            "--text",
            "Exactly",
            "--json",
        ],
    );

    let dumped = run(
        &provider,
        &[
            "dump",
            source.to_str().unwrap(),
            "--output",
            artifact.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(dumped["data"]["format"], "a3s.office.native-replay");
    assert_eq!(dumped["data"]["artifactSchemaVersion"], 1);
    assert_eq!(dumped["data"]["documentKind"], "word");
    let source_result = dumped["data"]["resultSha256"].clone();
    let artifact_json: serde_json::Value =
        serde_json::from_slice(&std::fs::read(&artifact).unwrap()).unwrap();
    assert_eq!(artifact_json["resultSha256"], source_result);

    let no_clobber = run_failure(
        &provider,
        &[
            "dump",
            source.to_str().unwrap(),
            "--output",
            artifact.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(no_clobber["error"]["code"], "use.office.dump_output_exists");

    create(&provider, &target);
    let replayed = run(
        &provider,
        &[
            "batch",
            target.to_str().unwrap(),
            "--input",
            artifact.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(replayed["data"]["replay"], true);
    assert_eq!(replayed["data"]["atomic"], true);
    assert_eq!(replayed["data"]["inputMutations"], 2);

    let target_dump = run(
        &provider,
        &[
            "dump",
            target.to_str().unwrap(),
            "--output",
            target_artifact.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(target_dump["data"]["resultSha256"], source_result);

    let target_before = std::fs::read(&target).unwrap();
    let wrong_base = run_failure(
        &provider,
        &[
            "batch",
            target.to_str().unwrap(),
            "--input",
            artifact.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(
        wrong_base["error"]["code"],
        "use.office.replay_base_mismatch"
    );
    assert_eq!(std::fs::read(&target).unwrap(), target_before);
}

#[test]
fn native_dump_stdout_is_a_batch_artifact_and_subtrees_fail_closed() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("blank.xlsx");
    create(&provider, &document);

    let output = execute(&provider, &["dump", document.to_str().unwrap()]);
    assert!(output.status.success(), "{output:?}");
    let artifact: serde_json::Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(artifact["format"], "a3s.office.native-replay");
    assert_eq!(artifact["mutations"], serde_json::json!([]));

    let unsupported = run_failure(
        &provider,
        &["dump", document.to_str().unwrap(), "/Sheet1", "--json"],
    );
    assert_eq!(
        unsupported["error"]["code"],
        "use.office.dump_scope_unsupported"
    );
}
