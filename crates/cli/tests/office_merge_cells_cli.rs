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
fn native_cli_merges_reads_and_precisely_unmerges_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("merged.xlsx");
    create(&provider, &document);

    let merged = run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/B2:A1",
            "--text",
            "Quarter",
            "--bold",
            "true",
            "--merge-cells",
            "true",
            "--json",
        ],
    );
    assert_eq!(merged["data"]["operation"], "set-and-merge-cells");
    assert_eq!(merged["data"]["changed"], true);
    assert_eq!(merged["data"]["node"]["format"]["merge"], "true");

    let anchor = run(
        &provider,
        &["get", document.to_str().unwrap(), "/Sheet1/A1", "--json"],
    );
    assert_eq!(anchor["data"]["node"]["text"], "Quarter");
    assert_eq!(anchor["data"]["node"]["format"]["bold"], "true");
    assert_eq!(anchor["data"]["node"]["format"]["merge"], "A1:B2");
    assert_eq!(anchor["data"]["node"]["format"]["mergeAnchor"], "true");

    let queried = run(
        &provider,
        &["query", document.to_str().unwrap(), "mergeCell", "--json"],
    );
    assert_eq!(queried["data"]["matches"], 1);
    assert_eq!(queried["data"]["results"][0]["format"]["ref"], "A1:B2");

    let idempotent = run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/sheet1/a1:b2",
            "--merge-cells",
            "true",
            "--json",
        ],
    );
    assert_eq!(idempotent["data"]["operation"], "merge-cells");
    assert_eq!(idempotent["data"]["changed"], false);

    let imprecise = run_failure(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/A1:C3",
            "--merge-cells",
            "false",
            "--json",
        ],
    );
    assert_eq!(
        imprecise["error"]["code"],
        "use.office.spreadsheet_merge_not_exact"
    );
    assert_eq!(
        imprecise["error"]["details"]["validRanges"],
        serde_json::json!(["A1:B2"])
    );

    let unmerged = run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/A1:B2",
            "--merge-cells",
            "false",
            "--json",
        ],
    );
    assert_eq!(unmerged["data"]["operation"], "unmerge-cells");
    assert_eq!(unmerged["data"]["changed"], true);
    assert_eq!(unmerged["data"]["node"]["format"]["merge"], "false");
    assert!(!provider.exists());
}

#[test]
fn native_merge_batch_rolls_back_and_cli_values_are_explicit() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("atomic.xlsx");
    let batch = temp.path().join("merges.json");
    create(&provider, &document);
    let before = std::fs::read(&document).unwrap();

    std::fs::write(
        &batch,
        serde_json::to_vec(&serde_json::json!({
            "schemaVersion": 1,
            "mutations": [
                { "operation": "merge-cells", "path": "/Sheet1/A1:B1" },
                { "operation": "merge-cells", "path": "/Sheet1/B1:C1" }
            ]
        }))
        .unwrap(),
    )
    .unwrap();
    let failed = run_failure(
        &provider,
        &[
            "batch",
            document.to_str().unwrap(),
            "--input",
            batch.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(
        failed["error"]["code"],
        "use.office.spreadsheet_merge_overlap"
    );
    assert_eq!(std::fs::read(&document).unwrap(), before);

    let invalid = run_failure(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/A1:B1",
            "--merge-cells",
            "yes",
            "--json",
        ],
    );
    assert_eq!(invalid["error"]["code"], "use.cli.invalid_usage");
    assert_eq!(std::fs::read(&document).unwrap(), before);
    assert!(!provider.exists());
}
