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

#[test]
fn native_cli_manages_complete_spreadsheet_named_range_lifecycle() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("names.xlsx");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);

    let added = run(
        &provider,
        &[
            "add",
            document.to_str().unwrap(),
            "/",
            "--type",
            "named-range",
            "--name",
            "Revenue",
            "--ref",
            "'Sheet1'!$A$2:$A$20",
            "--comment",
            "Workbook revenue",
            "--volatile",
            "true",
            "--json",
        ],
    );
    assert_eq!(added["data"]["operation"], "add-named-range");
    assert_eq!(
        added["data"]["path"],
        "/namedrange[@name=Revenue][@scope=workbook]"
    );
    assert_eq!(added["data"]["node"]["format"]["volatile"], "true");

    let local = run(
        &provider,
        &[
            "add",
            document.to_str().unwrap(),
            "/Sheet1",
            "--type",
            "namedrange",
            "--name",
            "LocalStatus",
            "--ref",
            "B2:B20",
            "--json",
        ],
    );
    assert_eq!(
        local["data"]["path"],
        "/namedrange[@name=LocalStatus][@scope=Sheet1]"
    );
    assert_eq!(local["data"]["node"]["format"]["ref"], "'Sheet1'!B2:B20");

    let listed = run(
        &provider,
        &[
            "get",
            document.to_str().unwrap(),
            "/namedrange",
            "--depth",
            "1",
            "--json",
        ],
    );
    assert_eq!(listed["data"]["node"]["childCount"], 2);
    let queried = run(
        &provider,
        &[
            "query",
            document.to_str().unwrap(),
            "namedrange[scope=Sheet1]",
            "--json",
        ],
    );
    assert_eq!(queried["data"]["matches"], 1);

    let updated = run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/namedrange[LocalStatus]",
            "--name",
            "WorkflowStatus",
            "--ref",
            "'Sheet1'!$C$2:$C$20",
            "--scope",
            "workbook",
            "--comment",
            "none",
            "--volatile",
            "false",
            "--json",
        ],
    );
    assert_eq!(updated["data"]["operation"], "set-named-range");
    assert_eq!(
        updated["data"]["path"],
        "/namedrange[@name=WorkflowStatus][@scope=workbook]"
    );
    assert_eq!(updated["data"]["node"]["format"]["volatile"], "false");
    assert!(updated["data"]["node"]["format"].get("comment").is_none());

    let removed = run(
        &provider,
        &[
            "remove",
            document.to_str().unwrap(),
            "/namedrange[Revenue]",
            "--json",
        ],
    );
    assert_eq!(removed["data"]["operation"], "remove");
    assert_eq!(removed["data"]["changed"], true);
    assert!(!provider.exists());
}

#[test]
fn native_cli_named_range_batch_is_versioned_and_atomic() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("batch.xlsx");
    let batch = temp.path().join("batch.json");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);
    let before = std::fs::read(&document).unwrap();
    std::fs::write(
        &batch,
        serde_json::to_vec(&serde_json::json!({
            "schemaVersion": 1,
            "mutations": [
                {
                    "operation": "add-named-range",
                    "namedRange": { "name": "Revenue", "ref": "'Sheet1'!$A$1" }
                },
                {
                    "operation": "add-named-range",
                    "namedRange": { "name": "revenue", "ref": "'Sheet1'!$B$1" }
                }
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
        failed["error"]["code"], "use.office.spreadsheet_named_range_duplicate",
        "{failed}"
    );
    assert_eq!(std::fs::read(&document).unwrap(), before);
    assert!(!provider.exists());
}
