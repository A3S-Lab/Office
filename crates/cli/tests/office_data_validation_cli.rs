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
fn native_cli_manages_complete_spreadsheet_data_validation_lifecycle() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("validation.xlsx");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);

    let added = run(
        &provider,
        &[
            "add",
            document.to_str().unwrap(),
            "/Sheet1",
            "--type",
            "data-validation",
            "--validation-type",
            "list",
            "--range",
            "A2:A20",
            "--range",
            "C2:C20",
            "--formula1",
            "Draft,Review,Approved",
            "--in-cell-dropdown",
            "false",
            "--prompt-title",
            "Status",
            "--prompt",
            "Choose a workflow state",
            "--json",
        ],
    );
    assert_eq!(added["data"]["operation"], "add-data-validation");
    assert_eq!(added["data"]["path"], "/Sheet1/dataValidation[1]");
    assert_eq!(added["data"]["node"]["format"]["type"], "list");
    assert_eq!(added["data"]["node"]["format"]["ref"], "A2:A20 C2:C20");
    assert_eq!(added["data"]["node"]["format"]["inCellDropdown"], "false");

    let queried = run(
        &provider,
        &[
            "query",
            document.to_str().unwrap(),
            "dataValidation[type=list]",
            "--json",
        ],
    );
    assert_eq!(queried["data"]["matches"], 1);

    let updated = run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/validation[1]",
            "--validation-type",
            "whole",
            "--range",
            "B2:B50",
            "--operator",
            "between",
            "--formula1",
            "18",
            "--formula2",
            "120",
            "--allow-blank",
            "false",
            "--error-style",
            "warning",
            "--error-title",
            "Age outside range",
            "--error-message",
            "Enter an age from 18 through 120",
            "--json",
        ],
    );
    assert_eq!(updated["data"]["operation"], "set-data-validation");
    assert_eq!(updated["data"]["path"], "/Sheet1/dataValidation[1]");
    assert_eq!(updated["data"]["node"]["format"]["type"], "whole");
    assert_eq!(updated["data"]["node"]["format"]["operator"], "between");
    assert_eq!(updated["data"]["node"]["format"]["allowBlank"], "false");
    assert_eq!(updated["data"]["node"]["format"]["errorStyle"], "warning");

    let invalid = run_failure(
        &provider,
        &[
            "add",
            document.to_str().unwrap(),
            "/Sheet1",
            "--type",
            "data-validation",
            "--validation-type",
            "list",
            "--range",
            "B50:C60",
            "--formula1",
            "Open,Closed",
            "--json",
        ],
    );
    assert_eq!(
        invalid["error"]["code"],
        "use.office.spreadsheet_validation_overlap"
    );

    let removed = run(
        &provider,
        &[
            "remove",
            document.to_str().unwrap(),
            "/Sheet1/dataValidation[1]",
            "--json",
        ],
    );
    assert_eq!(removed["data"]["operation"], "remove");
    assert_eq!(removed["data"]["changed"], true);
    assert!(!provider.exists());
}

#[test]
fn native_cli_data_validation_batch_is_versioned_and_atomic() {
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
                    "operation": "add-data-validation",
                    "sheet": "/Sheet1",
                    "validation": {
                        "type": "list",
                        "ranges": ["A1:A10"],
                        "formula1": "Open,Closed"
                    }
                },
                {
                    "operation": "add-data-validation",
                    "sheet": "/Sheet1",
                    "validation": {
                        "type": "list",
                        "ranges": ["A10:B20"],
                        "formula1": "Yes,No"
                    }
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
        failed["error"]["code"],
        "use.office.spreadsheet_validation_overlap"
    );
    assert_eq!(std::fs::read(&document).unwrap(), before);
    assert!(!provider.exists());
}
