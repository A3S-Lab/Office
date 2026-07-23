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
fn native_cli_manages_spreadsheet_table_lifecycle_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("tables.xlsx");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);

    let added = run(
        &provider,
        &[
            "add",
            document.to_str().unwrap(),
            "/Sheet1",
            "--type",
            "table",
            "--name",
            "Sales",
            "--range",
            "A1:C4",
            "--table-column",
            "Name",
            "--table-column",
            "Qty",
            "--table-column",
            "Price",
            "--style",
            "medium:4",
            "--show-last-column",
            "true",
            "--json",
        ],
    );
    assert_eq!(added["data"]["operation"], "add-spreadsheet-table");
    assert_eq!(added["data"]["path"], "/Sheet1/table[1]");
    assert_eq!(added["data"]["node"]["type"], "table");
    assert_eq!(added["data"]["node"]["format"]["name"], "Sales");
    assert_eq!(added["data"]["node"]["format"]["styleNumber"], "4");
    assert_eq!(added["data"]["node"]["children"][1]["text"], "Qty");

    let before_conflict = std::fs::read(&document).unwrap();
    let conflict = run_failure(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/table[1]",
            "--totals-row",
            "true",
            "--text",
            "blocked",
            "--json",
        ],
    );
    assert_eq!(conflict["error"]["code"], "use.cli.invalid_usage");
    assert_eq!(std::fs::read(&document).unwrap(), before_conflict);

    let queried = run(
        &provider,
        &[
            "query",
            document.to_str().unwrap(),
            "table[name=Sales]",
            "--json",
        ],
    );
    assert_eq!(queried["data"]["matches"], 1);

    let updated = run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/table[1]",
            "--name",
            "Inventory",
            "--display-name",
            "InventoryView",
            "--range",
            "B2:D6",
            "--table-column",
            "Item",
            "--table-column",
            "Units",
            "--table-column",
            "Cost",
            "--totals-row",
            "true",
            "--style",
            "dark:2",
            "--show-row-stripes",
            "false",
            "--show-column-stripes",
            "true",
            "--json",
        ],
    );
    assert_eq!(updated["data"]["operation"], "set-spreadsheet-table");
    assert_eq!(updated["data"]["node"]["format"]["name"], "Inventory");
    assert_eq!(
        updated["data"]["node"]["format"]["displayName"],
        "InventoryView"
    );
    assert_eq!(updated["data"]["node"]["format"]["ref"], "B2:D6");
    assert_eq!(updated["data"]["node"]["format"]["totalsRow"], "true");

    let dump = run(&provider, &["dump", document.to_str().unwrap(), "--json"]);
    assert!(dump["data"]["artifact"]["mutations"]
        .as_array()
        .unwrap()
        .iter()
        .any(|mutation| mutation["operation"] == "add-spreadsheet-table"));

    let removed = run(
        &provider,
        &[
            "remove",
            document.to_str().unwrap(),
            "/Sheet1/table[1]",
            "--json",
        ],
    );
    assert_eq!(removed["data"]["operation"], "remove");
    assert!(!provider.exists());
}

#[test]
fn native_cli_spreadsheet_table_batch_is_atomic() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("table-batch.xlsx");
    let batch = temp.path().join("batch.json");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);
    let before = std::fs::read(&document).unwrap();
    std::fs::write(
        &batch,
        serde_json::to_vec(&serde_json::json!({
            "schemaVersion": 1,
            "mutations": [
                {
                    "operation": "add-spreadsheet-table",
                    "sheet": "/Sheet1",
                    "table": {
                        "name": "First",
                        "range": "A1:B3",
                        "columns": [{"name": "A"}, {"name": "B"}]
                    }
                },
                {
                    "operation": "add-spreadsheet-table",
                    "sheet": "/Sheet1",
                    "table": {
                        "name": "Overlap",
                        "range": "B2:C4",
                        "columns": [{"name": "B"}, {"name": "C"}]
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
        "use.office.spreadsheet_table_overlap"
    );
    assert_eq!(std::fs::read(&document).unwrap(), before);
    assert!(!provider.exists());
}
