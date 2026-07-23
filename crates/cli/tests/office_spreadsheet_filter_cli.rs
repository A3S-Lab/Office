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
fn native_cli_manages_worksheet_and_table_filters_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("filters.xlsx");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);

    let added = run(
        &provider,
        &[
            "add",
            document.to_str().unwrap(),
            "/Sheet1",
            "--type",
            "auto-filter",
            "--range",
            "A1:C20",
            "--filter",
            r#"{"column":0,"criteria":{"type":"values","values":["Open","Closed"],"includeBlanks":true}}"#,
            "--filter",
            r#"{"column":2,"criteria":{"type":"greater-than","value":"100"}}"#,
            "--json",
        ],
    );
    assert_eq!(added["data"]["operation"], "add-spreadsheet-auto-filter");
    assert_eq!(added["data"]["path"], "/Sheet1/autofilter");
    assert_eq!(added["data"]["node"]["type"], "auto-filter");
    assert_eq!(
        added["data"]["node"]["children"][0]["format"]["criteriaType"],
        "values"
    );

    let before_conflict = std::fs::read(&document).unwrap();
    let conflict = run_failure(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/autofilter",
            "--filter",
            r#"{"column":0,"criteria":{"type":"blanks"}}"#,
            "--text",
            "blocked",
            "--json",
        ],
    );
    assert_eq!(conflict["error"]["code"], "use.cli.invalid_usage");
    assert_eq!(std::fs::read(&document).unwrap(), before_conflict);

    run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/autofilter",
            "--range",
            "A1:C25",
            "--json",
        ],
    );
    let range_only = run(
        &provider,
        &[
            "get",
            document.to_str().unwrap(),
            "/Sheet1/autofilter",
            "--depth",
            "2",
            "--json",
        ],
    );
    assert_eq!(range_only["data"]["node"]["childCount"], 2);

    let updated = run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/auto-filter",
            "--range",
            "B2:D30",
            "--filter",
            r#"{"column":1,"criteria":{"type":"dynamic","kind":"this-month"}}"#,
            "--json",
        ],
    );
    assert_eq!(updated["data"]["operation"], "set-spreadsheet-auto-filter");
    assert_eq!(updated["data"]["node"]["format"]["ref"], "B2:D30");

    let queried = run(
        &provider,
        &[
            "query",
            document.to_str().unwrap(),
            "filtercolumn[criteriaType=dynamic]",
            "--json",
        ],
    );
    assert_eq!(queried["data"]["matches"], 1);

    let cleared = run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/autofilter",
            "--clear-filters",
            "--json",
        ],
    );
    assert_eq!(cleared["data"]["operation"], "set-spreadsheet-auto-filter");
    let read = run(
        &provider,
        &[
            "get",
            document.to_str().unwrap(),
            "/Sheet1/autofilter",
            "--depth",
            "1",
            "--json",
        ],
    );
    assert_eq!(read["data"]["node"]["childCount"], 0);

    let dump = run(&provider, &["dump", document.to_str().unwrap(), "--json"]);
    assert!(dump["data"]["artifact"]["mutations"]
        .as_array()
        .unwrap()
        .iter()
        .any(|mutation| mutation["operation"] == "add-spreadsheet-auto-filter"));

    run(
        &provider,
        &[
            "remove",
            document.to_str().unwrap(),
            "/Sheet1/autofilter",
            "--json",
        ],
    );
    let table = run(
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
            "A1:B10",
            "--table-column",
            "Region",
            "--table-column",
            "Amount",
            "--filter",
            r#"{"column":0,"criteria":{"type":"contains","value":"West"}}"#,
            "--json",
        ],
    );
    assert!(table["data"]["node"]["children"]
        .as_array()
        .unwrap()
        .iter()
        .any(|child| child["type"] == "auto-filter"));
    run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/table[1]",
            "--show-row-stripes",
            "false",
            "--json",
        ],
    );
    let preserved_table = run(
        &provider,
        &[
            "get",
            document.to_str().unwrap(),
            "/Sheet1/table[1]",
            "--depth",
            "2",
            "--json",
        ],
    );
    let preserved_filter = preserved_table["data"]["node"]["children"]
        .as_array()
        .unwrap()
        .iter()
        .find(|child| child["type"] == "auto-filter")
        .unwrap();
    assert_eq!(preserved_filter["childCount"], 1);
    run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/table[1]",
            "--clear-filters",
            "--json",
        ],
    );
    let table_read = run(
        &provider,
        &[
            "get",
            document.to_str().unwrap(),
            "/Sheet1/table[1]",
            "--depth",
            "2",
            "--json",
        ],
    );
    let table_filter = table_read["data"]["node"]["children"]
        .as_array()
        .unwrap()
        .iter()
        .find(|child| child["type"] == "auto-filter")
        .unwrap();
    assert_eq!(table_filter["childCount"], 0);
    assert!(!provider.exists());
}

#[test]
fn native_cli_spreadsheet_filter_batch_is_atomic() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("filter-batch.xlsx");
    let batch = temp.path().join("batch.json");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);
    let before = std::fs::read(&document).unwrap();
    std::fs::write(
        &batch,
        serde_json::to_vec(&serde_json::json!({
            "schemaVersion": 1,
            "mutations": [
                {
                    "operation": "add-spreadsheet-auto-filter",
                    "sheet": "/Sheet1",
                    "filter": {
                        "range": "A1:B10",
                        "columns": [{
                            "column": 0,
                            "criteria": {"type": "non-blanks"}
                        }]
                    }
                },
                {
                    "operation": "add-spreadsheet-auto-filter",
                    "sheet": "/Sheet1",
                    "filter": {"range": "D1:E10"}
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
        "use.office.spreadsheet_filter_exists"
    );
    assert_eq!(std::fs::read(&document).unwrap(), before);
    assert!(!provider.exists());
}
