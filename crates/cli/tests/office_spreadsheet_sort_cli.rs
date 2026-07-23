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

fn set(provider: &Path, document: &Path, path: &str, option: &str, value: &str) {
    run(
        provider,
        &[
            "set",
            document.to_str().unwrap(),
            path,
            option,
            value,
            "--json",
        ],
    );
}

#[test]
fn native_cli_sorts_spreadsheet_rows_and_manages_persisted_sort_state() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("sorted.xlsx");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);
    for (path, value) in [
        ("/Sheet1/A1", "Name"),
        ("/Sheet1/B1", "Rank"),
        ("/Sheet1/A2", "Beta"),
        ("/Sheet1/A3", "Alpha"),
    ] {
        set(&provider, &document, path, "--text", value);
    }
    set(&provider, &document, "/Sheet1/B2", "--number", "2");
    set(&provider, &document, "/Sheet1/B3", "--number", "1");

    let sorted = run(
        &provider,
        &[
            "sort",
            document.to_str().unwrap(),
            "/Sheet1/A1:B3",
            "--key",
            "B:asc",
            "--header",
            "true",
            "--case-sensitive",
            "false",
            "--json",
        ],
    );
    assert_eq!(sorted["data"]["operation"], "sort-spreadsheet-range");
    assert_eq!(sorted["data"]["path"], "/Sheet1/sort");
    assert_eq!(sorted["data"]["sort"]["keys"][0]["column"], "B");
    assert_eq!(sorted["data"]["node"]["format"]["ref"], "A1:B3");
    assert_eq!(sorted["data"]["node"]["format"]["header"], "true");

    let first = run(
        &provider,
        &["get", document.to_str().unwrap(), "/Sheet1/A2", "--json"],
    );
    assert_eq!(first["data"]["node"]["text"], "Alpha");
    let dump = run(&provider, &["dump", document.to_str().unwrap(), "--json"]);
    assert!(dump["data"]["artifact"]["mutations"]
        .as_array()
        .unwrap()
        .iter()
        .any(|mutation| mutation["operation"] == "sort-spreadsheet-range"));

    let before_invalid = std::fs::read(&document).unwrap();
    let invalid = run_failure(
        &provider,
        &[
            "sort",
            document.to_str().unwrap(),
            "/Sheet1/A1:B3",
            "--key",
            "B:sideways",
            "--json",
        ],
    );
    assert_eq!(invalid["error"]["code"], "use.cli.invalid_usage");
    assert_eq!(std::fs::read(&document).unwrap(), before_invalid);

    run(
        &provider,
        &[
            "remove",
            document.to_str().unwrap(),
            "/Sheet1/sort",
            "--json",
        ],
    );
    let retained = run(
        &provider,
        &["get", document.to_str().unwrap(), "/Sheet1/A2", "--json"],
    );
    assert_eq!(retained["data"]["node"]["text"], "Alpha");
    assert!(!provider.exists());
}
