use std::io::Write;
use std::path::Path;
use std::process::{Command, Output, Stdio};

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

fn execute_stdin(provider: &Path, args: &[&str], input: &[u8]) -> Output {
    let mut child = Command::new(binary())
        .args(args)
        .env("A3S_OFFICECLI_EXECUTABLE", provider)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    child.stdin.take().unwrap().write_all(input).unwrap();
    child.wait_with_output().unwrap()
}

fn success(output: Output) -> serde_json::Value {
    assert!(output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

fn failure(output: Output) -> serde_json::Value {
    assert!(!output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

fn run(provider: &Path, args: &[&str]) -> serde_json::Value {
    success(execute(provider, args))
}

#[test]
fn native_cli_imports_files_and_stdin_without_invoking_the_compatibility_provider() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("import.xlsx");
    let source = temp.path().join("source.tab");
    std::fs::write(
        &source,
        b"Name\tAmount\tDate\nAlpha\t42\t2026-07-17\nBeta\tTRUE\t2026-07-18",
    )
    .unwrap();
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);

    let imported = run(
        &provider,
        &[
            "import",
            document.to_str().unwrap(),
            "/Sheet1",
            source.to_str().unwrap(),
            "--header",
            "--start-cell",
            "B2",
            "--json",
        ],
    );
    assert_eq!(
        imported["data"]["operation"],
        "import-spreadsheet-delimited"
    );
    assert_eq!(imported["data"]["result"]["format"], "tsv");
    assert_eq!(imported["data"]["result"]["range"], "B2:D4");
    assert_eq!(imported["data"]["result"]["rowCount"], 3);
    assert_eq!(imported["data"]["result"]["freezePath"], "/Sheet1/freeze");

    let amount = run(
        &provider,
        &["get", document.to_str().unwrap(), "/Sheet1/C3", "--json"],
    );
    assert_eq!(amount["data"]["node"]["text"], "42");
    assert_eq!(amount["data"]["node"]["format"]["valueType"], "Number");
    let freeze = run(
        &provider,
        &[
            "get",
            document.to_str().unwrap(),
            "/Sheet1/freeze",
            "--json",
        ],
    );
    assert_eq!(freeze["data"]["node"]["format"]["frozenRows"], "2");
    assert_eq!(freeze["data"]["node"]["format"]["topLeftCell"], "B3");
    let dump = run(&provider, &["dump", document.to_str().unwrap(), "--json"]);
    assert!(dump["data"]["artifact"]["mutations"]
        .as_array()
        .unwrap()
        .iter()
        .any(|mutation| mutation["operation"] == "set-spreadsheet-frozen-pane"));

    let stdin_document = temp.path().join("stdin.xlsx");
    let copy = temp.path().join("stdin-copy.xlsx");
    run(
        &provider,
        &["create", stdin_document.to_str().unwrap(), "--json"],
    );
    let output = execute_stdin(
        &provider,
        &[
            "import",
            stdin_document.to_str().unwrap(),
            "/Sheet1",
            "--stdin",
            "--format",
            "csv",
            "--output",
            copy.to_str().unwrap(),
            "--json",
        ],
        b"one,two\nthree,four",
    );
    let imported = success(output);
    assert_eq!(imported["data"]["source"], "stdin");
    assert_eq!(imported["data"]["inPlace"], false);
    assert!(copy.is_file());
    let copied = run(
        &provider,
        &["get", copy.to_str().unwrap(), "/Sheet1/B2", "--json"],
    );
    assert_eq!(copied["data"]["node"]["text"], "four");
    let original = failure(execute(
        &provider,
        &[
            "get",
            stdin_document.to_str().unwrap(),
            "/Sheet1/A1",
            "--json",
        ],
    ));
    assert_eq!(original["error"]["code"], "use.office.node_not_found");
    assert!(!provider.exists());
}

#[test]
fn native_cli_rejects_ambiguous_and_non_utf8_import_sources_before_mutation() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("safe.xlsx");
    let first = temp.path().join("first.csv");
    let second = temp.path().join("second.csv");
    let invalid = temp.path().join("invalid.csv");
    let malformed = temp.path().join("malformed.csv");
    std::fs::write(&first, b"a").unwrap();
    std::fs::write(&second, b"b").unwrap();
    std::fs::write(&invalid, [0xff, 0xfe]).unwrap();
    std::fs::write(&malformed, b"\"unterminated").unwrap();
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);
    let before = std::fs::read(&document).unwrap();

    let ambiguous = failure(execute(
        &provider,
        &[
            "import",
            document.to_str().unwrap(),
            "/Sheet1",
            first.to_str().unwrap(),
            "--file",
            second.to_str().unwrap(),
            "--json",
        ],
    ));
    assert_eq!(ambiguous["error"]["code"], "use.cli.invalid_usage");
    assert_eq!(std::fs::read(&document).unwrap(), before);

    let invalid_utf8 = failure(execute(
        &provider,
        &[
            "import",
            document.to_str().unwrap(),
            "/Sheet1",
            "--file",
            invalid.to_str().unwrap(),
            "--json",
        ],
    ));
    assert_eq!(
        invalid_utf8["error"]["code"],
        "use.office.spreadsheet_import_input_invalid"
    );
    assert_eq!(std::fs::read(&document).unwrap(), before);

    let malformed_csv = failure(execute(
        &provider,
        &[
            "import",
            document.to_str().unwrap(),
            "/Sheet1",
            malformed.to_str().unwrap(),
            "--json",
        ],
    ));
    assert_eq!(
        malformed_csv["error"]["code"],
        "use.office.spreadsheet_import_delimited_invalid"
    );
    assert_eq!(malformed_csv["error"]["details"]["row"], 1);
    assert_eq!(malformed_csv["error"]["details"]["column"], 1);
    assert_eq!(std::fs::read(&document).unwrap(), before);
    assert!(!provider.exists());
}
