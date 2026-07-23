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

fn success(output: Output) -> serde_json::Value {
    assert!(output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

fn run(provider: &Path, args: &[&str]) -> serde_json::Value {
    success(execute(provider, args))
}

#[test]
fn native_cli_recalculates_typed_formulas_and_spills_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let source = temp.path().join("formulas.xlsx");
    let output = temp.path().join("calculated.xlsx");
    run(&provider, &["create", source.to_str().unwrap(), "--json"]);
    for (path, option, value) in [
        ("/Sheet1/A1", "--number", "2"),
        ("/Sheet1/B1", "--formula", "A1*3"),
        ("/Sheet1/C1", "--formula", "SEQUENCE(2,2,1,1)"),
    ] {
        run(
            &provider,
            &[
                "set",
                source.to_str().unwrap(),
                path,
                option,
                value,
                "--json",
            ],
        );
    }

    let recalculated = run(
        &provider,
        &[
            "recalculate",
            source.to_str().unwrap(),
            "--output",
            output.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(
        recalculated["data"]["operation"],
        "recalculate-spreadsheet-formulas"
    );
    assert_eq!(recalculated["data"]["result"]["formulaCount"], 2);
    assert_eq!(recalculated["data"]["result"]["spillCellCount"], 3);
    assert_eq!(recalculated["data"]["inPlace"], false);

    let b1 = run(
        &provider,
        &["get", output.to_str().unwrap(), "/Sheet1/B1", "--json"],
    );
    assert_eq!(b1["data"]["node"]["text"], "6");
    assert_eq!(b1["data"]["node"]["format"]["formulaCached"], "true");
    let d2 = run(
        &provider,
        &["get", output.to_str().unwrap(), "/Sheet1/D2", "--json"],
    );
    assert_eq!(d2["data"]["node"]["text"], "4");
    assert!(d2["data"]["node"]["format"].get("formula").is_none());
    assert!(!provider.exists());
}

#[test]
fn native_cli_recalculation_failure_does_not_change_the_file() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("unsupported.xlsx");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);
    run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/A1",
            "--formula",
            "SHELL(\"unsafe\")",
            "--json",
        ],
    );
    let before = std::fs::read(&document).unwrap();
    let failure = execute(
        &provider,
        &["recalculate", document.to_str().unwrap(), "--json"],
    );
    assert!(!failure.status.success(), "{failure:?}");
    let error: serde_json::Value = serde_json::from_slice(&failure.stdout).unwrap();
    assert_eq!(
        error["error"]["code"],
        "use.office.spreadsheet_formula_function_unsupported"
    );
    assert_eq!(std::fs::read(&document).unwrap(), before);
    assert!(!provider.exists());
}
