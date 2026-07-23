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
fn native_cli_manages_conditional_formatting_lifecycle_and_visual_rules() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("conditional-formatting.xlsx");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);

    let comparison = run(
        &provider,
        &[
            "add",
            document.to_str().unwrap(),
            "/Sheet1",
            "--type",
            "conditional-format",
            "--rule-type",
            "cell-is",
            "--range",
            "A2:A20",
            "--operator",
            "greater-than",
            "--formula1",
            "80",
            "--fill",
            "C6EFCE",
            "--text-color",
            "006100",
            "--bold",
            "true",
            "--stop-if-true",
            "true",
            "--json",
        ],
    );
    assert_eq!(comparison["data"]["operation"], "add-conditional-format");
    assert_eq!(comparison["data"]["path"], "/Sheet1/cf[1]");
    assert_eq!(comparison["data"]["node"]["type"], "conditional-formatting");
    assert_eq!(comparison["data"]["node"]["format"]["type"], "cellIs");
    assert_eq!(comparison["data"]["node"]["format"]["fill"], "C6EFCE");

    for args in [
        vec![
            "--rule-type",
            "data-bar",
            "--range",
            "B2:B20",
            "--color",
            "638EC6",
            "--min",
            "min",
            "--max",
            "number:100",
        ],
        vec![
            "--rule-type",
            "color-scale",
            "--range",
            "C2:C20",
            "--min-color",
            "F8696B",
            "--midpoint",
            "percentile:50",
            "--mid-color",
            "FFEB84",
            "--max-color",
            "63BE7B",
        ],
        vec![
            "--rule-type",
            "icon-set",
            "--range",
            "D2:D20",
            "--icon-set",
            "3-traffic-lights-1",
            "--threshold",
            "percent:0",
            "--threshold",
            "percent:33",
            "--threshold",
            "percent:67",
            "--reverse",
            "true",
        ],
    ] {
        let mut command = vec!["add", document.to_str().unwrap(), "/Sheet1", "--type", "cf"];
        command.extend(args);
        command.push("--json");
        run(&provider, &command);
    }

    let queried = run(
        &provider,
        &[
            "query",
            document.to_str().unwrap(),
            "conditionalFormatting[type=iconSet]",
            "--json",
        ],
    );
    assert_eq!(queried["data"]["matches"], 1);

    let updated = run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/conditional-formatting[1]",
            "--formula1",
            "90",
            "--fill",
            "FFEB9C",
            "--json",
        ],
    );
    assert_eq!(updated["data"]["operation"], "set-conditional-format");
    assert_eq!(updated["data"]["path"], "/Sheet1/cf[1]");
    assert_eq!(updated["data"]["node"]["format"]["formula1"], "90");
    assert_eq!(updated["data"]["node"]["format"]["ref"], "A2:A20");
    assert_eq!(updated["data"]["node"]["format"]["fill"], "FFEB9C");
    assert_eq!(updated["data"]["node"]["format"]["fontBold"], "true");

    let reopened = run(
        &provider,
        &["get", document.to_str().unwrap(), "/Sheet1/cf[1]", "--json"],
    );
    assert_eq!(reopened["data"]["node"]["format"]["formula1"], "90");

    let removed = run(
        &provider,
        &[
            "remove",
            document.to_str().unwrap(),
            "/Sheet1/cf[2]",
            "--json",
        ],
    );
    assert_eq!(removed["data"]["operation"], "remove");
    assert!(!provider.exists());
}

#[test]
fn native_cli_conditional_format_batch_is_atomic() {
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
                    "operation": "add-conditional-format",
                    "sheet": "/Sheet1",
                    "conditionalFormat": {
                        "ranges": ["A1:A10"],
                        "rule": {
                            "type": "cellIs",
                            "operator": "greaterThan",
                            "formula1": "80",
                            "format": {"fill": {"red": 198, "green": 239, "blue": 206}}
                        }
                    }
                },
                {
                    "operation": "add-conditional-format",
                    "sheet": "/Sheet1",
                    "conditionalFormat": {
                        "ranges": ["B1:B10"],
                        "rule": {
                            "type": "colorScale",
                            "min": {"kind": "min"},
                            "minColor": {"red": 248, "green": 105, "blue": 107},
                            "mid": {"kind": "percentile", "value": "101"},
                            "midColor": {"red": 255, "green": 235, "blue": 132},
                            "max": {"kind": "max"},
                            "maxColor": {"red": 99, "green": 190, "blue": 123}
                        }
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
        "use.office.spreadsheet_conditional_format_threshold_invalid"
    );
    assert_eq!(std::fs::read(&document).unwrap(), before);
    assert!(!provider.exists());
}
