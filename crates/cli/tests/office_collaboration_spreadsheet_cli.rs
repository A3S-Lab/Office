mod support;

use std::fs;
use std::path::Path;
use std::process::{Command, Output};

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use support::{spreadsheet_cell, spreadsheet_collaboration_fixture};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

fn execute(args: &[&str]) -> Output {
    Command::new(binary()).args(args).output().unwrap()
}

fn run(args: &[&str]) -> serde_json::Value {
    let output = execute(args);
    assert!(output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

fn run_failure(args: &[&str]) -> serde_json::Value {
    let output = execute(args);
    assert!(!output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

#[test]
fn cli_manages_browser_compatible_spreadsheet_cells_with_conflict_guards() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("spreadsheet.replica");
    let fixture = temp.path().join("browser-spreadsheet.update");
    fs::write(&fixture, spreadsheet_collaboration_fixture()).unwrap();
    run(&[
        "collab",
        "join",
        replica.to_str().unwrap(),
        "--artifact-id",
        "fixture-spreadsheet",
        "--kind",
        "spreadsheet",
        "--actor-id",
        "coding-agent-spreadsheet-cli",
        "--actor-kind",
        "agent",
        "--mode",
        "edit",
        "--operation-id",
        "join-browser-spreadsheet-cli",
        "--input",
        fixture.to_str().unwrap(),
        "--client-id",
        "900014",
        "--json",
    ]);

    let set = mutate(
        &replica,
        "spreadsheet-set-cell-cli",
        r#"{"type":"spreadsheet-set-cell","sheetId":"sheet-data","row":1,"column":0,"expectedCell":{"v":10,"m":"10","ct":{"fa":"0.00","t":"n"}},"nextCell":{"v":12,"m":"12","f":"=6*2","ct":{"fa":"0.00","t":"n"}}}"#,
    );
    assert_eq!(set["data"]["sequence"], 1);
    assert_eq!(set["data"]["mutation"]["type"], "spreadsheet-set-cell");
    let create = mutate(
        &replica,
        "spreadsheet-create-cell-cli",
        r#"{"type":"spreadsheet-set-cell","sheetId":"sheet-empty","row":100,"column":5,"expectedCell":null,"nextCell":{"v":"CLI sparse","m":"CLI sparse","ps":{"value":"CLI note","isShow":false}}}"#,
    );
    assert_eq!(create["data"]["sequence"], 2);
    let delete = mutate(
        &replica,
        "spreadsheet-delete-cell-cli",
        r#"{"type":"spreadsheet-delete-cell","sheetId":"sheet-sparse","row":5,"column":3,"expectedCell":{"v":"edge","m":"edge"}}"#,
    );
    assert_eq!(delete["data"]["sequence"], 3);

    let conflict = run_failure(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        r#"{"type":"spreadsheet-set-cell","sheetId":"sheet-data","row":1,"column":0,"expectedCell":{"v":10,"m":"10","ct":{"fa":"0.00","t":"n"}},"nextCell":{"v":99,"m":"99","ct":{"fa":"0.00","t":"n"}}}"#,
        "--actor-id",
        "coding-agent-spreadsheet-cli",
        "--operation-id",
        "spreadsheet-conflict-cli",
        "--artifact-id",
        "fixture-spreadsheet",
        "--kind",
        "spreadsheet",
        "--mode",
        "edit",
        "--json",
    ]);
    assert_eq!(
        conflict["error"]["code"],
        "office.collaboration.mutation_match_conflict"
    );

    let exported = run(&["collab", "diff", replica.to_str().unwrap(), "--json"]);
    let update = STANDARD
        .decode(exported["data"]["updateBase64"].as_str().unwrap())
        .unwrap();
    assert_eq!(
        spreadsheet_cell(&update, "sheet-data", 1, 0),
        Some(serde_json::json!({
            "v": 12,
            "m": "12",
            "f": "=6*2",
            "ct": { "fa": "0.00", "t": "n" },
        }))
    );
    assert_eq!(
        spreadsheet_cell(&update, "sheet-empty", 100, 5),
        Some(serde_json::json!({
            "v": "CLI sparse",
            "m": "CLI sparse",
            "ps": { "value": "CLI note", "isShow": false },
        }))
    );
    assert_eq!(spreadsheet_cell(&update, "sheet-sparse", 5, 3), None);
    assert_eq!(
        run(&["collab", "inspect", replica.to_str().unwrap(), "--json"])["data"]["currentSequence"],
        3
    );
}

fn mutate(replica: &Path, operation_id: &str, mutation: &str) -> serde_json::Value {
    run(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        mutation,
        "--actor-id",
        "coding-agent-spreadsheet-cli",
        "--operation-id",
        operation_id,
        "--artifact-id",
        "fixture-spreadsheet",
        "--kind",
        "spreadsheet",
        "--mode",
        "edit",
        "--json",
    ])
}
