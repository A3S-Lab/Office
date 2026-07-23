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

fn create_presentation(provider: &Path, document: &Path) {
    run(provider, &["create", document.to_str().unwrap(), "--json"]);
    run(
        provider,
        &[
            "add",
            document.to_str().unwrap(),
            "/",
            "--type",
            "slide",
            "--text",
            "Native table",
            "--json",
        ],
    );
}

#[test]
fn native_cli_structurally_edits_presentation_tables_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("table.pptx");
    create_presentation(&provider, &document);

    let table = run(
        &provider,
        &[
            "add",
            document.to_str().unwrap(),
            "/slide[1]",
            "--type",
            "table",
            "--rows",
            "2",
            "--columns",
            "2",
            "--json",
        ],
    );
    assert_eq!(table["data"]["operation"], "add-table");
    assert_eq!(table["data"]["path"], "/slide[1]/table[1]");
    assert_eq!(table["data"]["node"]["type"], "table");
    assert_eq!(table["data"]["node"]["childCount"], 2);

    for (path, text) in [
        ("/slide[1]/table[1]/tr[1]/tc[1]", "Name"),
        ("/slide[1]/table[1]/tr[1]/tc[2]", "Value"),
    ] {
        let updated = run(
            &provider,
            &[
                "set",
                document.to_str().unwrap(),
                path,
                "--text",
                text,
                "--json",
            ],
        );
        assert_eq!(updated["data"]["node"]["text"], text);
    }

    let row = run(
        &provider,
        &[
            "add",
            document.to_str().unwrap(),
            "/slide[1]/table[1]",
            "--type",
            "row",
            "--columns",
            "2",
            "--json",
        ],
    );
    assert_eq!(row["data"]["operation"], "add-table-row");
    assert_eq!(row["data"]["path"], "/slide[1]/table[1]/tr[3]");
    assert_eq!(row["data"]["node"]["childCount"], 2);

    let updated = run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/slide[1]/table[1]/tr[3]/tc[1]",
            "--text",
            "Added row",
            "--json",
        ],
    );
    assert_eq!(updated["data"]["node"]["text"], "Added row");

    let column = run(
        &provider,
        &[
            "add",
            document.to_str().unwrap(),
            "/slide[1]/table[1]",
            "--type",
            "column",
            "--index",
            "1",
            "--text",
            "Inserted",
            "--json",
        ],
    );
    assert_eq!(column["data"]["operation"], "add-table-column");
    assert_eq!(column["data"]["path"], "/slide[1]/table[1]/col[2]");
    assert_eq!(column["data"]["node"]["type"], "table-column");
    assert_eq!(column["data"]["node"]["format"]["widthEmu"], "4114800");

    let resized = run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/slide[1]/table[1]/col[2]",
            "--width-emu",
            "2000000",
            "--json",
        ],
    );
    assert_eq!(resized["data"]["operation"], "set-table-column-width");
    assert_eq!(resized["data"]["node"]["format"]["widthEmu"], "2000000");

    let inserted = run(
        &provider,
        &[
            "get",
            document.to_str().unwrap(),
            "/slide[1]/table[1]/tr[1]/tc[2]",
            "--json",
        ],
    );
    assert_eq!(inserted["data"]["node"]["text"], "Inserted");

    run(
        &provider,
        &[
            "remove",
            document.to_str().unwrap(),
            "/slide[1]/table[1]/col[2]",
            "--json",
        ],
    );

    let moved = run(
        &provider,
        &[
            "move",
            document.to_str().unwrap(),
            "/slide[1]/table[1]/col[1]",
            "--after",
            "/slide[1]/table[1]/col[2]",
            "--json",
        ],
    );
    assert_eq!(moved["data"]["resultPath"], "/slide[1]/table[1]/col[2]");
    let copied = run(
        &provider,
        &[
            "copy",
            document.to_str().unwrap(),
            "/slide[1]/table[1]/col[2]",
            "--json",
        ],
    );
    assert_eq!(copied["data"]["resultPath"], "/slide[1]/table[1]/col[3]");
    let swapped = run(
        &provider,
        &[
            "swap",
            document.to_str().unwrap(),
            "/slide[1]/table[1]/col[1]",
            "/slide[1]/table[1]/col[3]",
            "--json",
        ],
    );
    assert_eq!(
        swapped["data"]["result"]["first"],
        "/slide[1]/table[1]/col[3]"
    );
    assert_eq!(
        swapped["data"]["result"]["second"],
        "/slide[1]/table[1]/col[1]"
    );

    let full_row = run_failure(
        &provider,
        &[
            "add",
            document.to_str().unwrap(),
            "/slide[1]/table[1]/tr[3]",
            "--type",
            "cell",
            "--text",
            "overflow",
            "--json",
        ],
    );
    assert_eq!(
        full_row["error"]["code"],
        "use.office.presentation_table_cell_grid_full"
    );

    run(
        &provider,
        &[
            "remove",
            document.to_str().unwrap(),
            "/slide[1]/table[1]/tr[3]",
            "--json",
        ],
    );
    let table = run(
        &provider,
        &[
            "get",
            document.to_str().unwrap(),
            "/slide[1]/table[1]",
            "--depth",
            "3",
            "--json",
        ],
    );
    assert_eq!(table["data"]["node"]["childCount"], 2);
    assert_eq!(
        table["data"]["node"]["children"][0]["children"][0]["text"],
        "Name"
    );

    run(
        &provider,
        &[
            "remove",
            document.to_str().unwrap(),
            "/slide[1]/table[1]",
            "--json",
        ],
    );
    let stats = run(
        &provider,
        &["view", document.to_str().unwrap(), "stats", "--json"],
    );
    assert_eq!(stats["data"]["result"]["tableCount"], 0);
}

#[test]
fn native_cli_batches_presentation_tables_atomically_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("batch-table.pptx");
    let batch = temp.path().join("table-batch.json");
    let failed_batch = temp.path().join("failed-table-batch.json");
    create_presentation(&provider, &document);

    std::fs::write(
        &batch,
        serde_json::to_vec(&serde_json::json!({
            "schemaVersion": 1,
            "mutations": [
                {
                    "operation": "add-table",
                    "parent": "/slide[1]",
                    "rows": 2,
                    "columns": 2
                },
                {
                    "operation": "set-text",
                    "path": "/slide[1]/table[1]/tr[1]/tc[1]",
                    "text": "Name"
                },
                {
                    "operation": "set-text",
                    "path": "/slide[1]/table[1]/tr[1]/tc[2]",
                    "text": "Value"
                },
                {
                    "operation": "add-table-column",
                    "parent": "/slide[1]/table[1]",
                    "index": 1,
                    "text": "Inserted"
                },
                {
                    "operation": "set-table-column-width",
                    "path": "/slide[1]/table[1]/col[2]",
                    "widthEmu": 2000000
                }
            ]
        }))
        .unwrap(),
    )
    .unwrap();
    let result = run(
        &provider,
        &[
            "batch",
            document.to_str().unwrap(),
            "--input",
            batch.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(result["data"]["result"]["applied"], 5);
    assert_eq!(result["data"]["result"]["paths"][0], "/slide[1]/table[1]");
    assert_eq!(
        result["data"]["result"]["paths"][3],
        "/slide[1]/table[1]/col[2]"
    );

    std::fs::write(
        &failed_batch,
        serde_json::to_vec(&serde_json::json!({
            "schemaVersion": 1,
            "mutations": [
                {
                    "operation": "add-table-row",
                    "parent": "/slide[1]/table[1]"
                },
                {
                    "operation": "set-text",
                    "path": "/slide[1]/table[1]/tr[99]/tc[1]",
                    "text": "missing"
                }
            ]
        }))
        .unwrap(),
    )
    .unwrap();
    let before = std::fs::read(&document).unwrap();
    let failure = run_failure(
        &provider,
        &[
            "batch",
            document.to_str().unwrap(),
            "--input",
            failed_batch.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(failure["error"]["code"], "use.office.node_not_found");
    assert_eq!(std::fs::read(&document).unwrap(), before);
}
