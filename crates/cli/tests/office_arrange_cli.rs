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
fn native_move_copy_and_swap_cover_all_formats_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let word = temp.path().join("arrange.docx");
    let spreadsheet = temp.path().join("arrange.xlsx");
    let presentation = temp.path().join("arrange.pptx");
    let word = word.to_str().unwrap();
    let spreadsheet = spreadsheet.to_str().unwrap();
    let presentation = presentation.to_str().unwrap();

    run(&provider, &["create", word, "--json"]);
    run(
        &provider,
        &["set", word, "/body/p[1]", "--text", "A", "--json"],
    );
    for text in ["B", "C"] {
        run(
            &provider,
            &[
                "add",
                word,
                "/body",
                "--type",
                "paragraph",
                "--text",
                text,
                "--json",
            ],
        );
    }
    let moved = run(
        &provider,
        &[
            "move",
            word,
            "/body/p[1]",
            "--after",
            "/body/p[2]",
            "--json",
        ],
    );
    assert_eq!(moved["data"]["resultPath"], "/body/p[2]");
    run(&provider, &["copy", word, "/body/p[2]", "--json"]);
    run(
        &provider,
        &["swap", word, "/body/p[1]", "/body/p[4]", "--json"],
    );
    assert_eq!(text_view(&provider, word), "C\nA\nA\nB");

    run(&provider, &["create", spreadsheet, "--json"]);
    for (cell, text) in [("A1", "A"), ("A2", "B"), ("A3", "C")] {
        let path = format!("/Sheet1/{cell}");
        run(
            &provider,
            &["set", spreadsheet, &path, "--text", text, "--json"],
        );
    }
    run(
        &provider,
        &[
            "move",
            spreadsheet,
            "/Sheet1/row[1]",
            "--after",
            "/Sheet1/row[2]",
            "--json",
        ],
    );
    run(
        &provider,
        &["copy", spreadsheet, "/Sheet1/row[2]", "--json"],
    );
    run(
        &provider,
        &[
            "swap",
            spreadsheet,
            "/Sheet1/row[1]",
            "/Sheet1/row[4]",
            "--json",
        ],
    );
    assert_eq!(
        text_view(&provider, spreadsheet),
        "/Sheet1/A1=C\n/Sheet1/A2=A\n/Sheet1/A3=A\n/Sheet1/A4=B"
    );

    run(&provider, &["create", presentation, "--json"]);
    for title in ["A", "B", "C"] {
        run(
            &provider,
            &[
                "add",
                presentation,
                "/",
                "--type",
                "slide",
                "--text",
                title,
                "--json",
            ],
        );
    }
    run(
        &provider,
        &[
            "move",
            presentation,
            "/slide[1]",
            "--after",
            "/slide[2]",
            "--json",
        ],
    );
    run(&provider, &["copy", presentation, "/slide[2]", "--json"]);
    run(
        &provider,
        &["swap", presentation, "/slide[1]", "/slide[4]", "--json"],
    );
    assert_eq!(text_view(&provider, presentation), "C\nA\nA\nB");

    for document in [word, spreadsheet, presentation] {
        assert_eq!(
            run(&provider, &["validate", document, "--json"])["data"]["valid"],
            true
        );
    }
}

#[test]
fn native_arrangement_rejects_ambiguous_positions_and_unsafe_rows_without_writes() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("safe.xlsx");
    let document = document.to_str().unwrap();
    run(&provider, &["create", document, "--json"]);
    for (path, value) in [("/Sheet1/A1", "A"), ("/Sheet1/A2", "B")] {
        run(
            &provider,
            &["set", document, path, "--text", value, "--json"],
        );
    }

    let ambiguous = run_failure(
        &provider,
        &[
            "move",
            document,
            "/Sheet1/row[1]",
            "--index",
            "0",
            "--after",
            "/Sheet1/row[2]",
            "--json",
        ],
    );
    assert_eq!(ambiguous["error"]["code"], "use.cli.invalid_usage");

    let zero_based = run(
        &provider,
        &["move", document, "/Sheet1/row[2]", "--index", "0", "--json"],
    );
    assert_eq!(zero_based["data"]["resultPath"], "/Sheet1/row[1]");
    assert_eq!(text_view(&provider, document), "/Sheet1/A1=B\n/Sheet1/A2=A");

    run(
        &provider,
        &["set", document, "/Sheet1/B1", "--formula", "A2", "--json"],
    );
    let before = std::fs::read(document).unwrap();
    let unsafe_move = run_failure(
        &provider,
        &[
            "move",
            document,
            "/Sheet1/row[1]",
            "--after",
            "/Sheet1/row[2]",
            "--json",
        ],
    );
    assert_eq!(
        unsafe_move["error"]["code"],
        "use.office.spreadsheet_row_arrange_unsupported"
    );
    assert_eq!(std::fs::read(document).unwrap(), before);
}

#[test]
fn native_arrangement_mutations_deserialize_and_report_swaps_in_atomic_batches() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("batch.docx");
    let batch = temp.path().join("arrange.json");
    let document = document.to_str().unwrap();

    run(&provider, &["create", document, "--json"]);
    run(
        &provider,
        &["set", document, "/body/p[1]", "--text", "A", "--json"],
    );
    for text in ["B", "C"] {
        run(
            &provider,
            &[
                "add",
                document,
                "/body",
                "--type",
                "paragraph",
                "--text",
                text,
                "--json",
            ],
        );
    }
    std::fs::write(
        &batch,
        serde_json::to_vec(&serde_json::json!({
            "schemaVersion": 1,
            "mutations": [
                {
                    "operation": "move",
                    "path": "/body/p[1]",
                    "position": { "kind": "after", "path": "/body/p[2]" }
                },
                {
                    "operation": "copy",
                    "path": "/body/p[2]"
                },
                {
                    "operation": "swap",
                    "path": "/body/p[1]",
                    "with": "/body/p[4]"
                }
            ]
        }))
        .unwrap(),
    )
    .unwrap();

    let applied = run(
        &provider,
        &[
            "batch",
            document,
            "--input",
            batch.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(applied["data"]["result"]["applied"], 3);
    assert_eq!(
        applied["data"]["result"]["paths"],
        serde_json::json!(["/body/p[2]", "/body/p[3]", "/body/p[4]"])
    );
    assert_eq!(
        applied["data"]["result"]["swaps"],
        serde_json::json!([{ "first": "/body/p[4]", "second": "/body/p[1]" }])
    );
    assert_eq!(text_view(&provider, document), "C\nA\nA\nB");
}

fn text_view(provider: &Path, document: &str) -> String {
    let value = run(provider, &["view", document, "text", "--json"]);
    value["data"]["result"]["text"]
        .as_str()
        .unwrap()
        .to_string()
}
