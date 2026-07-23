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

fn create(provider: &Path, document: &Path) {
    run(provider, &["create", document.to_str().unwrap(), "--json"]);
}

#[test]
fn native_cli_combines_content_text_and_cell_format_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("formatted.xlsx");
    create(&provider, &document);

    let result = run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/Sheet1/A1",
            "--number",
            "1234.5",
            "--bold",
            "true",
            "--text-color",
            "123456",
            "--align",
            "right",
            "--number-format",
            "currency",
            "--fill",
            "AABBCC",
            "--border-all",
            "thin",
            "--border-color",
            "112233",
            "--border-right",
            "medium-dashed",
            "--border-right-color",
            "445566",
            "--border-bottom",
            "none",
            "--border-diagonal",
            "slant-dash-dot",
            "--border-diagonal-color",
            "778899",
            "--border-diagonal-up",
            "true",
            "--border-diagonal-down",
            "false",
            "--vertical-align",
            "distributed",
            "--wrap-text",
            "true",
            "--text-rotation",
            "45",
            "--indent",
            "2",
            "--shrink-to-fit",
            "false",
            "--reading-order",
            "rtl",
            "--json",
        ],
    );

    assert_eq!(
        result["data"]["operation"],
        "set-content-and-text-and-cell-format"
    );
    let node = &result["data"]["node"];
    assert_eq!(node["text"], "1234.5");
    assert_eq!(node["format"]["bold"], "true");
    assert_eq!(node["format"]["color"], "123456");
    assert_eq!(node["format"]["alignment"], "right");
    assert_eq!(node["format"]["numberFormat"], "\"$\"#,##0.00");
    assert_eq!(node["format"]["fill"], "AABBCC");
    assert_eq!(node["format"]["borderLeft"], "thin");
    assert_eq!(node["format"]["borderLeftColor"], "112233");
    assert_eq!(node["format"]["borderRight"], "mediumDashed");
    assert_eq!(node["format"]["borderRightColor"], "445566");
    assert_eq!(node["format"]["borderTop"], "thin");
    assert_eq!(node["format"]["borderTopColor"], "112233");
    assert!(node["format"].get("borderBottom").is_none());
    assert_eq!(node["format"]["borderDiagonal"], "slantDashDot");
    assert_eq!(node["format"]["borderDiagonalColor"], "778899");
    assert_eq!(node["format"]["borderDiagonalUp"], "true");
    assert_eq!(node["format"]["borderDiagonalDown"], "false");
    assert_eq!(node["format"]["verticalAlignment"], "distributed");
    assert_eq!(node["format"]["wrapText"], "true");
    assert_eq!(node["format"]["textRotation"], "45");
    assert_eq!(node["format"]["indent"], "2");
    assert_eq!(node["format"]["shrinkToFit"], "false");
    assert_eq!(node["format"]["readingOrder"], "rtl");
    assert!(!provider.exists());
}

#[test]
fn native_cell_format_batch_uses_the_typed_json_contract() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("batch.xlsx");
    let mutations = temp.path().join("mutations.json");
    create(&provider, &document);
    std::fs::write(
        &mutations,
        serde_json::to_vec(&serde_json::json!({
            "schemaVersion": 1,
            "mutations": [
                {
                    "operation": "set-cell-value",
                    "path": "/Sheet1/C3",
                    "value": { "type": "number", "value": "0.125" }
                },
                {
                    "operation": "set-text-format",
                    "path": "/Sheet1/C3",
                    "format": { "italic": true }
                },
                {
                    "operation": "set-cell-format",
                    "path": "/Sheet1/C3",
                    "format": {
                        "numberFormat": "percent",
                        "fill": {
                            "kind": "solid",
                            "color": { "red": 1, "green": 2, "blue": 3 }
                        },
                        "border": {
                            "left": {
                                "kind": "line",
                                "style": "dashDot",
                                "color": { "red": 4, "green": 5, "blue": 6 }
                            },
                            "bottom": { "kind": "none" },
                            "diagonalUp": false
                        },
                        "verticalAlignment": "center",
                        "wrapText": false,
                        "textRotation": 255,
                        "indent": 1,
                        "shrinkToFit": true,
                        "readingOrder": "left-to-right"
                    }
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
            mutations.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(result["data"]["result"]["applied"], 3);
    assert_eq!(result["data"]["atomic"], true);

    let result = run(
        &provider,
        &["get", document.to_str().unwrap(), "/Sheet1/C3", "--json"],
    );
    let node = &result["data"]["node"];
    assert_eq!(node["text"], "0.125");
    assert_eq!(node["format"]["italic"], "true");
    assert_eq!(node["format"]["numberFormat"], "0.00%");
    assert_eq!(node["format"]["fill"], "010203");
    assert_eq!(node["format"]["borderLeft"], "dashDot");
    assert_eq!(node["format"]["borderLeftColor"], "040506");
    assert!(node["format"].get("borderBottom").is_none());
    assert_eq!(node["format"]["borderDiagonalUp"], "false");
    assert_eq!(node["format"]["verticalAlignment"], "center");
    assert_eq!(node["format"]["wrapText"], "false");
    assert_eq!(node["format"]["textRotation"], "255");
    assert_eq!(node["format"]["indent"], "1");
    assert_eq!(node["format"]["shrinkToFit"], "true");
    assert_eq!(node["format"]["readingOrder"], "ltr");
    assert!(!provider.exists());
}

#[test]
fn invalid_or_wrong_document_cell_format_never_writes() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let spreadsheet = temp.path().join("unchanged.xlsx");
    create(&provider, &spreadsheet);
    let before = std::fs::read(&spreadsheet).unwrap();

    let invalid_number_format = run_failure(
        &provider,
        &[
            "set",
            spreadsheet.to_str().unwrap(),
            "/Sheet1/A1",
            "--text",
            "must roll back",
            "--number-format",
            "0;0;0;0;0",
            "--json",
        ],
    );
    assert_eq!(
        invalid_number_format["error"]["code"],
        "use.office.number_format_invalid"
    );
    assert_eq!(std::fs::read(&spreadsheet).unwrap(), before);

    let invalid_rotation = run_failure(
        &provider,
        &[
            "set",
            spreadsheet.to_str().unwrap(),
            "/Sheet1/A1",
            "--text-rotation",
            "181",
            "--json",
        ],
    );
    assert_eq!(
        invalid_rotation["error"]["code"],
        "use.office.text_rotation_invalid"
    );
    assert_eq!(std::fs::read(&spreadsheet).unwrap(), before);

    let invalid_border = run_failure(
        &provider,
        &[
            "set",
            spreadsheet.to_str().unwrap(),
            "/Sheet1/A1",
            "--border-top",
            "triple",
            "--json",
        ],
    );
    assert_eq!(invalid_border["error"]["code"], "use.cli.invalid_usage");
    assert_eq!(std::fs::read(&spreadsheet).unwrap(), before);

    let orphan_border_color = run_failure(
        &provider,
        &[
            "set",
            spreadsheet.to_str().unwrap(),
            "/Sheet1/A1",
            "--border-top-color",
            "FF0000",
            "--json",
        ],
    );
    assert_eq!(
        orphan_border_color["error"]["code"],
        "use.cli.invalid_usage"
    );
    assert_eq!(std::fs::read(&spreadsheet).unwrap(), before);

    for (name, path) in [("word.docx", "/body/p[1]"), ("slides.pptx", "/slide[1]")] {
        let document = temp.path().join(name);
        create(&provider, &document);
        let before = std::fs::read(&document).unwrap();
        let error = run_failure(
            &provider,
            &[
                "set",
                document.to_str().unwrap(),
                path,
                "--wrap-text",
                "true",
                "--json",
            ],
        );
        assert_eq!(
            error["error"]["code"],
            "use.office.mutation_type_unsupported"
        );
        assert_eq!(std::fs::read(&document).unwrap(), before);
    }
    assert!(!provider.exists());
}
