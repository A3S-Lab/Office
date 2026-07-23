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
fn native_cli_formats_word_spreadsheet_and_presentation_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let word = temp.path().join("formatted.docx");
    let spreadsheet = temp.path().join("formatted.xlsx");
    let presentation = temp.path().join("formatted.pptx");
    let word = word.to_str().unwrap();
    let spreadsheet = spreadsheet.to_str().unwrap();
    let presentation = presentation.to_str().unwrap();

    run(&provider, &["create", word, "--json"]);
    run(
        &provider,
        &["set", word, "/body/p[1]", "--text", "Native Word", "--json"],
    );
    let word_run = run(
        &provider,
        &[
            "set",
            word,
            "/body/p[1]/r[1]",
            "--bold",
            "true",
            "--italic",
            "false",
            "--underline",
            "double",
            "--script",
            "superscript",
            "--strikethrough",
            "true",
            "--double-strikethrough",
            "true",
            "--text-case",
            "small-caps",
            "--highlight",
            "yellow",
            "--language",
            "en-US",
            "--font-family",
            "Aptos",
            "--font-size",
            "14pt",
            "--text-color",
            "#123456",
            "--json",
        ],
    );
    assert_eq!(word_run["data"]["operation"], "set-text-format");
    assert_eq!(word_run["data"]["node"]["format"]["bold"], "true");
    assert_eq!(word_run["data"]["node"]["format"]["underline"], "double");
    assert_eq!(word_run["data"]["node"]["format"]["script"], "superscript");
    assert_eq!(word_run["data"]["node"]["format"]["strike"], "true");
    assert_eq!(word_run["data"]["node"]["format"]["doubleStrike"], "true");
    assert_eq!(word_run["data"]["node"]["format"]["textCase"], "small-caps");
    assert_eq!(word_run["data"]["node"]["format"]["highlight"], "yellow");
    assert_eq!(word_run["data"]["node"]["format"]["language"], "en-US");
    assert_eq!(word_run["data"]["node"]["format"]["font"], "Aptos");
    assert_eq!(word_run["data"]["node"]["format"]["size"], "14pt");
    assert_eq!(word_run["data"]["node"]["format"]["color"], "123456");
    let word_paragraph = run(
        &provider,
        &["set", word, "/body/p[1]", "--align", "center", "--json"],
    );
    assert_eq!(
        word_paragraph["data"]["node"]["format"]["alignment"],
        "center"
    );

    run(&provider, &["create", spreadsheet, "--json"]);
    let cell = run(
        &provider,
        &[
            "set",
            spreadsheet,
            "/Sheet1/A1",
            "--text",
            "Revenue",
            "--bold",
            "true",
            "--underline",
            "single",
            "--script",
            "subscript",
            "--strikethrough",
            "false",
            "--font-size",
            "11.5",
            "--text-color",
            "0066CC",
            "--align",
            "right",
            "--json",
        ],
    );
    assert_eq!(cell["data"]["operation"], "set-content-and-text-format");
    assert_eq!(cell["data"]["node"]["text"], "Revenue");
    assert_eq!(cell["data"]["node"]["format"]["bold"], "true");
    assert_eq!(cell["data"]["node"]["format"]["underline"], "single");
    assert_eq!(cell["data"]["node"]["format"]["script"], "subscript");
    assert_eq!(cell["data"]["node"]["format"]["strike"], "false");
    assert_eq!(cell["data"]["node"]["format"]["size"], "11.5pt");
    assert_eq!(cell["data"]["node"]["format"]["color"], "0066CC");
    assert_eq!(cell["data"]["node"]["format"]["alignment"], "right");

    run(&provider, &["create", presentation, "--json"]);
    run(
        &provider,
        &[
            "add",
            presentation,
            "/",
            "--type",
            "slide",
            "--text",
            "Native Slides",
            "--json",
        ],
    );
    let slide_run = run(
        &provider,
        &[
            "set",
            presentation,
            "/slide[1]/shape[1]/paragraph[1]/run[1]",
            "--italic",
            "true",
            "--underline",
            "double",
            "--script",
            "superscript",
            "--text-case",
            "all-caps",
            "--highlight",
            "cyan",
            "--language",
            "zh-CN",
            "--font-family",
            "Aptos Display",
            "--font-size",
            "20",
            "--text-color",
            "AA2200",
            "--json",
        ],
    );
    assert_eq!(slide_run["data"]["node"]["format"]["italic"], "1");
    assert_eq!(slide_run["data"]["node"]["format"]["underline"], "double");
    assert_eq!(slide_run["data"]["node"]["format"]["script"], "superscript");
    assert_eq!(slide_run["data"]["node"]["format"]["textCase"], "all-caps");
    assert_eq!(slide_run["data"]["node"]["format"]["highlight"], "cyan");
    assert_eq!(slide_run["data"]["node"]["format"]["language"], "zh-CN");
    assert_eq!(slide_run["data"]["node"]["format"]["font"], "Aptos Display");
    assert_eq!(slide_run["data"]["node"]["format"]["size"], "20pt");
    assert_eq!(slide_run["data"]["node"]["format"]["color"], "AA2200");
    let slide_paragraph = run(
        &provider,
        &[
            "set",
            presentation,
            "/slide[1]/shape[1]/paragraph[1]",
            "--align",
            "justify",
            "--json",
        ],
    );
    assert_eq!(
        slide_paragraph["data"]["node"]["format"]["alignment"],
        "just"
    );

    for document in [word, spreadsheet, presentation] {
        assert_eq!(
            run(&provider, &["validate", document, "--json"])["data"]["valid"],
            true
        );
    }
    assert!(!provider.exists());
}

#[test]
fn native_format_batch_is_typed_and_invalid_word_sizes_do_not_write() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let spreadsheet = temp.path().join("batch.xlsx");
    let batch = temp.path().join("format.json");
    let spreadsheet = spreadsheet.to_str().unwrap();
    run(&provider, &["create", spreadsheet, "--json"]);
    std::fs::write(
        &batch,
        serde_json::to_vec(&serde_json::json!({
            "schemaVersion": 1,
            "mutations": [{
                "operation": "set-text-format",
                "path": "/Sheet1/B2",
                "format": {
                    "bold": true,
                    "underline": "double",
                    "script": "subscript",
                    "strikethrough": true,
                    "fontSizeCentipoints": 1200,
                    "textColor": { "red": 17, "green": 34, "blue": 51 },
                    "alignment": "center"
                }
            }]
        }))
        .unwrap(),
    )
    .unwrap();
    let batch_path = batch.to_str().unwrap();
    let result = run(
        &provider,
        &["batch", spreadsheet, "--input", batch_path, "--json"],
    );
    assert_eq!(result["data"]["result"]["applied"], 1);
    let cell = run(&provider, &["get", spreadsheet, "/Sheet1/B2", "--json"]);
    assert_eq!(cell["data"]["node"]["format"]["bold"], "true");
    assert_eq!(cell["data"]["node"]["format"]["underline"], "double");
    assert_eq!(cell["data"]["node"]["format"]["script"], "subscript");
    assert_eq!(cell["data"]["node"]["format"]["strike"], "true");
    assert_eq!(cell["data"]["node"]["format"]["color"], "112233");

    let before = std::fs::read(spreadsheet).unwrap();
    let error = run_failure(
        &provider,
        &[
            "set",
            spreadsheet,
            "/Sheet1/B2",
            "--highlight",
            "yellow",
            "--json",
        ],
    );
    assert_eq!(
        error["error"]["code"],
        "use.office.spreadsheet_run_format_unsupported"
    );
    assert_eq!(std::fs::read(spreadsheet).unwrap(), before);

    let word = temp.path().join("invalid.docx");
    let word = word.to_str().unwrap();
    run(&provider, &["create", word, "--json"]);
    run(
        &provider,
        &["set", word, "/body/p[1]", "--text", "Unchanged", "--json"],
    );
    let before = std::fs::read(word).unwrap();
    let error = run_failure(
        &provider,
        &[
            "set",
            word,
            "/body/p[1]/r[1]",
            "--font-size",
            "11.25",
            "--json",
        ],
    );
    assert_eq!(error["error"]["code"], "use.office.font_size_unsupported");
    assert_eq!(std::fs::read(word).unwrap(), before);

    let presentation = temp.path().join("unsupported-strike.pptx");
    let presentation = presentation.to_str().unwrap();
    run(&provider, &["create", presentation, "--json"]);
    run(
        &provider,
        &[
            "add",
            presentation,
            "/",
            "--type",
            "slide",
            "--text",
            "Unchanged",
            "--json",
        ],
    );
    let before = std::fs::read(presentation).unwrap();
    let error = run_failure(
        &provider,
        &[
            "set",
            presentation,
            "/slide[1]/shape[1]/paragraph[1]/run[1]",
            "--strikethrough",
            "true",
            "--json",
        ],
    );
    assert_eq!(
        error["error"]["code"],
        "use.office.presentation_strikethrough_unsupported"
    );
    assert_eq!(std::fs::read(presentation).unwrap(), before);
    let error = run_failure(
        &provider,
        &[
            "set",
            presentation,
            "/slide[1]/shape[1]/paragraph[1]/run[1]",
            "--double-strikethrough",
            "true",
            "--json",
        ],
    );
    assert_eq!(
        error["error"]["code"],
        "use.office.presentation_double_strikethrough_unsupported"
    );
    assert_eq!(std::fs::read(presentation).unwrap(), before);
    assert!(!provider.exists());
}
