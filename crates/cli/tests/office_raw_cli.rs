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

fn word_document(text: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>{text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#
    )
}

fn text_view(provider: &Path, document: &Path) -> String {
    let result = run(
        provider,
        &["view", document.to_str().unwrap(), "text", "--json"],
    );
    result["data"]["result"]["text"]
        .as_str()
        .unwrap()
        .to_string()
}

#[test]
fn native_raw_cli_inspects_exports_and_replaces_xml_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let source = temp.path().join("source.docx");
    let changed = temp.path().join("changed.docx");
    let xml_input = temp.path().join("document.xml");
    let exported = temp.path().join("exported.xml");
    let provider = temp.path().join("must-not-be-invoked");

    run(&provider, &["create", source.to_str().unwrap(), "--json"]);
    let source_before = std::fs::read(&source).unwrap();
    let inspected = run(
        &provider,
        &[
            "raw",
            source.to_str().unwrap(),
            "/word/document.xml",
            "--json",
        ],
    );
    assert_eq!(inspected["data"]["operation"], "raw");
    assert_eq!(inspected["data"]["part"]["name"], "/word/document.xml");
    assert_eq!(inspected["data"]["part"]["root"]["localName"], "document");
    assert_eq!(
        inspected["data"]["part"]["sha256"].as_str().unwrap().len(),
        64
    );
    assert!(inspected["data"]["xml"]
        .as_str()
        .unwrap()
        .contains("<w:body>"));

    let export = run(
        &provider,
        &[
            "raw",
            source.to_str().unwrap(),
            "word/document.xml",
            "--output",
            exported.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(export["data"]["exported"], true);
    assert_eq!(
        std::fs::read_to_string(&exported).unwrap(),
        inspected["data"]["xml"].as_str().unwrap()
    );
    assert_eq!(std::fs::read(&source).unwrap(), source_before);

    let no_clobber = run_failure(
        &provider,
        &[
            "raw",
            source.to_str().unwrap(),
            "/word/document.xml",
            "--output",
            source.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(no_clobber["error"]["code"], "use.office.raw_output_exists");
    assert_eq!(std::fs::read(&source).unwrap(), source_before);

    std::fs::write(&xml_input, word_document("Native raw CLI")).unwrap();
    let replaced = run(
        &provider,
        &[
            "raw-set",
            source.to_str().unwrap(),
            "/word/document.xml",
            "--input",
            xml_input.to_str().unwrap(),
            "--output",
            changed.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(replaced["data"]["operation"], "raw-set");
    assert_eq!(replaced["data"]["part"], "/word/document.xml");
    assert_eq!(replaced["data"]["inPlace"], false);
    assert_eq!(text_view(&provider, &source), "");
    assert_eq!(text_view(&provider, &changed), "Native raw CLI");
    assert_eq!(std::fs::read(&source).unwrap(), source_before);
}

#[test]
fn native_raw_cli_rejects_unsafe_replacements_and_rolls_back_batches() {
    let temp = tempfile::tempdir().unwrap();
    let document = temp.path().join("safe.docx");
    let xml_input = temp.path().join("replacement.xml");
    let mutations = temp.path().join("mutations.json");
    let provider = temp.path().join("must-not-be-invoked");
    let document_text = document.to_str().unwrap();

    run(&provider, &["create", document_text, "--json"]);
    let original = std::fs::read(&document).unwrap();

    std::fs::write(&xml_input, "<document/>").unwrap();
    let mismatch = run_failure(
        &provider,
        &[
            "raw-set",
            document_text,
            "/word/document.xml",
            "--input",
            xml_input.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(mismatch["error"]["code"], "use.office.raw_root_mismatch");
    assert_eq!(std::fs::read(&document).unwrap(), original);

    std::fs::write(&xml_input, "<Types/>").unwrap();
    let protected = run_failure(
        &provider,
        &[
            "raw-set",
            document_text,
            "/[Content_Types].xml",
            "--input",
            xml_input.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(protected["error"]["code"], "use.office.raw_part_protected");
    assert_eq!(std::fs::read(&document).unwrap(), original);

    std::fs::write(
        &mutations,
        serde_json::to_vec(&serde_json::json!({
            "schemaVersion": 1,
            "mutations": [
                {
                    "operation": "replace-xml-part",
                    "part": "/word/document.xml",
                    "xml": word_document("must roll back")
                },
                {
                    "operation": "replace-xml-part",
                    "part": "/_rels/.rels",
                    "xml": "<Relationships/>"
                }
            ]
        }))
        .unwrap(),
    )
    .unwrap();
    let failed_batch = run_failure(
        &provider,
        &[
            "batch",
            document_text,
            "--input",
            mutations.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(
        failed_batch["error"]["code"],
        "use.office.raw_part_protected"
    );
    assert_eq!(std::fs::read(&document).unwrap(), original);
    assert_eq!(text_view(&provider, &document), "");
}
