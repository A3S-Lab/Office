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

fn raw(provider: &Path, document: &Path, part: &str) -> serde_json::Value {
    run(
        provider,
        &["raw", document.to_str().unwrap(), part, "--json"],
    )
}

#[test]
fn native_add_part_creates_known_word_spreadsheet_and_presentation_parts() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");

    let word = temp.path().join("parts.docx");
    create(&provider, &word);
    let header = run(
        &provider,
        &[
            "add-part",
            word.to_str().unwrap(),
            "/",
            "--type",
            "header",
            "--json",
        ],
    );
    assert_eq!(header["data"]["operation"], "add-part");
    assert_eq!(header["data"]["createdPart"]["path"], "/header[1]");
    assert_eq!(header["data"]["createdPart"]["part"], "/word/header1.xml");
    assert_eq!(header["data"]["createdPart"]["type"], "header");
    assert!(header["data"]["createdPart"]["relationshipId"]
        .as_str()
        .unwrap()
        .starts_with("rId"));
    let header_xml = raw(&provider, &word, "/word/header1.xml");
    assert_eq!(header_xml["data"]["part"]["root"]["localName"], "hdr");

    let word_chart = run(
        &provider,
        &[
            "add-part",
            word.to_str().unwrap(),
            "/",
            "--type",
            "chart",
            "--json",
        ],
    );
    assert_eq!(
        word_chart["data"]["createdPart"]["part"],
        "/word/charts/chart1.xml"
    );
    assert_eq!(
        raw(&provider, &word, "/word/charts/chart1.xml")["data"]["part"]["root"]["localName"],
        "chartSpace"
    );

    let spreadsheet = temp.path().join("parts.xlsx");
    create(&provider, &spreadsheet);
    let spreadsheet_chart = run(
        &provider,
        &[
            "add-part",
            spreadsheet.to_str().unwrap(),
            "/Sheet1",
            "--type",
            "chart",
            "--json",
        ],
    );
    assert_eq!(
        spreadsheet_chart["data"]["createdPart"]["ownerPart"],
        "/xl/drawings/drawing1.xml"
    );
    assert_eq!(
        spreadsheet_chart["data"]["createdPart"]["part"],
        "/xl/charts/chart1.xml"
    );
    assert!(
        raw(&provider, &spreadsheet, "/xl/worksheets/sheet1.xml")["data"]["xml"]
            .as_str()
            .unwrap()
            .contains("<drawing")
    );

    let presentation = temp.path().join("parts.pptx");
    create(&provider, &presentation);
    run(
        &provider,
        &[
            "add",
            presentation.to_str().unwrap(),
            "/",
            "--type",
            "slide",
            "--text",
            "Chart carrier",
            "--json",
        ],
    );
    let presentation_chart = run(
        &provider,
        &[
            "add-part",
            presentation.to_str().unwrap(),
            "/slide[1]",
            "--type",
            "chart",
            "--json",
        ],
    );
    assert_eq!(
        presentation_chart["data"]["createdPart"]["path"],
        "/slide[1]/chart[1]"
    );
    assert_eq!(
        presentation_chart["data"]["createdPart"]["ownerPart"],
        "/ppt/slides/slide1.xml"
    );
}

#[test]
fn native_add_part_batch_returns_receipts_and_rolls_back_on_failure() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let word = temp.path().join("batch.docx");
    let word_batch = temp.path().join("word-batch.json");
    create(&provider, &word);
    std::fs::write(
        &word_batch,
        serde_json::to_vec(&serde_json::json!({
            "schemaVersion": 1,
            "mutations": [
                {"operation": "add-part", "parent": "/", "type": "header"},
                {"operation": "add-part", "parent": "/", "type": "footer"}
            ]
        }))
        .unwrap(),
    )
    .unwrap();
    let result = run(
        &provider,
        &[
            "batch",
            word.to_str().unwrap(),
            "--input",
            word_batch.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(result["data"]["result"]["applied"], 2);
    assert_eq!(
        result["data"]["result"]["createdParts"]
            .as_array()
            .unwrap()
            .len(),
        2
    );
    assert_eq!(
        result["data"]["result"]["createdParts"][1]["part"],
        "/word/footer1.xml"
    );

    let spreadsheet = temp.path().join("rollback.xlsx");
    let failing_batch = temp.path().join("failing-batch.json");
    create(&provider, &spreadsheet);
    let original = std::fs::read(&spreadsheet).unwrap();
    std::fs::write(
        &failing_batch,
        serde_json::to_vec(&serde_json::json!({
            "schemaVersion": 1,
            "mutations": [
                {"operation": "add-part", "parent": "/Sheet1", "type": "chart"},
                {"operation": "add-part", "parent": "/Sheet1", "type": "header"}
            ]
        }))
        .unwrap(),
    )
    .unwrap();
    let failure = run_failure(
        &provider,
        &[
            "batch",
            spreadsheet.to_str().unwrap(),
            "--input",
            failing_batch.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(failure["error"]["code"], "use.office.part_type_unsupported");
    assert_eq!(std::fs::read(&spreadsheet).unwrap(), original);
}
