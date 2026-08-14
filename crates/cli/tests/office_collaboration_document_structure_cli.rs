mod support;

use std::fs;
use std::process::{Command, Output};

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use support::{complex_document_fixture, document_snapshot, DocumentParagraph, DocumentRow};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

fn run(args: &[&str]) -> serde_json::Value {
    let output: Output = Command::new(binary()).args(args).output().unwrap();
    assert!(output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

#[test]
fn cli_mutates_paragraphs_in_nested_list_and_table_containers() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("complex-document.replica");
    let fixture = temp.path().join("complex-document.update");
    fs::write(&fixture, complex_document_fixture()).unwrap();
    run(&[
        "collab",
        "join",
        replica.to_str().unwrap(),
        "--artifact-id",
        "fixture-complex-document",
        "--kind",
        "document",
        "--actor-id",
        "coding-agent-complex-cli",
        "--actor-kind",
        "agent",
        "--mode",
        "edit",
        "--operation-id",
        "join-complex-document-cli",
        "--input",
        fixture.to_str().unwrap(),
        "--client-id",
        "900011",
        "--json",
    ]);

    let list_inserted = mutate(
        &replica,
        "cli-insert-list-paragraph",
        r#"{"type":"document-insert-paragraph","anchorParagraphId":"00000101","position":"after","paragraphId":"00000105","textId":"00000106","text":"CLI list"}"#,
    );
    assert_eq!(list_inserted["data"]["sequence"], 1);
    let cell_inserted = mutate(
        &replica,
        "cli-insert-nested-cell-paragraph",
        r#"{"type":"document-insert-paragraph","anchorParagraphId":"00000311","position":"after","paragraphId":"00000315","textId":"00000316","text":"CLI cell"}"#,
    );
    assert_eq!(cell_inserted["data"]["sequence"], 2);

    let exported = run(&["collab", "diff", replica.to_str().unwrap(), "--json"]);
    let update = STANDARD
        .decode(exported["data"]["updateBase64"].as_str().unwrap())
        .unwrap();
    let snapshot = document_snapshot(&update);
    assert!(snapshot
        .paragraphs
        .contains(&paragraph("00000105", "00000106", "listItem", "CLI list",)));
    assert!(snapshot.paragraphs.contains(&paragraph(
        "00000315",
        "00000316",
        "tableCell",
        "CLI cell",
    )));
    assert_eq!(
        snapshot.rows,
        vec![row("00000201", "00000203"), row("00000301", "00000303")]
    );
}

fn mutate(replica: &std::path::Path, operation_id: &str, mutation: &str) -> serde_json::Value {
    run(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        mutation,
        "--actor-id",
        "coding-agent-complex-cli",
        "--operation-id",
        operation_id,
        "--artifact-id",
        "fixture-complex-document",
        "--kind",
        "document",
        "--mode",
        "edit",
        "--json",
    ])
}

fn paragraph(paragraph_id: &str, text_id: &str, parent_tag: &str, text: &str) -> DocumentParagraph {
    DocumentParagraph {
        paragraph_id: paragraph_id.to_owned(),
        text_id: text_id.to_owned(),
        parent_tag: parent_tag.to_owned(),
        text: text.to_owned(),
    }
}

fn row(row_id: &str, row_text_id: &str) -> DocumentRow {
    DocumentRow {
        row_id: row_id.to_owned(),
        row_text_id: row_text_id.to_owned(),
    }
}
