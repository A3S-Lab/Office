mod support;

use std::fs;
use std::path::Path;
use std::process::{Command, Output};

use base64::Engine as _;
use serde_json::{json, Value};
use support::{
    pdf_collaboration_fixture, presentation_collaboration_fixture, presentation_scene_element,
    spreadsheet_collaboration_fixture,
};

const YJS_MARKDOWN_UPDATE_BASE64: &str = "AQey8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtbWFya2Rvd24oARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhtYXJrZG93bigBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjQyNDI0Mjpicm93c2VyLWZpeHR1cmUEARphM3Mub2ZmaWNlLm1hcmtkb3duLnNvdXJjZRUjIFNoYXJlZAoKWWpzIHRvIFlycy4A";
const YJS_DOCUMENT_UPDATE_BASE64: &str = "AQ+y8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtZG9jdW1lbnQoARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhkb2N1bWVudCgBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjQyNDI0Mjpicm93c2VyLWZpeHR1cmUHARthM3Mub2ZmaWNlLmRvY3VtZW50LmNvbnRlbnQDD2RvY3VtZW50U2VjdGlvbgcAsvIZBgMJcGFyYWdyYXBoBwCy8hkHBgQAsvIZCBBIZWxsbyDwn5iAIHdvcmxkKACy8hkHC3BhcmFncmFwaElkAXcIMDAwMDAwMDEoALLyGQcGdGV4dElkAXcIMDAwMDAwMDIoALLyGQYCaWQBdxJkb2N1bWVudC1zZWN0aW9uLTEoARthM3Mub2ZmaWNlLmRvY3VtZW50Lm9wdGlvbnMJcGFnZUNvbG9yAXcHI0Y4RkFGQygBG2Ezcy5vZmZpY2UuZG9jdW1lbnQub3B0aW9ucwx0cmFja0NoYW5nZXMBeAA=";

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

fn execute(args: &[String]) -> Output {
    Command::new(binary()).args(args).output().unwrap()
}

fn run(args: &[String]) -> Value {
    let output = execute(args);
    let value = parse_output(&output);
    assert!(output.status.success(), "{value}");
    value
}

fn run_failure(args: &[String]) -> Value {
    let output = execute(args);
    let value = parse_output(&output);
    assert!(!output.status.success(), "{value}");
    value
}

fn parse_output(output: &Output) -> Value {
    serde_json::from_slice(&output.stdout).unwrap_or_else(|_| {
        panic!(
            "collaboration command did not return JSON:\nstdout: {}\nstderr: {}",
            String::from_utf8_lossy(&output.stdout),
            String::from_utf8_lossy(&output.stderr)
        )
    })
}

fn join(replica: &Path, artifact: &str, kind: &str, actor: &str, client_id: &str, update: &[u8]) {
    let fixture = replica.with_extension("update");
    fs::write(&fixture, update).unwrap();
    run(&args([
        "collab",
        "join",
        replica.to_str().unwrap(),
        "--artifact-id",
        artifact,
        "--kind",
        kind,
        "--actor-id",
        actor,
        "--actor-kind",
        "agent",
        "--mode",
        "edit",
        "--operation-id",
        &format!("join-{kind}-{client_id}"),
        "--input",
        fixture.to_str().unwrap(),
        "--client-id",
        client_id,
        "--json",
    ]));
}

fn read_replica(replica: &Path) -> Value {
    let value = run(&args([
        "collab",
        "read",
        replica.to_str().unwrap(),
        "--json",
    ]));
    let encoded = value.to_string();
    assert!(!encoded.contains("projection_unsupported"), "{value}");
    assert!(!encoded.contains("byteLength"), "{value}");
    value
}

fn mutate(
    replica: &Path,
    artifact: &str,
    kind: &str,
    actor: &str,
    operation: &str,
    mutation: Value,
) -> Value {
    let output = execute(&args([
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        &mutation.to_string(),
        "--actor-id",
        actor,
        "--operation-id",
        operation,
        "--artifact-id",
        artifact,
        "--kind",
        kind,
        "--mode",
        "edit",
        "--json",
    ]));
    let value = parse_output(&output);
    assert!(
        output.status.success(),
        "mutate failed op={operation} mutation={mutation} response={value}"
    );
    value
}

fn mutate_conflict(
    replica: &Path,
    artifact: &str,
    kind: &str,
    actor: &str,
    operation: &str,
    mutation: Value,
) -> Value {
    let value = run_failure(&args([
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        &mutation.to_string(),
        "--actor-id",
        actor,
        "--operation-id",
        operation,
        "--artifact-id",
        artifact,
        "--kind",
        kind,
        "--mode",
        "edit",
        "--json",
    ]));
    assert_eq!(
        value["error"]["code"],
        "office.collaboration.mutation_match_conflict"
    );
    value
}

fn content(read: &Value) -> &Value {
    &read["data"]["content"]
}

#[test]
fn cli_locate_then_patch_preserves_unrelated_regions() {
    let temp = tempfile::tempdir().unwrap();
    cli_markdown(temp.path());
    cli_document(temp.path());
    cli_spreadsheet(temp.path());
    cli_presentation(temp.path());
    cli_pdf(temp.path());
}

#[tokio::test]
async fn mcp_locate_then_patch_preserves_unrelated_regions() {
    use std::time::Duration;
    use tokio::io::{AsyncWriteExt, BufReader};
    use tokio::process::Command;

    const TIMEOUT: Duration = Duration::from_secs(20);
    let temp = tempfile::tempdir().unwrap();
    let mut child = Command::new(binary())
        .args(["mcp"])
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .stderr(std::process::Stdio::piped())
        .kill_on_drop(true)
        .spawn()
        .unwrap();
    let mut stdin = child.stdin.take().unwrap();
    let mut stdout = BufReader::new(child.stdout.take().unwrap());
    mcp_request(
        &mut stdin,
        &mut stdout,
        json!({
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2025-06-18",
                "capabilities": {},
                "clientInfo": { "name": "office-locate-test", "version": "1" }
            }
        }),
        TIMEOUT,
    )
    .await;
    stdin
        .write_all(
            b"{\"jsonrpc\":\"2.0\",\"method\":\"notifications/initialized\",\"params\":{}}\n",
        )
        .await
        .unwrap();
    stdin.flush().await.unwrap();

    let mut id = 2_u32;
    mcp_markdown(&mut stdin, &mut stdout, &mut id, temp.path(), TIMEOUT).await;
    mcp_document(&mut stdin, &mut stdout, &mut id, temp.path(), TIMEOUT).await;
    mcp_spreadsheet(&mut stdin, &mut stdout, &mut id, temp.path(), TIMEOUT).await;
    mcp_presentation(&mut stdin, &mut stdout, &mut id, temp.path(), TIMEOUT).await;
    mcp_pdf(&mut stdin, &mut stdout, &mut id, temp.path(), TIMEOUT).await;
    drop(stdin);
    let status = tokio::time::timeout(TIMEOUT, child.wait())
        .await
        .unwrap()
        .unwrap();
    assert!(status.success());
}

fn cli_markdown(root: &Path) {
    let replica = root.join("markdown-patch.replica");
    join(
        &replica,
        "fixture-markdown",
        "markdown",
        "agent-markdown",
        "910001",
        &decode(YJS_MARKDOWN_UPDATE_BASE64),
    );
    let read = read_replica(&replica);
    assert_eq!(content(&read)["kind"], "markdown");
    let alpha = slice(&read, "# Shared");
    let beta = slice(&read, "Yjs to Yrs.");
    mutate(
        &replica,
        "fixture-markdown",
        "markdown",
        "agent-markdown",
        "markdown-outside",
        splice(&beta, "Yjs to Yrs!"),
    );
    let moved = read_replica(&replica);
    assert_eq!(slice(&moved, "Yjs to Yrs!")["text"], "Yjs to Yrs!");
    mutate(
        &replica,
        "fixture-markdown",
        "markdown",
        "agent-markdown",
        "markdown-patch",
        splice(&alpha, "# Shared live"),
    );
    assert_eq!(
        content(&read_replica(&replica))["source"],
        "# Shared live\n\nYjs to Yrs!"
    );

    let stale = root.join("markdown-stale.replica");
    join(
        &stale,
        "fixture-markdown",
        "markdown",
        "agent-markdown",
        "910002",
        &decode(YJS_MARKDOWN_UPDATE_BASE64),
    );
    let stale_read = read_replica(&stale);
    let stale_alpha = slice(&stale_read, "# Shared");
    mutate(
        &stale,
        "fixture-markdown",
        "markdown",
        "agent-markdown",
        "markdown-stale-outside",
        splice(&stale_alpha, "Title"),
    );
    mutate_conflict(
        &stale,
        "fixture-markdown",
        "markdown",
        "agent-markdown",
        "markdown-stale-patch",
        splice(&stale_alpha, "# Shared live"),
    );
    assert_eq!(
        content(&read_replica(&stale))["source"],
        "Title\n\nYjs to Yrs."
    );

    let whole = root.join("markdown-whole.replica");
    join(
        &whole,
        "fixture-markdown",
        "markdown",
        "agent-markdown",
        "910003",
        &decode(YJS_MARKDOWN_UPDATE_BASE64),
    );
    let whole_read = read_replica(&whole);
    let whole_beta = slice(&whole_read, "Yjs to Yrs.");
    mutate(
        &whole,
        "fixture-markdown",
        "markdown",
        "agent-markdown",
        "markdown-whole-outside",
        splice(&whole_beta, "Yjs to Yrs!"),
    );
    let source_after_splice = content(&read_replica(&whole))["source"]
        .as_str()
        .unwrap()
        .to_owned();
    mutate(
        &whole,
        "fixture-markdown",
        "markdown",
        "agent-markdown",
        "markdown-whole-replace",
        json!({
            "type": "markdown-replace",
            "expectedMarkdown": source_after_splice,
            "markdown": source_after_splice.replacen("# Shared", "# Shared live", 1),
        }),
    );
    assert_eq!(
        content(&read_replica(&whole))["source"],
        "# Shared live\n\nYjs to Yrs!"
    );
}

fn cli_document(root: &Path) {
    let replica = root.join("document-patch.replica");
    join(
        &replica,
        "fixture-document",
        "document",
        "agent-document",
        "910011",
        &decode(YJS_DOCUMENT_UPDATE_BASE64),
    );
    let seeded = seed_document(&replica, "document-seed");
    let alpha = paragraph(&seeded, "Hello 😀 world");
    let beta = paragraph(&seeded, "BETA");
    assert_eq!(alpha["startUtf16"], 0);
    assert!(alpha["endUtf16"].as_u64().unwrap() > 0);
    mutate(
        &replica,
        "fixture-document",
        "document",
        "agent-document",
        "document-outside",
        replace_paragraph(&beta, "BETA2"),
    );
    mutate(
        &replica,
        "fixture-document",
        "document",
        "agent-document",
        "document-patch",
        replace_paragraph(&alpha, "Hello agent"),
    );
    let patched = read_replica(&replica);
    assert_eq!(paragraph(&patched, "Hello agent")["text"], "Hello agent");
    assert_eq!(paragraph(&patched, "BETA2")["text"], "BETA2");

    let stale = root.join("document-stale.replica");
    join(
        &stale,
        "fixture-document",
        "document",
        "agent-document",
        "910012",
        &decode(YJS_DOCUMENT_UPDATE_BASE64),
    );
    let stale_read = seed_document(&stale, "document-stale-seed");
    let stale_alpha = paragraph(&stale_read, "Hello 😀 world");
    mutate(
        &stale,
        "fixture-document",
        "document",
        "agent-document",
        "document-stale-outside",
        replace_paragraph(&stale_alpha, "Hello human"),
    );
    mutate_conflict(
        &stale,
        "fixture-document",
        "document",
        "agent-document",
        "document-stale-patch",
        replace_paragraph(&stale_alpha, "Hello agent"),
    );
    assert_eq!(paragraph(&read_replica(&stale), "BETA")["text"], "BETA");

    let whole = root.join("document-whole.replica");
    join(
        &whole,
        "fixture-document",
        "document",
        "agent-document",
        "910013",
        &decode(YJS_DOCUMENT_UPDATE_BASE64),
    );
    let whole_read = seed_document(&whole, "document-whole-seed");
    let whole_beta = paragraph(&whole_read, "BETA");
    let plain = content(&whole_read)["plainText"]
        .as_str()
        .unwrap()
        .to_owned();
    mutate(
        &whole,
        "fixture-document",
        "document",
        "agent-document",
        "document-whole-outside",
        replace_paragraph(&whole_beta, "BETA2"),
    );
    mutate_conflict(
        &whole,
        "fixture-document",
        "document",
        "agent-document",
        "document-whole-replace",
        json!({
            "type": "document-replace-text",
            "search": plain,
            "replacement": plain.replacen("Hello 😀 world", "Hello agent", 1),
            "expectedMatches": 1,
        }),
    );
    assert_eq!(paragraph(&read_replica(&whole), "BETA2")["text"], "BETA2");
}

fn cli_spreadsheet(root: &Path) {
    let replica = root.join("spreadsheet-patch.replica");
    join(
        &replica,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        "910021",
        &spreadsheet_collaboration_fixture(),
    );
    mutate(
        &replica,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        "spreadsheet-seed",
        json!({
            "type": "spreadsheet-set-cell",
            "sheetId": "sheet-data",
            "row": 2,
            "column": 0,
            "expectedCell": null,
            "nextCell": { "v": 2, "m": "2" }
        }),
    );
    let read = read_replica(&replica);
    assert_eq!(content(&read)["kind"], "spreadsheet");
    let alpha = cell(&read, 1, 0);
    let beta = cell(&read, 2, 0);
    mutate(
        &replica,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        "spreadsheet-outside",
        set_cell(2, 0, &beta, json!({ "v": 3, "m": "3" })),
    );
    let mut next = alpha.clone();
    next["v"] = json!(11);
    next["m"] = json!("11");
    mutate(
        &replica,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        "spreadsheet-patch",
        set_cell(1, 0, &alpha, next),
    );
    let patched = read_replica(&replica);
    assert_eq!(cell(&patched, 1, 0)["v"].as_f64(), Some(11.0));
    assert_eq!(cell(&patched, 2, 0)["v"].as_f64(), Some(3.0));

    let stale = root.join("spreadsheet-stale.replica");
    join(
        &stale,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        "910022",
        &spreadsheet_collaboration_fixture(),
    );
    mutate(
        &stale,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        "spreadsheet-stale-seed",
        json!({
            "type": "spreadsheet-set-cell",
            "sheetId": "sheet-data",
            "row": 2,
            "column": 0,
            "expectedCell": null,
            "nextCell": { "v": 2, "m": "2" }
        }),
    );
    let stale_read = read_replica(&stale);
    let stale_alpha = cell(&stale_read, 1, 0);
    let mut conflicting = stale_alpha.clone();
    conflicting["v"] = json!(99);
    conflicting["m"] = json!("99");
    mutate(
        &stale,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        "spreadsheet-stale-outside",
        set_cell(1, 0, &stale_alpha, conflicting),
    );
    let mut agent = stale_alpha.clone();
    agent["v"] = json!(11);
    agent["m"] = json!("11");
    mutate_conflict(
        &stale,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        "spreadsheet-stale-patch",
        set_cell(1, 0, &stale_alpha, agent),
    );
    assert_eq!(cell(&read_replica(&stale), 2, 0)["v"].as_f64(), Some(2.0));

    let whole = root.join("spreadsheet-whole.replica");
    join(
        &whole,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        "910023",
        &spreadsheet_collaboration_fixture(),
    );
    mutate(
        &whole,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        "spreadsheet-whole-seed",
        json!({
            "type": "spreadsheet-set-cell",
            "sheetId": "sheet-data",
            "row": 2,
            "column": 0,
            "expectedCell": null,
            "nextCell": { "v": 2, "m": "2" }
        }),
    );
    let whole_read = read_replica(&whole);
    let whole_alpha = cell(&whole_read, 1, 0);
    let whole_beta = cell(&whole_read, 2, 0);
    mutate(
        &whole,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        "spreadsheet-whole-outside",
        set_cell(2, 0, &whole_beta, json!({ "v": 3, "m": "3" })),
    );
    let mut whole_next = whole_alpha.clone();
    whole_next["v"] = json!(11);
    whole_next["m"] = json!("11");
    mutate(
        &whole,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        "spreadsheet-whole-replace",
        json!({
            "type": "spreadsheet-batch-cells",
            "sheetId": "sheet-data",
            "changes": [
                { "row": 1, "column": 0, "expectedCell": whole_alpha, "nextCell": whole_next },
                { "row": 2, "column": 0, "expectedCell": whole_beta, "nextCell": whole_beta.clone() }
            ]
        }),
    );
    let replaced = read_replica(&whole);
    assert_eq!(cell(&replaced, 1, 0)["v"].as_f64(), Some(11.0));
    assert_eq!(cell(&replaced, 2, 0)["v"].as_f64(), Some(3.0));
}

fn cli_presentation(root: &Path) {
    let replica = root.join("presentation-patch.replica");
    join(
        &replica,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        "910031",
        &presentation_collaboration_fixture(),
    );
    mutate(
        &replica,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        "presentation-seed",
        json!({
            "type": "presentation-create-element",
            "containerKind": "slide",
            "containerId": "slide-1",
            "afterElementId": "element-title",
            "element": presentation_scene_element("element-beta", "BETA", "shape")
        }),
    );
    let read = read_replica(&replica);
    assert_eq!(content(&read)["kind"], "presentation");
    let alpha = element(&read, "element-title");
    let beta = element(&read, "element-beta");
    let mut beta_next = beta.clone();
    beta_next["text"] = json!("BETA2");
    mutate(
        &replica,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        "presentation-outside",
        update_element("element-beta", &beta, beta_next),
    );
    let mut alpha_next = alpha.clone();
    alpha_next["text"] = json!("ALPHA2");
    mutate(
        &replica,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        "presentation-patch",
        update_element("element-title", &alpha, alpha_next),
    );
    let patched = read_replica(&replica);
    assert_eq!(element(&patched, "element-title")["text"], "ALPHA2");
    assert_eq!(element(&patched, "element-beta")["text"], "BETA2");

    let stale = root.join("presentation-stale.replica");
    join(
        &stale,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        "910032",
        &presentation_collaboration_fixture(),
    );
    mutate(
        &stale,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        "presentation-stale-seed",
        json!({
            "type": "presentation-create-element",
            "containerKind": "slide",
            "containerId": "slide-1",
            "afterElementId": "element-title",
            "element": presentation_scene_element("element-beta", "BETA", "shape")
        }),
    );
    let stale_read = read_replica(&stale);
    let stale_alpha = element(&stale_read, "element-title");
    let mut conflicting = stale_alpha.clone();
    conflicting["text"] = json!("HUMAN");
    mutate(
        &stale,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        "presentation-stale-outside",
        update_element("element-title", &stale_alpha, conflicting),
    );
    let mut agent = stale_alpha.clone();
    agent["text"] = json!("ALPHA2");
    mutate_conflict(
        &stale,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        "presentation-stale-patch",
        update_element("element-title", &stale_alpha, agent),
    );
    assert_eq!(
        element(&read_replica(&stale), "element-beta")["text"],
        "BETA"
    );

    let whole = root.join("presentation-whole.replica");
    join(
        &whole,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        "910033",
        &presentation_collaboration_fixture(),
    );
    mutate(
        &whole,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        "presentation-whole-seed",
        json!({
            "type": "presentation-create-element",
            "containerKind": "slide",
            "containerId": "slide-1",
            "afterElementId": "element-title",
            "element": presentation_scene_element("element-beta", "BETA", "shape")
        }),
    );
    let whole_read = read_replica(&whole);
    let whole_alpha = element(&whole_read, "element-title");
    let whole_beta = element(&whole_read, "element-beta");
    let mut live_beta = whole_beta.clone();
    live_beta["text"] = json!("BETA2");
    mutate(
        &whole,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        "presentation-whole-outside",
        update_element("element-beta", &whole_beta, live_beta),
    );
    let mut whole_alpha_next = whole_alpha.clone();
    whole_alpha_next["text"] = json!("ALPHA2");
    mutate(
        &whole,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        "presentation-whole-alpha",
        update_element("element-title", &whole_alpha, whole_alpha_next),
    );
    mutate(
        &whole,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        "presentation-whole-echo",
        update_element("element-beta", &whole_beta, whole_beta.clone()),
    );
    let replaced = read_replica(&whole);
    assert_eq!(element(&replaced, "element-title")["text"], "ALPHA2");
    assert_eq!(element(&replaced, "element-beta")["text"], "BETA2");
}

fn cli_pdf(root: &Path) {
    let replica = root.join("pdf-patch.replica");
    join(
        &replica,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        "910041",
        &pdf_collaboration_fixture(),
    );
    mutate(
        &replica,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        "pdf-seed",
        json!({
            "type": "pdf-set-form-value",
            "fieldId": "Applicant.Email",
            "value": "beta@example.test",
            "expectedValue": ""
        }),
    );
    let read = read_replica(&replica);
    assert_eq!(content(&read)["kind"], "pdf");
    assert_eq!(content(&read)["pageCount"], 3);
    assert!(!read
        .to_string()
        .contains("0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef"));
    let name = field(&read, "Applicant.Name");
    let email = field(&read, "Applicant.Email");
    mutate(
        &replica,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        "pdf-outside",
        set_field("Applicant.Email", "beta2@example.test", &email),
    );
    mutate(
        &replica,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        "pdf-patch",
        set_field("Applicant.Name", "Grace", &name),
    );
    let patched = read_replica(&replica);
    assert_eq!(field(&patched, "Applicant.Name"), "Grace");
    assert_eq!(field(&patched, "Applicant.Email"), "beta2@example.test");

    let stale = root.join("pdf-stale.replica");
    join(
        &stale,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        "910042",
        &pdf_collaboration_fixture(),
    );
    mutate(
        &stale,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        "pdf-stale-seed",
        json!({
            "type": "pdf-set-form-value",
            "fieldId": "Applicant.Email",
            "value": "beta@example.test",
            "expectedValue": ""
        }),
    );
    let stale_read = read_replica(&stale);
    let stale_name = field(&stale_read, "Applicant.Name");
    mutate(
        &stale,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        "pdf-stale-outside",
        set_field("Applicant.Name", "Other", &stale_name),
    );
    mutate_conflict(
        &stale,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        "pdf-stale-patch",
        set_field("Applicant.Name", "Grace", &stale_name),
    );
    assert_eq!(
        field(&read_replica(&stale), "Applicant.Email"),
        "beta@example.test"
    );

    let whole = root.join("pdf-whole.replica");
    join(
        &whole,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        "910043",
        &pdf_collaboration_fixture(),
    );
    mutate(
        &whole,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        "pdf-whole-seed",
        json!({
            "type": "pdf-set-form-value",
            "fieldId": "Applicant.Email",
            "value": "beta@example.test",
            "expectedValue": ""
        }),
    );
    let whole_read = read_replica(&whole);
    let whole_email = field(&whole_read, "Applicant.Email");
    let whole_name = field(&whole_read, "Applicant.Name");
    mutate(
        &whole,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        "pdf-whole-outside",
        set_field("Applicant.Email", "beta2@example.test", &whole_email),
    );
    mutate(
        &whole,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        "pdf-whole-name",
        set_field("Applicant.Name", "Grace", &whole_name),
    );
    mutate_conflict(
        &whole,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        "pdf-whole-echo",
        set_field("Applicant.Email", &whole_email, &whole_email),
    );
    let replaced = read_replica(&whole);
    assert_eq!(field(&replaced, "Applicant.Name"), "Grace");
    assert_eq!(field(&replaced, "Applicant.Email"), "beta2@example.test");
}

fn seed_document(replica: &Path, operation: &str) -> Value {
    let read = read_replica(replica);
    let anchor = paragraph(&read, "Hello 😀 world")["paragraphId"]
        .as_str()
        .unwrap()
        .to_owned();
    mutate(
        replica,
        "fixture-document",
        "document",
        "agent-document",
        operation,
        json!({
            "type": "document-insert-paragraph",
            "anchorParagraphId": anchor,
            "position": "after",
            "paragraphId": "00000020",
            "textId": "00000021",
            "text": "BETA"
        }),
    );
    read_replica(replica)
}

fn slice(read: &Value, text: &str) -> Value {
    content(read)["slices"]
        .as_array()
        .unwrap()
        .iter()
        .find(|slice| slice["text"] == text)
        .unwrap_or_else(|| panic!("missing slice {text} in {read}"))
        .clone()
}

fn splice(slice: &Value, insert: &str) -> Value {
    let start = slice["startUtf16"].as_u64().unwrap();
    let end = slice["endUtf16"].as_u64().unwrap();
    json!({
        "type": "markdown-splice",
        "indexUtf16": start,
        "deleteUtf16": end - start,
        "expectedSlice": slice["text"],
        "insert": insert,
    })
}

fn paragraph(read: &Value, text: &str) -> Value {
    content(read)["paragraphs"]
        .as_array()
        .unwrap()
        .iter()
        .find(|paragraph| paragraph["text"] == text)
        .unwrap_or_else(|| panic!("missing paragraph {text} in {read}"))
        .clone()
}

fn replace_paragraph(paragraph: &Value, replacement: &str) -> Value {
    json!({
        "type": "document-replace-paragraph",
        "paragraphId": paragraph["paragraphId"],
        "expectedTextId": paragraph["textId"],
        "expectedText": paragraph["text"],
        "replacement": replacement,
    })
}

fn cell(read: &Value, row: u64, column: u64) -> Value {
    content(read)["sheets"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|sheet| sheet["cells"].as_array().unwrap().iter())
        .find(|cell| cell["row"] == row && cell["column"] == column)
        .unwrap_or_else(|| panic!("missing cell {row}:{column} in {read}"))["cell"]
        .clone()
}

fn set_cell(row: u64, column: u64, expected: &Value, next: Value) -> Value {
    json!({
        "type": "spreadsheet-set-cell",
        "sheetId": "sheet-data",
        "row": row,
        "column": column,
        "expectedCell": expected,
        "nextCell": next,
    })
}

fn element(read: &Value, element_id: &str) -> Value {
    content(read)["containers"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|container| container["elements"].as_array().unwrap().iter())
        .find(|element| element["elementId"] == element_id)
        .unwrap_or_else(|| panic!("missing element {element_id} in {read}"))["element"]
        .clone()
}

fn update_element(element_id: &str, expected: &Value, next: Value) -> Value {
    json!({
        "type": "presentation-update-element",
        "containerKind": "slide",
        "containerId": "slide-1",
        "elementId": element_id,
        "expectedElement": expected,
        "nextElement": next,
    })
}

fn field(read: &Value, field_id: &str) -> String {
    content(read)["formFields"]
        .as_array()
        .unwrap()
        .iter()
        .find(|field| field["fieldId"] == field_id)
        .unwrap_or_else(|| panic!("missing field {field_id} in {read}"))["value"]
        .as_str()
        .unwrap()
        .to_owned()
}

fn set_field(field_id: &str, value: &str, expected: &str) -> Value {
    json!({
        "type": "pdf-set-form-value",
        "fieldId": field_id,
        "value": value,
        "expectedValue": expected,
    })
}

fn decode(value: &str) -> Vec<u8> {
    base64::engine::general_purpose::STANDARD
        .decode(value)
        .unwrap()
}

fn args<const N: usize>(values: [&str; N]) -> Vec<String> {
    values.into_iter().map(str::to_owned).collect()
}

async fn mcp_request(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    value: Value,
    timeout: std::time::Duration,
) -> Value {
    use tokio::io::{AsyncBufReadExt, AsyncWriteExt};
    let mut encoded = serde_json::to_vec(&value).unwrap();
    encoded.push(b'\n');
    stdin.write_all(&encoded).await.unwrap();
    stdin.flush().await.unwrap();
    let mut line = String::new();
    let bytes = tokio::time::timeout(timeout, stdout.read_line(&mut line))
        .await
        .unwrap()
        .unwrap();
    assert!(bytes > 0, "native Office MCP closed before responding");
    serde_json::from_str(&line).unwrap()
}

async fn mcp_call(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    id: u32,
    name: &str,
    arguments: Value,
    timeout: std::time::Duration,
) -> Value {
    mcp_request(
        stdin,
        stdout,
        json!({
            "jsonrpc": "2.0",
            "id": id,
            "method": "tools/call",
            "params": { "name": name, "arguments": arguments }
        }),
        timeout,
    )
    .await
}

fn mcp_ok(response: &Value) -> &Value {
    assert_ne!(response["result"]["isError"], true, "{response}");
    &response["result"]["structuredContent"]
}

fn mcp_wrap(projection: &Value) -> Value {
    json!({ "data": projection })
}

async fn mcp_create(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    id: &mut u32,
    store: &Path,
    artifact: &str,
    kind: &str,
    actor: &str,
    client_id: u64,
    update_base64: &str,
    timeout: std::time::Duration,
) {
    let response = mcp_call(
        stdin,
        stdout,
        *id,
        "office_collaboration_create",
        json!({
            "store": store.to_str().unwrap(),
            "artifactId": artifact,
            "kind": kind,
            "actorId": actor,
            "actorKind": "agent",
            "mode": "edit",
            "operationId": format!("join-{kind}-{client_id}"),
            "clientId": client_id,
            "initialUpdateBase64": update_base64
        }),
        timeout,
    )
    .await;
    *id += 1;
    mcp_ok(&response);
}

async fn mcp_read(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    id: &mut u32,
    store: &Path,
    timeout: std::time::Duration,
) -> Value {
    let response = mcp_call(
        stdin,
        stdout,
        *id,
        "office_collaboration_read",
        json!({ "store": store.to_str().unwrap() }),
        timeout,
    )
    .await;
    *id += 1;
    let projection = mcp_ok(&response).clone();
    let encoded = projection.to_string();
    assert!(!encoded.contains("projection_unsupported"), "{projection}");
    assert!(!encoded.contains("byteLength"), "{projection}");
    mcp_wrap(&projection)
}

async fn mcp_mutate(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    id: &mut u32,
    store: &Path,
    artifact: &str,
    kind: &str,
    actor: &str,
    operation: &str,
    mutation: Value,
    timeout: std::time::Duration,
) -> Value {
    let response = mcp_call(
        stdin,
        stdout,
        *id,
        "office_collaboration_mutate",
        json!({
            "store": store.to_str().unwrap(),
            "operationId": operation,
            "actorId": actor,
            "mode": "edit",
            "artifactId": artifact,
            "kind": kind,
            "mutation": mutation
        }),
        timeout,
    )
    .await;
    *id += 1;
    response
}

fn mcp_conflict(response: &Value) {
    assert_eq!(response["result"]["isError"], true, "{response}");
    assert_eq!(
        response["result"]["structuredContent"]["code"],
        "office.collaboration.mutation_match_conflict"
    );
}

async fn mcp_markdown(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    id: &mut u32,
    root: &Path,
    timeout: std::time::Duration,
) {
    let replica = root.join("mcp-markdown.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &replica,
        "fixture-markdown",
        "markdown",
        "agent-markdown",
        920001,
        YJS_MARKDOWN_UPDATE_BASE64,
        timeout,
    )
    .await;
    let read = mcp_read(stdin, stdout, id, &replica, timeout).await;
    let alpha = slice(&read, "# Shared");
    let beta = slice(&read, "Yjs to Yrs.");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &replica,
            "fixture-markdown",
            "markdown",
            "agent-markdown",
            "mcp-markdown-outside",
            splice(&beta, "Yjs to Yrs!"),
            timeout,
        )
        .await,
    );
    let moved = mcp_read(stdin, stdout, id, &replica, timeout).await;
    assert_eq!(slice(&moved, "Yjs to Yrs!")["text"], "Yjs to Yrs!");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &replica,
            "fixture-markdown",
            "markdown",
            "agent-markdown",
            "mcp-markdown-patch",
            splice(&alpha, "# Shared live"),
            timeout,
        )
        .await,
    );
    assert_eq!(
        content(&mcp_read(stdin, stdout, id, &replica, timeout).await)["source"],
        "# Shared live\n\nYjs to Yrs!"
    );

    let stale = root.join("mcp-markdown-stale.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &stale,
        "fixture-markdown",
        "markdown",
        "agent-markdown",
        920002,
        YJS_MARKDOWN_UPDATE_BASE64,
        timeout,
    )
    .await;
    let stale_read = mcp_read(stdin, stdout, id, &stale, timeout).await;
    let stale_alpha = slice(&stale_read, "# Shared");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &stale,
            "fixture-markdown",
            "markdown",
            "agent-markdown",
            "mcp-markdown-stale-outside",
            splice(&stale_alpha, "Title"),
            timeout,
        )
        .await,
    );
    mcp_conflict(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &stale,
            "fixture-markdown",
            "markdown",
            "agent-markdown",
            "mcp-markdown-stale-patch",
            splice(&stale_alpha, "# Shared live"),
            timeout,
        )
        .await,
    );
    assert_eq!(
        content(&mcp_read(stdin, stdout, id, &stale, timeout).await)["source"],
        "Title\n\nYjs to Yrs."
    );

    let whole = root.join("mcp-markdown-whole.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &whole,
        "fixture-markdown",
        "markdown",
        "agent-markdown",
        920003,
        YJS_MARKDOWN_UPDATE_BASE64,
        timeout,
    )
    .await;
    let whole_read = mcp_read(stdin, stdout, id, &whole, timeout).await;
    let whole_beta = slice(&whole_read, "Yjs to Yrs.");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &whole,
            "fixture-markdown",
            "markdown",
            "agent-markdown",
            "mcp-markdown-whole-outside",
            splice(&whole_beta, "Yjs to Yrs!"),
            timeout,
        )
        .await,
    );
    let source_after_splice = content(&mcp_read(stdin, stdout, id, &whole, timeout).await)
        ["source"]
        .as_str()
        .unwrap()
        .to_owned();
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &whole,
            "fixture-markdown",
            "markdown",
            "agent-markdown",
            "mcp-markdown-whole-replace",
            json!({
                "type": "markdown-replace",
                "expectedMarkdown": source_after_splice,
                "markdown": source_after_splice.replacen("# Shared", "# Shared live", 1),
            }),
            timeout,
        )
        .await,
    );
    assert_eq!(
        content(&mcp_read(stdin, stdout, id, &whole, timeout).await)["source"],
        "# Shared live\n\nYjs to Yrs!"
    );
}

async fn mcp_document(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    id: &mut u32,
    root: &Path,
    timeout: std::time::Duration,
) {
    let replica = root.join("mcp-document.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &replica,
        "fixture-document",
        "document",
        "agent-document",
        920011,
        YJS_DOCUMENT_UPDATE_BASE64,
        timeout,
    )
    .await;
    let seeded = mcp_seed_document(stdin, stdout, id, &replica, "mcp-document-seed", timeout).await;
    let alpha = paragraph(&seeded, "Hello 😀 world");
    let beta = paragraph(&seeded, "BETA");
    assert_eq!(alpha["startUtf16"], 0);
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &replica,
            "fixture-document",
            "document",
            "agent-document",
            "mcp-document-outside",
            replace_paragraph(&beta, "BETA2"),
            timeout,
        )
        .await,
    );
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &replica,
            "fixture-document",
            "document",
            "agent-document",
            "mcp-document-patch",
            replace_paragraph(&alpha, "Hello agent"),
            timeout,
        )
        .await,
    );
    let patched = mcp_read(stdin, stdout, id, &replica, timeout).await;
    assert_eq!(paragraph(&patched, "BETA2")["text"], "BETA2");

    let stale = root.join("mcp-document-stale.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &stale,
        "fixture-document",
        "document",
        "agent-document",
        920012,
        YJS_DOCUMENT_UPDATE_BASE64,
        timeout,
    )
    .await;
    let stale_read = mcp_seed_document(
        stdin,
        stdout,
        id,
        &stale,
        "mcp-document-stale-seed",
        timeout,
    )
    .await;
    let stale_alpha = paragraph(&stale_read, "Hello 😀 world");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &stale,
            "fixture-document",
            "document",
            "agent-document",
            "mcp-document-stale-outside",
            replace_paragraph(&stale_alpha, "Hello human"),
            timeout,
        )
        .await,
    );
    mcp_conflict(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &stale,
            "fixture-document",
            "document",
            "agent-document",
            "mcp-document-stale-patch",
            replace_paragraph(&stale_alpha, "Hello agent"),
            timeout,
        )
        .await,
    );
    assert_eq!(
        paragraph(&mcp_read(stdin, stdout, id, &stale, timeout).await, "BETA")["text"],
        "BETA"
    );

    let whole = root.join("mcp-document-whole.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &whole,
        "fixture-document",
        "document",
        "agent-document",
        920013,
        YJS_DOCUMENT_UPDATE_BASE64,
        timeout,
    )
    .await;
    let whole_read = mcp_seed_document(
        stdin,
        stdout,
        id,
        &whole,
        "mcp-document-whole-seed",
        timeout,
    )
    .await;
    let plain = content(&whole_read)["plainText"]
        .as_str()
        .unwrap()
        .to_owned();
    let whole_beta = paragraph(&whole_read, "BETA");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &whole,
            "fixture-document",
            "document",
            "agent-document",
            "mcp-document-whole-outside",
            replace_paragraph(&whole_beta, "BETA2"),
            timeout,
        )
        .await,
    );
    mcp_conflict(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &whole,
            "fixture-document",
            "document",
            "agent-document",
            "mcp-document-whole-replace",
            json!({
                "type": "document-replace-text",
                "search": plain,
                "replacement": plain.replacen("Hello 😀 world", "Hello agent", 1),
                "expectedMatches": 1
            }),
            timeout,
        )
        .await,
    );
    assert_eq!(
        paragraph(&mcp_read(stdin, stdout, id, &whole, timeout).await, "BETA2")["text"],
        "BETA2"
    );
}

async fn mcp_seed_document(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    id: &mut u32,
    replica: &Path,
    operation: &str,
    timeout: std::time::Duration,
) -> Value {
    let read = mcp_read(stdin, stdout, id, replica, timeout).await;
    let anchor = paragraph(&read, "Hello 😀 world")["paragraphId"]
        .as_str()
        .unwrap()
        .to_owned();
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            replica,
            "fixture-document",
            "document",
            "agent-document",
            operation,
            json!({
                "type": "document-insert-paragraph",
                "anchorParagraphId": anchor,
                "position": "after",
                "paragraphId": "00000020",
                "textId": "00000021",
                "text": "BETA"
            }),
            timeout,
        )
        .await,
    );
    mcp_read(stdin, stdout, id, replica, timeout).await
}

async fn mcp_spreadsheet(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    id: &mut u32,
    root: &Path,
    timeout: std::time::Duration,
) {
    let update =
        base64::engine::general_purpose::STANDARD.encode(spreadsheet_collaboration_fixture());
    let replica = root.join("mcp-spreadsheet.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &replica,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        920021,
        &update,
        timeout,
    )
    .await;
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &replica,
            "fixture-spreadsheet",
            "spreadsheet",
            "agent-spreadsheet",
            "mcp-spreadsheet-seed",
            json!({
                "type": "spreadsheet-set-cell",
                "sheetId": "sheet-data",
                "row": 2,
                "column": 0,
                "expectedCell": null,
                "nextCell": { "v": 2, "m": "2" }
            }),
            timeout,
        )
        .await,
    );
    let read = mcp_read(stdin, stdout, id, &replica, timeout).await;
    let alpha = cell(&read, 1, 0);
    let beta = cell(&read, 2, 0);
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &replica,
            "fixture-spreadsheet",
            "spreadsheet",
            "agent-spreadsheet",
            "mcp-spreadsheet-outside",
            set_cell(2, 0, &beta, json!({ "v": 3, "m": "3" })),
            timeout,
        )
        .await,
    );
    let mut next = alpha.clone();
    next["v"] = json!(11);
    next["m"] = json!("11");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &replica,
            "fixture-spreadsheet",
            "spreadsheet",
            "agent-spreadsheet",
            "mcp-spreadsheet-patch",
            set_cell(1, 0, &alpha, next),
            timeout,
        )
        .await,
    );
    let patched = mcp_read(stdin, stdout, id, &replica, timeout).await;
    assert_eq!(cell(&patched, 2, 0)["v"].as_f64(), Some(3.0));

    let stale = root.join("mcp-spreadsheet-stale.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &stale,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        920022,
        &update,
        timeout,
    )
    .await;
    mcp_ok(&mcp_mutate(stdin, stdout, id, &stale, "fixture-spreadsheet", "spreadsheet", "agent-spreadsheet", "mcp-spreadsheet-stale-seed", json!({
        "type": "spreadsheet-set-cell", "sheetId": "sheet-data", "row": 2, "column": 0, "expectedCell": null, "nextCell": { "v": 2, "m": "2" }
    }), timeout).await);
    let stale_read = mcp_read(stdin, stdout, id, &stale, timeout).await;
    let stale_alpha = cell(&stale_read, 1, 0);
    let mut conflicting = stale_alpha.clone();
    conflicting["v"] = json!(99);
    conflicting["m"] = json!("99");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &stale,
            "fixture-spreadsheet",
            "spreadsheet",
            "agent-spreadsheet",
            "mcp-spreadsheet-stale-outside",
            set_cell(1, 0, &stale_alpha, conflicting),
            timeout,
        )
        .await,
    );
    let mut agent = stale_alpha.clone();
    agent["v"] = json!(11);
    agent["m"] = json!("11");
    mcp_conflict(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &stale,
            "fixture-spreadsheet",
            "spreadsheet",
            "agent-spreadsheet",
            "mcp-spreadsheet-stale-patch",
            set_cell(1, 0, &stale_alpha, agent),
            timeout,
        )
        .await,
    );
    assert_eq!(
        cell(&mcp_read(stdin, stdout, id, &stale, timeout).await, 2, 0)["v"].as_f64(),
        Some(2.0)
    );

    let whole = root.join("mcp-spreadsheet-whole.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &whole,
        "fixture-spreadsheet",
        "spreadsheet",
        "agent-spreadsheet",
        920023,
        &update,
        timeout,
    )
    .await;
    mcp_ok(&mcp_mutate(stdin, stdout, id, &whole, "fixture-spreadsheet", "spreadsheet", "agent-spreadsheet", "mcp-spreadsheet-whole-seed", json!({
        "type": "spreadsheet-set-cell", "sheetId": "sheet-data", "row": 2, "column": 0, "expectedCell": null, "nextCell": { "v": 2, "m": "2" }
    }), timeout).await);
    let whole_read = mcp_read(stdin, stdout, id, &whole, timeout).await;
    let whole_alpha = cell(&whole_read, 1, 0);
    let whole_beta = cell(&whole_read, 2, 0);
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &whole,
            "fixture-spreadsheet",
            "spreadsheet",
            "agent-spreadsheet",
            "mcp-spreadsheet-whole-outside",
            set_cell(2, 0, &whole_beta, json!({ "v": 3, "m": "3" })),
            timeout,
        )
        .await,
    );
    let mut whole_next = whole_alpha.clone();
    whole_next["v"] = json!(11);
    whole_next["m"] = json!("11");
    mcp_ok(&mcp_mutate(stdin, stdout, id, &whole, "fixture-spreadsheet", "spreadsheet", "agent-spreadsheet", "mcp-spreadsheet-whole-replace", json!({
        "type": "spreadsheet-batch-cells",
        "sheetId": "sheet-data",
        "changes": [
            { "row": 1, "column": 0, "expectedCell": whole_alpha, "nextCell": whole_next },
            { "row": 2, "column": 0, "expectedCell": whole_beta, "nextCell": whole_beta.clone() }
        ]
    }), timeout).await);
    let replaced = mcp_read(stdin, stdout, id, &whole, timeout).await;
    assert_eq!(cell(&replaced, 1, 0)["v"].as_f64(), Some(11.0));
    assert_eq!(cell(&replaced, 2, 0)["v"].as_f64(), Some(3.0));
}

async fn mcp_presentation(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    id: &mut u32,
    root: &Path,
    timeout: std::time::Duration,
) {
    let update =
        base64::engine::general_purpose::STANDARD.encode(presentation_collaboration_fixture());
    let element = presentation_scene_element("element-beta", "BETA", "shape");
    let replica = root.join("mcp-presentation.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &replica,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        920031,
        &update,
        timeout,
    )
    .await;
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &replica,
            "fixture-presentation",
            "presentation",
            "agent-presentation",
            "mcp-presentation-seed",
            json!({
                "type": "presentation-create-element",
                "containerKind": "slide",
                "containerId": "slide-1",
                "afterElementId": "element-title",
                "element": element
            }),
            timeout,
        )
        .await,
    );
    let read = mcp_read(stdin, stdout, id, &replica, timeout).await;
    let alpha = element_value(&read, "element-title");
    let beta = element_value(&read, "element-beta");
    let mut beta_next = beta.clone();
    beta_next["text"] = json!("BETA2");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &replica,
            "fixture-presentation",
            "presentation",
            "agent-presentation",
            "mcp-presentation-outside",
            update_element("element-beta", &beta, beta_next),
            timeout,
        )
        .await,
    );
    let mut alpha_next = alpha.clone();
    alpha_next["text"] = json!("ALPHA2");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &replica,
            "fixture-presentation",
            "presentation",
            "agent-presentation",
            "mcp-presentation-patch",
            update_element("element-title", &alpha, alpha_next),
            timeout,
        )
        .await,
    );
    assert_eq!(
        element_value(
            &mcp_read(stdin, stdout, id, &replica, timeout).await,
            "element-beta"
        )["text"],
        "BETA2"
    );

    let stale = root.join("mcp-presentation-stale.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &stale,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        920032,
        &update,
        timeout,
    )
    .await;
    mcp_ok(&mcp_mutate(stdin, stdout, id, &stale, "fixture-presentation", "presentation", "agent-presentation", "mcp-presentation-stale-seed", json!({
        "type": "presentation-create-element", "containerKind": "slide", "containerId": "slide-1", "afterElementId": "element-title",
        "element": presentation_scene_element("element-beta", "BETA", "shape")
    }), timeout).await);
    let stale_read = mcp_read(stdin, stdout, id, &stale, timeout).await;
    let stale_alpha = element_value(&stale_read, "element-title");
    let mut conflicting = stale_alpha.clone();
    conflicting["text"] = json!("HUMAN");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &stale,
            "fixture-presentation",
            "presentation",
            "agent-presentation",
            "mcp-presentation-stale-outside",
            update_element("element-title", &stale_alpha, conflicting),
            timeout,
        )
        .await,
    );
    let mut agent = stale_alpha.clone();
    agent["text"] = json!("ALPHA2");
    mcp_conflict(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &stale,
            "fixture-presentation",
            "presentation",
            "agent-presentation",
            "mcp-presentation-stale-patch",
            update_element("element-title", &stale_alpha, agent),
            timeout,
        )
        .await,
    );
    assert_eq!(
        element_value(
            &mcp_read(stdin, stdout, id, &stale, timeout).await,
            "element-beta"
        )["text"],
        "BETA"
    );

    let whole = root.join("mcp-presentation-whole.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &whole,
        "fixture-presentation",
        "presentation",
        "agent-presentation",
        920033,
        &update,
        timeout,
    )
    .await;
    mcp_ok(&mcp_mutate(stdin, stdout, id, &whole, "fixture-presentation", "presentation", "agent-presentation", "mcp-presentation-whole-seed", json!({
        "type": "presentation-create-element", "containerKind": "slide", "containerId": "slide-1", "afterElementId": "element-title",
        "element": presentation_scene_element("element-beta", "BETA", "shape")
    }), timeout).await);
    let whole_read = mcp_read(stdin, stdout, id, &whole, timeout).await;
    let whole_alpha = element_value(&whole_read, "element-title");
    let whole_beta = element_value(&whole_read, "element-beta");
    let mut live_beta = whole_beta.clone();
    live_beta["text"] = json!("BETA2");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &whole,
            "fixture-presentation",
            "presentation",
            "agent-presentation",
            "mcp-presentation-whole-outside",
            update_element("element-beta", &whole_beta, live_beta),
            timeout,
        )
        .await,
    );
    let mut whole_alpha_next = whole_alpha.clone();
    whole_alpha_next["text"] = json!("ALPHA2");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &whole,
            "fixture-presentation",
            "presentation",
            "agent-presentation",
            "mcp-presentation-whole-alpha",
            update_element("element-title", &whole_alpha, whole_alpha_next),
            timeout,
        )
        .await,
    );
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &whole,
            "fixture-presentation",
            "presentation",
            "agent-presentation",
            "mcp-presentation-whole-echo",
            update_element("element-beta", &whole_beta, whole_beta.clone()),
            timeout,
        )
        .await,
    );
    let replaced = mcp_read(stdin, stdout, id, &whole, timeout).await;
    assert_eq!(element_value(&replaced, "element-title")["text"], "ALPHA2");
    assert_eq!(element_value(&replaced, "element-beta")["text"], "BETA2");
}

async fn mcp_pdf(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    id: &mut u32,
    root: &Path,
    timeout: std::time::Duration,
) {
    let update = base64::engine::general_purpose::STANDARD.encode(pdf_collaboration_fixture());
    let replica = root.join("mcp-pdf.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &replica,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        920041,
        &update,
        timeout,
    )
    .await;
    mcp_ok(&mcp_mutate(stdin, stdout, id, &replica, "fixture-pdf", "pdf", "agent-pdf", "mcp-pdf-seed", json!({
        "type": "pdf-set-form-value", "fieldId": "Applicant.Email", "value": "beta@example.test", "expectedValue": ""
    }), timeout).await);
    let read = mcp_read(stdin, stdout, id, &replica, timeout).await;
    assert_eq!(content(&read)["pageCount"], 3);
    let name = field(&read, "Applicant.Name");
    let email = field(&read, "Applicant.Email");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &replica,
            "fixture-pdf",
            "pdf",
            "agent-pdf",
            "mcp-pdf-outside",
            set_field("Applicant.Email", "beta2@example.test", &email),
            timeout,
        )
        .await,
    );
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &replica,
            "fixture-pdf",
            "pdf",
            "agent-pdf",
            "mcp-pdf-patch",
            set_field("Applicant.Name", "Grace", &name),
            timeout,
        )
        .await,
    );
    let patched = mcp_read(stdin, stdout, id, &replica, timeout).await;
    assert_eq!(field(&patched, "Applicant.Email"), "beta2@example.test");

    let stale = root.join("mcp-pdf-stale.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &stale,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        920042,
        &update,
        timeout,
    )
    .await;
    mcp_ok(&mcp_mutate(stdin, stdout, id, &stale, "fixture-pdf", "pdf", "agent-pdf", "mcp-pdf-stale-seed", json!({
        "type": "pdf-set-form-value", "fieldId": "Applicant.Email", "value": "beta@example.test", "expectedValue": ""
    }), timeout).await);
    let stale_read = mcp_read(stdin, stdout, id, &stale, timeout).await;
    let stale_name = field(&stale_read, "Applicant.Name");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &stale,
            "fixture-pdf",
            "pdf",
            "agent-pdf",
            "mcp-pdf-stale-outside",
            set_field("Applicant.Name", "Other", &stale_name),
            timeout,
        )
        .await,
    );
    mcp_conflict(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &stale,
            "fixture-pdf",
            "pdf",
            "agent-pdf",
            "mcp-pdf-stale-patch",
            set_field("Applicant.Name", "Grace", &stale_name),
            timeout,
        )
        .await,
    );
    assert_eq!(
        field(
            &mcp_read(stdin, stdout, id, &stale, timeout).await,
            "Applicant.Email"
        ),
        "beta@example.test"
    );

    let whole = root.join("mcp-pdf-whole.replica");
    mcp_create(
        stdin,
        stdout,
        id,
        &whole,
        "fixture-pdf",
        "pdf",
        "agent-pdf",
        920043,
        &update,
        timeout,
    )
    .await;
    mcp_ok(&mcp_mutate(stdin, stdout, id, &whole, "fixture-pdf", "pdf", "agent-pdf", "mcp-pdf-whole-seed", json!({
        "type": "pdf-set-form-value", "fieldId": "Applicant.Email", "value": "beta@example.test", "expectedValue": ""
    }), timeout).await);
    let whole_read = mcp_read(stdin, stdout, id, &whole, timeout).await;
    let whole_email = field(&whole_read, "Applicant.Email");
    let whole_name = field(&whole_read, "Applicant.Name");
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &whole,
            "fixture-pdf",
            "pdf",
            "agent-pdf",
            "mcp-pdf-whole-outside",
            set_field("Applicant.Email", "beta2@example.test", &whole_email),
            timeout,
        )
        .await,
    );
    mcp_ok(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &whole,
            "fixture-pdf",
            "pdf",
            "agent-pdf",
            "mcp-pdf-whole-name",
            set_field("Applicant.Name", "Grace", &whole_name),
            timeout,
        )
        .await,
    );
    mcp_conflict(
        &mcp_mutate(
            stdin,
            stdout,
            id,
            &whole,
            "fixture-pdf",
            "pdf",
            "agent-pdf",
            "mcp-pdf-whole-echo",
            set_field("Applicant.Email", &whole_email, &whole_email),
            timeout,
        )
        .await,
    );
    let replaced = mcp_read(stdin, stdout, id, &whole, timeout).await;
    assert_eq!(field(&replaced, "Applicant.Name"), "Grace");
    assert_eq!(field(&replaced, "Applicant.Email"), "beta2@example.test");
}

fn element_value(read: &Value, element_id: &str) -> Value {
    element(read, element_id)
}
