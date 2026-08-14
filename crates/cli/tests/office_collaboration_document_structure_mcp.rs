mod support;

use std::process::Stdio;
use std::time::Duration;

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use support::{complex_document_fixture, document_snapshot, DocumentRow};
use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

#[tokio::test]
async fn mcp_mutates_a_paragraph_in_a_nested_table_cell() {
    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("complex-document.replica");
    let mut child = tokio::process::Command::new(binary())
        .args(["mcp"])
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .kill_on_drop(true)
        .spawn()
        .unwrap();
    let mut stdin = child.stdin.take().unwrap();
    let mut stdout = BufReader::new(child.stdout.take().unwrap());
    let mut stderr = child.stderr.take().unwrap();
    request(
        &mut stdin,
        &mut stdout,
        serde_json::json!({
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2025-06-18",
                "capabilities": {},
                "clientInfo": { "name": "complex-document-test", "version": "1" }
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

    let created = call(
        &mut stdin,
        &mut stdout,
        2,
        "office_collaboration_create",
        serde_json::json!({
            "store": replica.to_str().unwrap(),
            "artifactId": "fixture-complex-document",
            "kind": "document",
            "actorId": "coding-agent-complex-mcp",
            "actorKind": "agent",
            "mode": "edit",
            "operationId": "join-complex-document-mcp",
            "clientId": 900012,
            "initialUpdateBase64": STANDARD.encode(complex_document_fixture())
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(created["result"]["isError"], true, "{created}");

    let inserted = call(
        &mut stdin,
        &mut stdout,
        3,
        "office_collaboration_mutate",
        serde_json::json!({
            "store": replica.to_str().unwrap(),
            "operationId": "mcp-insert-nested-cell-paragraph",
            "actorId": "coding-agent-complex-mcp",
            "mode": "edit",
            "artifactId": "fixture-complex-document",
            "kind": "document",
            "mutation": {
                "type": "document-insert-paragraph",
                "anchorParagraphId": "00000311",
                "position": "after",
                "paragraphId": "00000315",
                "textId": "00000316",
                "text": "MCP cell"
            }
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(inserted["result"]["isError"], true, "{inserted}");
    assert_eq!(inserted["result"]["structuredContent"]["sequence"], 1);

    let exported = call(
        &mut stdin,
        &mut stdout,
        4,
        "office_collaboration_diff",
        serde_json::json!({"store": replica.to_str().unwrap()}),
        TIMEOUT,
    )
    .await;
    let update = STANDARD
        .decode(
            exported["result"]["structuredContent"]["updateBase64"]
                .as_str()
                .unwrap(),
        )
        .unwrap();
    let snapshot = document_snapshot(&update);
    assert!(snapshot.paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000315"
            && paragraph.text_id == "00000316"
            && paragraph.parent_tag == "tableCell"
            && paragraph.text == "MCP cell"
    }));
    assert_eq!(
        snapshot.rows,
        vec![row("00000201", "00000203"), row("00000301", "00000303")]
    );

    drop(stdin);
    let status = tokio::time::timeout(TIMEOUT, child.wait())
        .await
        .unwrap()
        .unwrap();
    assert!(status.success());
    let mut diagnostics = Vec::new();
    stderr.read_to_end(&mut diagnostics).await.unwrap();
    assert!(
        diagnostics.is_empty(),
        "{}",
        String::from_utf8_lossy(&diagnostics)
    );
}

fn row(row_id: &str, row_text_id: &str) -> DocumentRow {
    DocumentRow {
        row_id: row_id.to_owned(),
        row_text_id: row_text_id.to_owned(),
    }
}

async fn call(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut BufReader<tokio::process::ChildStdout>,
    id: u32,
    name: &str,
    arguments: serde_json::Value,
    timeout: Duration,
) -> serde_json::Value {
    request(
        stdin,
        stdout,
        serde_json::json!({
            "jsonrpc": "2.0",
            "id": id,
            "method": "tools/call",
            "params": { "name": name, "arguments": arguments }
        }),
        timeout,
    )
    .await
}

async fn request(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut BufReader<tokio::process::ChildStdout>,
    value: serde_json::Value,
    timeout: Duration,
) -> serde_json::Value {
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
