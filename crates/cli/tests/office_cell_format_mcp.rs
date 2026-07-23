use std::process::Stdio;
use std::time::Duration;

use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

#[tokio::test]
async fn native_standard_mcp_applies_typed_spreadsheet_cell_format_without_officecli() {
    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("mcp-cell-format.xlsx");
    let mut child = tokio::process::Command::new(binary())
        .args(["mcp"])
        .env("A3S_OFFICECLI_EXECUTABLE", &provider)
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
                "clientInfo": { "name": "office-cell-format-test", "version": "1" }
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

    let tools = request(
        &mut stdin,
        &mut stdout,
        serde_json::json!({"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}),
        TIMEOUT,
    )
    .await;
    let schema = tools["result"]["tools"]
        .as_array()
        .unwrap()
        .iter()
        .find(|tool| tool["name"] == "office_apply_batch")
        .unwrap()["inputSchema"]
        .to_string();
    for expected in [
        "set-cell-format",
        "numberFormat",
        "fill",
        "border",
        "diagonalUp",
        "mediumDashDotDot",
        "slantDashDot",
        "verticalAlignment",
        "wrapText",
        "textRotation",
        "shrinkToFit",
        "readingOrder",
        "right-to-left",
    ] {
        assert!(schema.contains(expected), "missing {expected}");
    }

    let created = call(
        &mut stdin,
        &mut stdout,
        3,
        "office_create",
        serde_json::json!({"session":"workbook","file":document}),
        TIMEOUT,
    )
    .await;
    assert_ne!(created["result"]["isError"], true, "{created}");

    let applied = call(
        &mut stdin,
        &mut stdout,
        4,
        "office_apply_batch",
        serde_json::json!({
            "session": "workbook",
            "mutations": [
                {
                    "operation": "set-cell-value",
                    "path": "/Sheet1/A1",
                    "value": { "type": "number", "value": "1234.5" }
                },
                {
                    "operation": "set-text-format",
                    "path": "/Sheet1/A1",
                    "format": { "bold": true }
                },
                {
                    "operation": "set-cell-format",
                    "path": "/Sheet1/A1",
                    "format": {
                        "numberFormat": "currency",
                        "fill": {
                            "kind": "solid",
                            "color": { "red": 170, "green": 187, "blue": 204 }
                        },
                        "border": {
                            "left": {
                                "kind": "line",
                                "style": "mediumDashDotDot",
                                "color": { "red": 17, "green": 34, "blue": 51 }
                            },
                            "right": { "kind": "none" },
                            "diagonal": {
                                "kind": "line",
                                "style": "slantDashDot"
                            },
                            "diagonalUp": true,
                            "diagonalDown": false
                        },
                        "verticalAlignment": "distributed",
                        "wrapText": true,
                        "textRotation": 45,
                        "indent": 2,
                        "shrinkToFit": false,
                        "readingOrder": "right-to-left"
                    }
                }
            ]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(applied["result"]["isError"], true, "{applied}");
    assert_eq!(
        applied["result"]["structuredContent"]["result"]["applied"],
        3
    );
    assert_eq!(applied["result"]["structuredContent"]["persisted"], false);

    let invalid_formula = call(
        &mut stdin,
        &mut stdout,
        40,
        "office_apply_batch",
        serde_json::json!({
            "session": "workbook",
            "mutations": [
                {
                    "operation": "set-cell-value",
                    "path": "/Sheet1/A1",
                    "value": { "type": "number", "value": "999" }
                },
                {
                    "operation": "set-cell-value",
                    "path": "/Sheet1/B1",
                    "value": { "type": "formula", "expression": "SUM(A1" }
                }
            ]
        }),
        TIMEOUT,
    )
    .await;
    assert_eq!(
        invalid_formula["result"]["isError"], true,
        "{invalid_formula}"
    );
    assert_eq!(
        invalid_formula["result"]["structuredContent"]["code"],
        "use.office.spreadsheet_formula_invalid"
    );
    assert_eq!(
        invalid_formula["result"]["structuredContent"]["details"]["characterOffset"],
        6
    );

    let read = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_get",
        serde_json::json!({
            "session": "workbook",
            "path": "/Sheet1/A1",
            "depth": 0
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(read["result"]["isError"], true, "{read}");
    let node = &read["result"]["structuredContent"]["node"];
    assert_eq!(node["text"], "1234.5");
    assert_eq!(node["format"]["bold"], "true");
    assert_eq!(node["format"]["numberFormat"], "\"$\"#,##0.00");
    assert_eq!(node["format"]["fill"], "AABBCC");
    assert_eq!(node["format"]["borderLeft"], "mediumDashDotDot");
    assert_eq!(node["format"]["borderLeftColor"], "112233");
    assert!(node["format"].get("borderRight").is_none());
    assert_eq!(node["format"]["borderDiagonal"], "slantDashDot");
    assert_eq!(node["format"]["borderDiagonalUp"], "true");
    assert_eq!(node["format"]["borderDiagonalDown"], "false");
    assert_eq!(node["format"]["verticalAlignment"], "distributed");
    assert_eq!(node["format"]["wrapText"], "true");
    assert_eq!(node["format"]["textRotation"], "45");
    assert_eq!(node["format"]["indent"], "2");
    assert_eq!(node["format"]["shrinkToFit"], "false");
    assert_eq!(node["format"]["readingOrder"], "rtl");

    let saved = call(
        &mut stdin,
        &mut stdout,
        6,
        "office_save",
        serde_json::json!({"session":"workbook"}),
        TIMEOUT,
    )
    .await;
    assert_ne!(saved["result"]["isError"], true, "{saved}");
    let closed = call(
        &mut stdin,
        &mut stdout,
        7,
        "office_close",
        serde_json::json!({"session":"workbook"}),
        TIMEOUT,
    )
    .await;
    assert_eq!(closed["result"]["structuredContent"]["closed"], true);

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
    assert!(document.is_file());
    assert!(!provider.exists());
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
