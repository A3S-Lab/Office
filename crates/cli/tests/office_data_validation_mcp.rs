use std::process::Stdio;
use std::time::Duration;

use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

#[tokio::test]
async fn standard_mcp_manages_spreadsheet_data_validation_without_officecli() {
    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("mcp-validation.xlsx");
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
                "clientInfo": { "name": "office-validation-test", "version": "1" }
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
    assert!(schema.contains("add-data-validation"));
    assert!(schema.contains("set-data-validation"));

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

    let added = call(
        &mut stdin,
        &mut stdout,
        4,
        "office_apply_batch",
        serde_json::json!({
            "session": "workbook",
            "mutations": [{
                "operation": "add-data-validation",
                "sheet": "/Sheet1",
                "validation": {
                    "type": "list",
                    "ranges": ["A2:A20", "C2:C20"],
                    "formula1": "Draft,Review,Approved",
                    "inCellDropdown": false,
                    "promptTitle": "Status",
                    "prompt": "Choose a workflow state"
                }
            }]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(added["result"]["isError"], true, "{added}");
    assert_eq!(
        added["result"]["structuredContent"]["result"]["paths"][0],
        "/Sheet1/dataValidation[1]"
    );
    assert_eq!(added["result"]["structuredContent"]["persisted"], false);

    let read = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_get",
        serde_json::json!({
            "session": "workbook",
            "path": "/Sheet1/dataValidation[1]",
            "depth": 0
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(read["result"]["isError"], true, "{read}");
    let node = &read["result"]["structuredContent"]["node"];
    assert_eq!(node["type"], "data-validation");
    assert_eq!(node["format"]["type"], "list");
    assert_eq!(node["format"]["ref"], "A2:A20 C2:C20");

    let updated = call(
        &mut stdin,
        &mut stdout,
        6,
        "office_apply_batch",
        serde_json::json!({
            "session": "workbook",
            "mutations": [{
                "operation": "set-data-validation",
                "path": "/Sheet1/dataValidation[1]",
                "validation": {
                    "type": "whole",
                    "ranges": ["B2:B50"],
                    "operator": "between",
                    "formula1": "18",
                    "formula2": "120",
                    "allowBlank": false,
                    "errorStyle": "warning"
                }
            }]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(updated["result"]["isError"], true, "{updated}");

    let removed = call(
        &mut stdin,
        &mut stdout,
        7,
        "office_apply_batch",
        serde_json::json!({
            "session": "workbook",
            "mutations": [{
                "operation": "remove",
                "path": "/Sheet1/dataValidation[1]"
            }]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(removed["result"]["isError"], true, "{removed}");

    let saved = call(
        &mut stdin,
        &mut stdout,
        8,
        "office_save",
        serde_json::json!({"session":"workbook"}),
        TIMEOUT,
    )
    .await;
    assert_ne!(saved["result"]["isError"], true, "{saved}");
    let closed = call(
        &mut stdin,
        &mut stdout,
        9,
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
    tokio::time::timeout(timeout, stdout.read_line(&mut line))
        .await
        .unwrap()
        .unwrap();
    assert!(!line.is_empty());
    serde_json::from_str(&line).unwrap()
}
