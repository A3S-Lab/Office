use std::process::Stdio;
use std::time::Duration;

use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

#[tokio::test]
async fn standard_mcp_manages_spreadsheet_tables_without_officecli() {
    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("mcp-tables.xlsx");
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
                "clientInfo": { "name": "office-table-test", "version": "1" }
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
    assert!(schema.contains("add-spreadsheet-table"));
    assert!(schema.contains("set-spreadsheet-table"));
    assert!(schema.contains("showColumnStripes"));
    assert!(schema.contains("filters"));

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
                "operation": "add-spreadsheet-table",
                "sheet": "/Sheet1",
                "table": {
                    "name": "Sales",
                    "range": "A1:C4",
                    "columns": [
                        {"name": "Name"},
                        {"name": "Qty"},
                        {"name": "Price"}
                    ],
                    "filters": [{
                        "column": 1,
                        "criteria": {"type": "top", "count": 10}
                    }],
                    "style": {"family": "medium", "number": 4},
                    "showLastColumn": true
                }
            }]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(added["result"]["isError"], true, "{added}");
    assert_eq!(
        added["result"]["structuredContent"]["result"]["paths"],
        serde_json::json!(["/Sheet1/table[1]"])
    );
    assert_eq!(added["result"]["structuredContent"]["persisted"], false);

    let read = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_get",
        serde_json::json!({
            "session": "workbook",
            "path": "/Sheet1/table[1]",
            "depth": 2
        }),
        TIMEOUT,
    )
    .await;
    let node = &read["result"]["structuredContent"]["node"];
    assert_eq!(node["type"], "table");
    assert_eq!(node["format"]["name"], "Sales");
    assert_eq!(node["children"][2]["text"], "Price");
    assert_eq!(node["children"][3]["type"], "auto-filter");
    assert_eq!(
        node["children"][3]["children"][0]["format"]["criteriaType"],
        "top"
    );

    let updated = call(
        &mut stdin,
        &mut stdout,
        6,
        "office_apply_batch",
        serde_json::json!({
            "session": "workbook",
            "mutations": [{
                "operation": "set-spreadsheet-table",
                "path": "/Sheet1/table[1]",
                "table": {
                    "name": "Inventory",
                    "displayName": "InventoryView",
                    "range": "B2:D6",
                    "columns": [
                        {"name": "Item"},
                        {"name": "Units"},
                        {"name": "Cost"}
                    ],
                    "totalsRow": true,
                    "style": {"family": "dark", "number": 2},
                    "showRowStripes": false,
                    "showColumnStripes": true
                }
            }]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(updated["result"]["isError"], true, "{updated}");

    let queried = call(
        &mut stdin,
        &mut stdout,
        7,
        "office_query",
        serde_json::json!({
            "session": "workbook",
            "selector": "table[name=Inventory]",
            "limit": 10
        }),
        TIMEOUT,
    )
    .await;
    assert_eq!(queried["result"]["structuredContent"]["matches"], 1);

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
    call(
        &mut stdin,
        &mut stdout,
        9,
        "office_close",
        serde_json::json!({"session":"workbook"}),
        TIMEOUT,
    )
    .await;

    let reopened = call(
        &mut stdin,
        &mut stdout,
        10,
        "office_open",
        serde_json::json!({"session":"reopened","file":document}),
        TIMEOUT,
    )
    .await;
    assert_ne!(reopened["result"]["isError"], true, "{reopened}");
    let persisted = call(
        &mut stdin,
        &mut stdout,
        11,
        "office_get",
        serde_json::json!({
            "session": "reopened",
            "path": "/Sheet1/table[1]",
            "depth": 0
        }),
        TIMEOUT,
    )
    .await;
    assert_eq!(
        persisted["result"]["structuredContent"]["node"]["format"]["displayName"],
        "InventoryView"
    );
    let removed = call(
        &mut stdin,
        &mut stdout,
        12,
        "office_apply_batch",
        serde_json::json!({
            "session": "reopened",
            "mutations": [{"operation":"remove","path":"/Sheet1/table[1]"}]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(removed["result"]["isError"], true, "{removed}");
    call(
        &mut stdin,
        &mut stdout,
        13,
        "office_save",
        serde_json::json!({"session":"reopened"}),
        TIMEOUT,
    )
    .await;
    call(
        &mut stdin,
        &mut stdout,
        14,
        "office_close",
        serde_json::json!({"session":"reopened"}),
        TIMEOUT,
    )
    .await;

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
