use std::process::Stdio;
use std::time::Duration;

use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

#[tokio::test]
async fn standard_mcp_manages_spreadsheet_filters_without_officecli() {
    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("mcp-filter.xlsx");
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
                "clientInfo": { "name": "office-filter-test", "version": "1" }
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
    assert!(schema.contains("add-spreadsheet-auto-filter"));
    assert!(schema.contains("set-spreadsheet-auto-filter"));
    assert!(schema.contains("does-not-contain"));
    assert!(schema.contains("above-average"));

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
                "operation": "add-spreadsheet-auto-filter",
                "sheet": "/Sheet1",
                "filter": {
                    "range": "A1:C50",
                    "columns": [
                        {
                            "column": 0,
                            "criteria": {
                                "type": "values",
                                "values": ["Open", "Closed"],
                                "includeBlanks": true
                            }
                        },
                        {
                            "column": 2,
                            "criteria": {"type": "between", "lower": "10", "upper": "100"}
                        }
                    ]
                }
            }]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(added["result"]["isError"], true, "{added}");
    assert_eq!(
        added["result"]["structuredContent"]["result"]["paths"][0],
        "/Sheet1/autofilter"
    );

    let read = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_get",
        serde_json::json!({
            "session": "workbook",
            "path": "/Sheet1/autofilter",
            "depth": 2
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(read["result"]["isError"], true, "{read}");
    let node = &read["result"]["structuredContent"]["node"];
    assert_eq!(node["type"], "auto-filter");
    assert_eq!(node["format"]["ref"], "A1:C50");
    assert_eq!(node["children"][0]["children"][1]["text"], "Closed");

    let queried = call(
        &mut stdin,
        &mut stdout,
        6,
        "office_query",
        serde_json::json!({
            "session": "workbook",
            "selector": "filtercolumn[criteriaType=between]"
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(queried["result"]["isError"], true, "{queried}");
    assert_eq!(queried["result"]["structuredContent"]["matches"], 1);

    let updated = call(
        &mut stdin,
        &mut stdout,
        7,
        "office_apply_batch",
        serde_json::json!({
            "session": "workbook",
            "mutations": [{
                "operation": "set-spreadsheet-auto-filter",
                "path": "/Sheet1/autofilter",
                "filter": {
                    "range": "B2:D30",
                    "columns": [{
                        "column": 1,
                        "criteria": {"type": "dynamic", "kind": "this-month"}
                    }]
                }
            }]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(updated["result"]["isError"], true, "{updated}");

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
            "path": "/Sheet1/autofilter",
            "depth": 1
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(persisted["result"]["isError"], true, "{persisted}");
    assert_eq!(
        persisted["result"]["structuredContent"]["node"]["format"]["ref"],
        "B2:D30"
    );

    let removed = call(
        &mut stdin,
        &mut stdout,
        12,
        "office_apply_batch",
        serde_json::json!({
            "session": "reopened",
            "mutations": [{
                "operation": "remove",
                "path": "/Sheet1/autofilter"
            }]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(removed["result"]["isError"], true, "{removed}");
    let saved = call(
        &mut stdin,
        &mut stdout,
        13,
        "office_save",
        serde_json::json!({"session":"reopened"}),
        TIMEOUT,
    )
    .await;
    assert_ne!(saved["result"]["isError"], true, "{saved}");
    let closed = call(
        &mut stdin,
        &mut stdout,
        14,
        "office_close",
        serde_json::json!({"session":"reopened"}),
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
