use std::process::Stdio;
use std::time::Duration;

use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

#[tokio::test]
async fn standard_mcp_sorts_spreadsheet_rows_before_save_and_after_reopen() {
    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("mcp-sort.xlsx");
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
                "clientInfo": { "name": "office-sort-test", "version": "1" }
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
    assert!(schema.contains("sort-spreadsheet-range"));
    assert!(schema.contains("caseSensitive"));
    assert!(schema.contains("descending"));

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
                {"operation":"set-cell-value","path":"/Sheet1/A1","value":{"type":"text","value":"Name"}},
                {"operation":"set-cell-value","path":"/Sheet1/B1","value":{"type":"text","value":"Rank"}},
                {"operation":"set-cell-value","path":"/Sheet1/A2","value":{"type":"text","value":"Beta"}},
                {"operation":"set-cell-value","path":"/Sheet1/B2","value":{"type":"number","value":"2"}},
                {"operation":"set-cell-value","path":"/Sheet1/A3","value":{"type":"text","value":"Alpha"}},
                {"operation":"set-cell-value","path":"/Sheet1/B3","value":{"type":"number","value":"1"}},
                {
                    "operation": "sort-spreadsheet-range",
                    "path": "/Sheet1/A1:B3",
                    "sort": {
                        "keys": [{"column":"B","direction":"ascending"}],
                        "header": true
                    }
                }
            ]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(applied["result"]["isError"], true, "{applied}");
    assert_eq!(applied["result"]["structuredContent"]["persisted"], false);
    assert_eq!(
        applied["result"]["structuredContent"]["result"]["paths"][6],
        "/Sheet1/sort"
    );

    let unsaved_cell = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_get",
        serde_json::json!({"session":"workbook","path":"/Sheet1/A2","depth":0}),
        TIMEOUT,
    )
    .await;
    assert_eq!(
        unsaved_cell["result"]["structuredContent"]["node"]["text"],
        "Alpha"
    );
    let unsaved_state = call(
        &mut stdin,
        &mut stdout,
        6,
        "office_get",
        serde_json::json!({"session":"workbook","path":"/Sheet1/sort","depth":1}),
        TIMEOUT,
    )
    .await;
    assert_eq!(
        unsaved_state["result"]["structuredContent"]["node"]["format"]["ref"],
        "A1:B3"
    );

    call(
        &mut stdin,
        &mut stdout,
        7,
        "office_save",
        serde_json::json!({"session":"workbook"}),
        TIMEOUT,
    )
    .await;
    call(
        &mut stdin,
        &mut stdout,
        8,
        "office_close",
        serde_json::json!({"session":"workbook"}),
        TIMEOUT,
    )
    .await;
    let reopened = call(
        &mut stdin,
        &mut stdout,
        9,
        "office_open",
        serde_json::json!({"session":"reopened","file":document}),
        TIMEOUT,
    )
    .await;
    assert_ne!(reopened["result"]["isError"], true, "{reopened}");
    let persisted = call(
        &mut stdin,
        &mut stdout,
        10,
        "office_get",
        serde_json::json!({"session":"reopened","path":"/Sheet1/sort","depth":1}),
        TIMEOUT,
    )
    .await;
    assert_eq!(
        persisted["result"]["structuredContent"]["node"]["children"][0]["format"]["direction"],
        "ascending"
    );

    let removed = call(
        &mut stdin,
        &mut stdout,
        11,
        "office_apply_batch",
        serde_json::json!({
            "session":"reopened",
            "mutations":[{"operation":"remove","path":"/Sheet1/sort"}]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(removed["result"]["isError"], true, "{removed}");
    let retained = call(
        &mut stdin,
        &mut stdout,
        12,
        "office_get",
        serde_json::json!({"session":"reopened","path":"/Sheet1/A2","depth":0}),
        TIMEOUT,
    )
    .await;
    assert_eq!(
        retained["result"]["structuredContent"]["node"]["text"],
        "Alpha"
    );
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
