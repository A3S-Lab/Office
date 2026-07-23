use std::process::Stdio;
use std::time::Duration;

use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

#[tokio::test]
async fn standard_mcp_imports_delimited_content_before_save_and_after_reopen() {
    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("mcp-import.xlsx");
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
                "clientInfo": { "name": "office-import-test", "version": "1" }
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
    assert!(schema.contains("import-spreadsheet-delimited"));
    assert!(schema.contains("set-spreadsheet-frozen-pane"));
    assert!(schema.contains("startCell"));

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
            "mutations": [{
                "operation": "import-spreadsheet-delimited",
                "sheet": "/Sheet1",
                "import": {
                    "content": "Name,Amount,Date\nAlpha,42,2026-07-17\nBeta,TRUE,2026-07-18",
                    "format": "csv",
                    "header": true,
                    "startCell": "B2"
                }
            }]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(applied["result"]["isError"], true, "{applied}");
    let content = &applied["result"]["structuredContent"];
    assert_eq!(content["persisted"], false);
    assert_eq!(content["result"]["paths"][0], "/Sheet1/B2:D4");
    assert_eq!(content["result"]["spreadsheetImports"][0]["rowCount"], 3);
    assert_eq!(
        content["result"]["spreadsheetImports"][0]["freezePath"],
        "/Sheet1/freeze"
    );

    let unsaved_date = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_get",
        serde_json::json!({"session":"workbook","path":"/Sheet1/D3","depth":0}),
        TIMEOUT,
    )
    .await;
    assert_eq!(
        unsaved_date["result"]["structuredContent"]["node"]["format"]["numberFormat"],
        "yyyy-mm-dd"
    );
    let unsaved_freeze = call(
        &mut stdin,
        &mut stdout,
        6,
        "office_get",
        serde_json::json!({"session":"workbook","path":"/Sheet1/freeze","depth":0}),
        TIMEOUT,
    )
    .await;
    assert_eq!(
        unsaved_freeze["result"]["structuredContent"]["node"]["format"]["topLeftCell"],
        "B3"
    );
    let unsaved_filter = call(
        &mut stdin,
        &mut stdout,
        7,
        "office_get",
        serde_json::json!({"session":"workbook","path":"/Sheet1/autofilter","depth":0}),
        TIMEOUT,
    )
    .await;
    assert_eq!(
        unsaved_filter["result"]["structuredContent"]["node"]["format"]["ref"],
        "B2:D4"
    );
    let pane_set = call(
        &mut stdin,
        &mut stdout,
        70,
        "office_apply_batch",
        serde_json::json!({
            "session": "workbook",
            "mutations": [{
                "operation": "set-spreadsheet-frozen-pane",
                "sheet": "/Sheet1",
                "pane": {
                    "frozenRows": 1,
                    "frozenColumns": 1,
                    "topLeftCell": "B2"
                }
            }]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(pane_set["result"]["isError"], true, "{pane_set}");
    let replaced_pane = call(
        &mut stdin,
        &mut stdout,
        71,
        "office_get",
        serde_json::json!({"session":"workbook","path":"/Sheet1/freeze","depth":0}),
        TIMEOUT,
    )
    .await;
    assert_eq!(
        replaced_pane["result"]["structuredContent"]["node"]["format"]["frozenRows"],
        "1"
    );
    assert_eq!(
        replaced_pane["result"]["structuredContent"]["node"]["format"]["frozenColumns"],
        "1"
    );
    assert_eq!(
        replaced_pane["result"]["structuredContent"]["node"]["format"]["topLeftCell"],
        "B2"
    );

    call(
        &mut stdin,
        &mut stdout,
        8,
        "office_save",
        serde_json::json!({"session":"workbook"}),
        TIMEOUT,
    )
    .await;
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
        serde_json::json!({"session":"reopened","path":"/Sheet1/B4","depth":0}),
        TIMEOUT,
    )
    .await;
    assert_eq!(
        persisted["result"]["structuredContent"]["node"]["text"],
        "Beta"
    );
    let persisted_pane = call(
        &mut stdin,
        &mut stdout,
        72,
        "office_get",
        serde_json::json!({"session":"reopened","path":"/Sheet1/freeze","depth":0}),
        TIMEOUT,
    )
    .await;
    assert_eq!(
        persisted_pane["result"]["structuredContent"]["node"]["format"]["topLeftCell"],
        "B2"
    );
    let removed = call(
        &mut stdin,
        &mut stdout,
        12,
        "office_apply_batch",
        serde_json::json!({
            "session":"reopened",
            "mutations":[{"operation":"remove","path":"/Sheet1/freeze"}]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(removed["result"]["isError"], true, "{removed}");
    call(
        &mut stdin,
        &mut stdout,
        13,
        "office_close",
        serde_json::json!({"session":"reopened","discard":true}),
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
