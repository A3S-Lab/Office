use std::process::Stdio;
use std::time::Duration;

use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

#[tokio::test]
async fn standard_mcp_recalculates_formulas_atomically_without_officecli() {
    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("mcp-formulas.xlsx");
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
                "clientInfo": { "name": "office-formula-test", "version": "1" }
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
    assert!(schema.contains("recalculate-spreadsheet-formulas"));

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
                    "value": { "type": "number", "value": "2" }
                },
                {
                    "operation": "set-cell-value",
                    "path": "/Sheet1/B1",
                    "value": { "type": "formula", "expression": "A1*3" }
                },
                {
                    "operation": "set-cell-value",
                    "path": "/Sheet1/C1",
                    "value": { "type": "formula", "expression": "SEQUENCE(2,2,1,1)" }
                },
                {
                    "operation": "recalculate-spreadsheet-formulas"
                }
            ]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(applied["result"]["isError"], true, "{applied}");
    let content = &applied["result"]["structuredContent"];
    assert_eq!(content["persisted"], false);
    assert_eq!(content["result"]["applied"], 4);
    assert_eq!(
        content["result"]["spreadsheetCalculations"][0]["formulaCount"],
        2
    );
    assert_eq!(
        content["result"]["spreadsheetCalculations"][0]["spillCellCount"],
        3
    );

    let spill = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_get",
        serde_json::json!({"session":"workbook","path":"/Sheet1/D2","depth":0}),
        TIMEOUT,
    )
    .await;
    assert_eq!(spill["result"]["structuredContent"]["node"]["text"], "4");

    let rejected = call(
        &mut stdin,
        &mut stdout,
        6,
        "office_apply_batch",
        serde_json::json!({
            "session": "workbook",
            "mutations": [
                {
                    "operation": "set-cell-value",
                    "path": "/Sheet1/E1",
                    "value": { "type": "text", "value": "must roll back" }
                },
                {
                    "operation": "set-cell-value",
                    "path": "/Sheet1/F1",
                    "value": { "type": "formula", "expression": "SHELL(\"unsafe\")" }
                },
                {
                    "operation": "recalculate-spreadsheet-formulas"
                }
            ]
        }),
        TIMEOUT,
    )
    .await;
    assert_eq!(rejected["result"]["isError"], true, "{rejected}");
    assert_eq!(
        rejected["result"]["structuredContent"]["code"],
        "use.office.spreadsheet_formula_function_unsupported"
    );
    assert_eq!(
        rejected["result"]["structuredContent"]["suggestion"],
        "Use a closed-registry function listed on recalculate-spreadsheet-formulas. Do not calculate the formula outside the native engine."
    );
    let rolled_back = call(
        &mut stdin,
        &mut stdout,
        7,
        "office_get",
        serde_json::json!({"session":"workbook","path":"/Sheet1/E1","depth":0}),
        TIMEOUT,
    )
    .await;
    assert_eq!(rolled_back["result"]["isError"], true, "{rolled_back}");
    assert_eq!(
        rolled_back["result"]["structuredContent"]["code"],
        "use.office.node_not_found"
    );

    let qualified = call(
        &mut stdin,
        &mut stdout,
        8,
        "office_apply_batch",
        serde_json::json!({
            "session": "workbook",
            "mutations": [
                {
                    "operation": "set-cell-value",
                    "path": "/Sheet1/G1",
                    "value": { "type": "formula", "expression": "Sheet1!SUM(1,2)" }
                },
                {
                    "operation": "recalculate-spreadsheet-formulas"
                }
            ]
        }),
        TIMEOUT,
    )
    .await;
    assert_eq!(qualified["result"]["isError"], true, "{qualified}");
    assert_eq!(
        qualified["result"]["structuredContent"]["code"],
        "use.office.spreadsheet_formula_function_unsupported"
    );
    assert_eq!(
        qualified["result"]["structuredContent"]["suggestion"],
        "Call closed-registry functions with bare names (no Sheet! or workbook qualifier). Recalculate in-process; do not evaluate the formula outside the native engine."
    );

    let external = call(
        &mut stdin,
        &mut stdout,
        9,
        "office_apply_batch",
        serde_json::json!({
            "session": "workbook",
            "mutations": [
                {
                    "operation": "set-cell-value",
                    "path": "/Sheet1/H1",
                    "value": { "type": "formula", "expression": "'[Book.xlsx]Sheet1'!A1" }
                },
                {
                    "operation": "recalculate-spreadsheet-formulas"
                }
            ]
        }),
        TIMEOUT,
    )
    .await;
    assert_eq!(external["result"]["isError"], true, "{external}");
    assert_eq!(
        external["result"]["structuredContent"]["code"],
        "use.office.spreadsheet_formula_external_reference_unsupported"
    );
    assert_eq!(
        external["result"]["structuredContent"]["suggestion"],
        "Rewrite the formula to in-workbook sheet ranges or values. Native calculation never opens other workbooks. Recalculate in-process; do not evaluate the formula outside the native engine."
    );

    let named_3d = call(
        &mut stdin,
        &mut stdout,
        10,
        "office_apply_batch",
        serde_json::json!({
            "session": "workbook",
            "mutations": [
                {
                    "operation": "set-cell-value",
                    "path": "/Sheet1/I1",
                    "value": { "type": "formula", "expression": "Sheet1:Sheet2!Revenue" }
                },
                {
                    "operation": "recalculate-spreadsheet-formulas"
                }
            ]
        }),
        TIMEOUT,
    )
    .await;
    assert_eq!(named_3d["result"]["isError"], true, "{named_3d}");
    assert_eq!(
        named_3d["result"]["structuredContent"]["code"],
        "use.office.spreadsheet_formula_named_reference_unsupported"
    );
    assert_eq!(
        named_3d["result"]["structuredContent"]["suggestion"],
        "Resolve names with a single sheet scope or workbook scope (no Sheet1:Sheet2!Name). Rewrite to in-workbook ranges or a non-3D defined name; recalculate in-process; do not evaluate the formula outside the native engine."
    );

    call(
        &mut stdin,
        &mut stdout,
        11,
        "office_close",
        serde_json::json!({"session":"workbook","discard":true}),
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
            "params": {"name": name, "arguments": arguments}
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
    stdin
        .write_all(format!("{value}\n").as_bytes())
        .await
        .unwrap();
    stdin.flush().await.unwrap();
    let mut line = String::new();
    tokio::time::timeout(timeout, stdout.read_line(&mut line))
        .await
        .unwrap()
        .unwrap();
    assert!(!line.is_empty());
    serde_json::from_str(&line).unwrap()
}
