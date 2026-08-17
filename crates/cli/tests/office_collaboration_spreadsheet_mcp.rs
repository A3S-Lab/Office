mod support;

use std::process::Stdio;
use std::time::Duration;

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use support::{spreadsheet_cell, spreadsheet_collaboration_fixture};
use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

#[tokio::test]
async fn mcp_sets_creates_and_deletes_browser_compatible_spreadsheet_cells() {
    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("spreadsheet.replica");
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
    let fixture_update = spreadsheet_collaboration_fixture();
    let header = spreadsheet_cell(&fixture_update, "sheet-data", 0, 0).unwrap();
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
                "clientInfo": { "name": "spreadsheet-collaboration-test", "version": "1" }
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
            "artifactId": "fixture-spreadsheet",
            "kind": "spreadsheet",
            "actorId": "coding-agent-spreadsheet-mcp",
            "actorKind": "agent",
            "mode": "edit",
            "operationId": "join-spreadsheet-mcp",
            "clientId": 900015,
            "initialUpdateBase64": STANDARD.encode(&fixture_update)
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(created["result"]["isError"], true, "{created}");

    let set = call(
        &mut stdin,
        &mut stdout,
        3,
        "office_collaboration_mutate",
        mutation_arguments(
            &replica,
            "spreadsheet-set-cell-mcp",
            serde_json::json!({
                "type": "spreadsheet-set-cell",
                "sheetId": "sheet-data",
                "row": 1,
                "column": 0,
                "expectedCell": {
                    "v": 10,
                    "m": "10",
                    "ct": { "fa": "0.00", "t": "n" }
                },
                "nextCell": {
                    "v": 14,
                    "m": "14",
                    "f": "=7*2",
                    "ct": { "fa": "0.00", "t": "n" }
                }
            }),
        ),
        TIMEOUT,
    )
    .await;
    assert_ne!(set["result"]["isError"], true, "{set}");
    assert_eq!(set["result"]["structuredContent"]["sequence"], 1);

    let batch = call(
        &mut stdin,
        &mut stdout,
        4,
        "office_collaboration_mutate",
        mutation_arguments(
            &replica,
            "spreadsheet-batch-cells-mcp",
            serde_json::json!({
                "type": "spreadsheet-batch-cells",
                "sheetId": "sheet-data",
                "changes": [
                    {
                        "row": 1,
                        "column": 0,
                        "expectedCell": {
                            "v": 14,
                            "m": "14",
                            "f": "=7*2",
                            "ct": { "fa": "0.00", "t": "n" }
                        },
                        "nextCell": {
                            "v": 18,
                            "m": "18",
                            "f": "=9*2",
                            "ct": { "fa": "0.00", "t": "n" }
                        }
                    },
                    {
                        "row": 3,
                        "column": 4,
                        "expectedCell": null,
                        "nextCell": { "v": "MCP batch", "m": "MCP batch" }
                    },
                    {
                        "row": 0,
                        "column": 0,
                        "expectedCell": header,
                        "nextCell": null
                    }
                ]
            }),
        ),
        TIMEOUT,
    )
    .await;
    assert_ne!(batch["result"]["isError"], true, "{batch}");
    assert_eq!(batch["result"]["structuredContent"]["sequence"], 2);

    let create = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_collaboration_mutate",
        mutation_arguments(
            &replica,
            "spreadsheet-create-cell-mcp",
            serde_json::json!({
                "type": "spreadsheet-set-cell",
                "sheetId": "sheet-empty",
                "row": 8,
                "column": 2,
                "expectedCell": null,
                "nextCell": { "v": "MCP sparse", "m": "MCP sparse" }
            }),
        ),
        TIMEOUT,
    )
    .await;
    assert_ne!(create["result"]["isError"], true, "{create}");
    assert_eq!(create["result"]["structuredContent"]["sequence"], 3);

    let delete = call(
        &mut stdin,
        &mut stdout,
        6,
        "office_collaboration_mutate",
        mutation_arguments(
            &replica,
            "spreadsheet-delete-cell-mcp",
            serde_json::json!({
                "type": "spreadsheet-delete-cell",
                "sheetId": "sheet-sparse",
                "row": 5,
                "column": 3,
                "expectedCell": { "v": "edge", "m": "edge" }
            }),
        ),
        TIMEOUT,
    )
    .await;
    assert_ne!(delete["result"]["isError"], true, "{delete}");
    assert_eq!(delete["result"]["structuredContent"]["sequence"], 4);

    let exported = call(
        &mut stdin,
        &mut stdout,
        7,
        "office_collaboration_diff",
        serde_json::json!({ "store": replica.to_str().unwrap() }),
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
    assert_eq!(
        spreadsheet_cell(&update, "sheet-data", 1, 0),
        Some(serde_json::json!({
            "v": 18,
            "m": "18",
            "f": "=9*2",
            "ct": { "fa": "0.00", "t": "n" },
        }))
    );
    assert_eq!(
        spreadsheet_cell(&update, "sheet-data", 3, 4),
        Some(serde_json::json!({ "v": "MCP batch", "m": "MCP batch" }))
    );
    assert_eq!(spreadsheet_cell(&update, "sheet-data", 0, 0), None);
    assert_eq!(
        spreadsheet_cell(&update, "sheet-empty", 8, 2),
        Some(serde_json::json!({ "v": "MCP sparse", "m": "MCP sparse" }))
    );
    assert_eq!(spreadsheet_cell(&update, "sheet-sparse", 5, 3), None);

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

fn mutation_arguments(
    replica: &std::path::Path,
    operation_id: &str,
    mutation: serde_json::Value,
) -> serde_json::Value {
    serde_json::json!({
        "store": replica.to_str().unwrap(),
        "operationId": operation_id,
        "actorId": "coding-agent-spreadsheet-mcp",
        "mode": "edit",
        "artifactId": "fixture-spreadsheet",
        "kind": "spreadsheet",
        "mutation": mutation
    })
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
