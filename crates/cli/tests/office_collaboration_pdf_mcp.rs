mod support;

use std::process::Stdio;
use std::time::Duration;

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use support::{pdf_collaboration_fixture, pdf_snapshot};
use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

#[tokio::test]
async fn mcp_sets_existing_and_new_pdf_form_values() {
    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("pdf.replica");
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
                "clientInfo": { "name": "pdf-collaboration-test", "version": "1" }
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
            "artifactId": "fixture-pdf",
            "kind": "pdf",
            "actorId": "coding-agent-pdf-mcp",
            "actorKind": "agent",
            "mode": "edit",
            "operationId": "join-browser-pdf-mcp",
            "clientId": 900014,
            "initialUpdateBase64": STANDARD.encode(pdf_collaboration_fixture())
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(created["result"]["isError"], true, "{created}");

    let name = mutate(
        &mut stdin,
        &mut stdout,
        3,
        &replica,
        "pdf-set-name-mcp",
        "Applicant.Name",
        "Katherine",
        TIMEOUT,
    )
    .await;
    assert_ne!(name["result"]["isError"], true, "{name}");
    assert_eq!(name["result"]["structuredContent"]["sequence"], 1);
    let email = mutate(
        &mut stdin,
        &mut stdout,
        4,
        &replica,
        "pdf-set-email-mcp",
        "Applicant.Email",
        "katherine@example.test",
        TIMEOUT,
    )
    .await;
    assert_ne!(email["result"]["isError"], true, "{email}");
    assert_eq!(email["result"]["structuredContent"]["sequence"], 2);

    let proposed = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_collaboration_mutate",
        serde_json::json!({
            "store": replica.to_str().unwrap(),
            "operationId": "pdf-propose-redaction-mcp",
            "actorId": "coding-agent-pdf-mcp",
            "mode": "edit",
            "artifactId": "fixture-pdf",
            "kind": "pdf",
            "mutation": {
                "type": "pdf-propose-redaction",
                "proposalId": "redaction-mcp-1",
                "pageIndex": 2,
                "rects": [
                    { "left": 12.5, "top": 25, "right": 82.5, "bottom": 45 }
                ],
                "proposedAt": "2026-08-15T06:00:00.000Z",
                "reason": "Confidential"
            }
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(proposed["result"]["isError"], true, "{proposed}");
    assert_eq!(proposed["result"]["structuredContent"]["sequence"], 3);

    let decided = call(
        &mut stdin,
        &mut stdout,
        6,
        "office_collaboration_mutate",
        serde_json::json!({
            "store": replica.to_str().unwrap(),
            "operationId": "pdf-decide-review-mcp",
            "actorId": "coding-agent-pdf-mcp",
            "mode": "edit",
            "artifactId": "fixture-pdf",
            "kind": "pdf",
            "mutation": {
                "type": "pdf-decide-review",
                "decisionId": "decision-mcp-1",
                "targetKind": "redaction",
                "targetId": "redaction-mcp-1",
                "decision": "reject",
                "createdAt": "2026-08-15T06:05:00.000Z"
            }
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(decided["result"]["isError"], true, "{decided}");
    assert_eq!(decided["result"]["structuredContent"]["sequence"], 4);

    for (id, operation_id, mutation, sequence) in [
        (
            7,
            "pdf-propose-page-rotation-mcp",
            serde_json::json!({
                "type": "pdf-propose-page-rotation",
                "pageOperationId": "page-operation-rotate-mcp",
                "pageIndices": [0, 2],
                "degrees": 270,
                "proposedAt": "2026-08-15T06:10:00.000Z"
            }),
            5,
        ),
        (
            8,
            "pdf-propose-page-deletion-mcp",
            serde_json::json!({
                "type": "pdf-propose-page-deletion",
                "pageOperationId": "page-operation-delete-mcp",
                "pageIndices": [1],
                "proposedAt": "2026-08-15T06:11:00.000Z"
            }),
            6,
        ),
        (
            9,
            "pdf-propose-page-reorder-mcp",
            serde_json::json!({
                "type": "pdf-propose-page-reorder",
                "pageOperationId": "page-operation-reorder-mcp",
                "pageOrder": [1, 2, 0],
                "proposedAt": "2026-08-15T06:12:00.000Z"
            }),
            7,
        ),
        (
            10,
            "pdf-decide-page-operation-mcp",
            serde_json::json!({
                "type": "pdf-decide-review",
                "decisionId": "decision-page-operation-mcp",
                "targetKind": "page-operation",
                "targetId": "page-operation-delete-mcp",
                "decision": "approve",
                "createdAt": "2026-08-15T06:15:00.000Z"
            }),
            8,
        ),
    ] {
        let result = call(
            &mut stdin,
            &mut stdout,
            id,
            "office_collaboration_mutate",
            serde_json::json!({
                "store": replica.to_str().unwrap(),
                "operationId": operation_id,
                "actorId": "coding-agent-pdf-mcp",
                "mode": "edit",
                "artifactId": "fixture-pdf",
                "kind": "pdf",
                "mutation": mutation
            }),
            TIMEOUT,
        )
        .await;
        assert_ne!(result["result"]["isError"], true, "{result}");
        assert_eq!(result["result"]["structuredContent"]["sequence"], sequence);
    }

    let annotation = serde_json::json!({
        "id": "annotation-mcp-1",
        "pageIndex": 1,
        "type": 15,
        "rect": {
            "origin": { "x": 20, "y": 30 },
            "size": { "width": 120, "height": 40 },
        },
        "inkList": [{
            "points": [
                { "x": 20, "y": 30 },
                { "x": 40, "y": 45 },
                { "x": 80, "y": 50 },
            ],
        }],
        "strokeColor": "#2563eb",
        "color": "#2563eb",
        "opacity": 0.8,
        "strokeWidth": 2,
        "contents": "MCP ink",
    });
    let created_annotation = call(
        &mut stdin,
        &mut stdout,
        11,
        "office_collaboration_mutate",
        serde_json::json!({
            "store": replica.to_str().unwrap(),
            "operationId": "pdf-create-annotation-mcp",
            "actorId": "coding-agent-pdf-mcp",
            "mode": "edit",
            "artifactId": "fixture-pdf",
            "kind": "pdf",
            "mutation": {
                "type": "pdf-create-annotation",
                "annotationId": "annotation-mcp-1",
                "pageIndex": 1,
                "annotation": annotation
            }
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(
        created_annotation["result"]["isError"], true,
        "{created_annotation}"
    );
    assert_eq!(
        created_annotation["result"]["structuredContent"]["sequence"],
        9
    );

    let mut next_annotation = annotation.clone();
    next_annotation["contents"] = serde_json::json!("Updated MCP ink");
    let updated_annotation = call(
        &mut stdin,
        &mut stdout,
        12,
        "office_collaboration_mutate",
        serde_json::json!({
            "store": replica.to_str().unwrap(),
            "operationId": "pdf-update-annotation-mcp",
            "actorId": "coding-agent-pdf-mcp",
            "mode": "edit",
            "artifactId": "fixture-pdf",
            "kind": "pdf",
            "mutation": {
                "type": "pdf-update-annotation",
                "annotationId": "annotation-mcp-1",
                "expectedAnnotation": annotation,
                "nextAnnotation": next_annotation
            }
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(
        updated_annotation["result"]["isError"], true,
        "{updated_annotation}"
    );
    assert_eq!(
        updated_annotation["result"]["structuredContent"]["sequence"],
        10
    );

    let deleted_annotation = call(
        &mut stdin,
        &mut stdout,
        13,
        "office_collaboration_mutate",
        serde_json::json!({
            "store": replica.to_str().unwrap(),
            "operationId": "pdf-delete-annotation-mcp",
            "actorId": "coding-agent-pdf-mcp",
            "mode": "edit",
            "artifactId": "fixture-pdf",
            "kind": "pdf",
            "mutation": {
                "type": "pdf-delete-annotation",
                "annotationId": "annotation-mcp-1",
                "expectedSource": "created",
                "expectedPageIndex": 1,
                "expectedType": 15
            }
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(
        deleted_annotation["result"]["isError"], true,
        "{deleted_annotation}"
    );
    assert_eq!(
        deleted_annotation["result"]["structuredContent"]["sequence"],
        11
    );

    let exported = call(
        &mut stdin,
        &mut stdout,
        14,
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
    let snapshot = pdf_snapshot(&update);
    assert_eq!(snapshot.form_values["Applicant.Name"], "Katherine");
    assert_eq!(
        snapshot.form_values["Applicant.Email"],
        "katherine@example.test"
    );
    assert_eq!(
        snapshot.redaction_proposals["redaction-mcp-1"]["proposedBy"],
        "coding-agent-pdf-mcp"
    );
    assert_eq!(
        snapshot.review_decisions["decision-mcp-1"],
        serde_json::json!({
            "id": "decision-mcp-1",
            "targetKind": "redaction",
            "targetId": "redaction-mcp-1",
            "decision": "reject",
            "actorId": "coding-agent-pdf-mcp",
            "createdAt": "2026-08-15T06:05:00.000Z"
        })
    );
    assert_eq!(
        snapshot.page_operations["page-operation-rotate-mcp"]["degrees"],
        270
    );
    assert_eq!(
        snapshot.page_operations["page-operation-delete-mcp"]["proposedBy"],
        "coding-agent-pdf-mcp"
    );
    assert_eq!(
        snapshot.page_operations["page-operation-reorder-mcp"]["pageOrder"],
        serde_json::json!([1, 2, 0])
    );
    assert_eq!(
        snapshot.review_decisions["decision-page-operation-mcp"]["targetKind"],
        "page-operation"
    );
    assert_eq!(snapshot.annotations["annotation-mcp-1"]["deleted"], true);
    assert_eq!(
        snapshot.annotations["annotation-mcp-1"]["annotation"]["contents"],
        "Updated MCP ink"
    );
    assert_eq!(
        snapshot.annotations["annotation-mcp-1"]["annotation"]["inkList"][0]["points"][2]["x"],
        80.0
    );
    assert_eq!(snapshot.record_claims.len(), 7);

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

#[allow(clippy::too_many_arguments)]
async fn mutate(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut BufReader<tokio::process::ChildStdout>,
    id: u32,
    replica: &std::path::Path,
    operation_id: &str,
    field_id: &str,
    value: &str,
    timeout: Duration,
) -> serde_json::Value {
    call(
        stdin,
        stdout,
        id,
        "office_collaboration_mutate",
        serde_json::json!({
            "store": replica.to_str().unwrap(),
            "operationId": operation_id,
            "actorId": "coding-agent-pdf-mcp",
            "mode": "edit",
            "artifactId": "fixture-pdf",
            "kind": "pdf",
            "mutation": {
                "type": "pdf-set-form-value",
                "fieldId": field_id,
                "value": value
            }
        }),
        timeout,
    )
    .await
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
