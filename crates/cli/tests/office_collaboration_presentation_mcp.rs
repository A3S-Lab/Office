mod support;

use std::process::Stdio;
use std::time::Duration;

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use support::{
    presentation_collaboration_fixture, presentation_element, presentation_element_order,
    presentation_element_tombstoned, presentation_scene_element, presentation_slide_body_element,
    presentation_slide_title_element,
};
use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

#[tokio::test]
async fn mcp_manages_browser_compatible_presentation_scene_elements() {
    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("presentation.replica");
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
                "clientInfo": { "name": "presentation-collaboration-test", "version": "1" }
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
            "artifactId": "fixture-presentation",
            "kind": "presentation",
            "actorId": "coding-agent-presentation-mcp",
            "actorKind": "agent",
            "mode": "edit",
            "operationId": "join-presentation-mcp",
            "clientId": 900017,
            "initialUpdateBase64": STANDARD.encode(presentation_collaboration_fixture())
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(created["result"]["isError"], true, "{created}");

    let expected = presentation_slide_title_element();
    let mut next = expected.clone();
    next["text"] = serde_json::json!("MCP shared title");
    next["fill"] = serde_json::json!("#E0E7FF");
    let updated = call(
        &mut stdin,
        &mut stdout,
        3,
        "office_collaboration_mutate",
        mutation_arguments(
            &replica,
            "presentation-update-element-mcp",
            serde_json::json!({
                "type": "presentation-update-element",
                "containerKind": "slide",
                "containerId": "slide-1",
                "elementId": "element-title",
                "expectedElement": expected,
                "nextElement": next,
            }),
        ),
        TIMEOUT,
    )
    .await;
    assert_ne!(updated["result"]["isError"], true, "{updated}");
    assert_eq!(updated["result"]["structuredContent"]["sequence"], 1);

    let created_element = call(
        &mut stdin,
        &mut stdout,
        4,
        "office_collaboration_mutate",
        mutation_arguments(
            &replica,
            "presentation-create-element-mcp",
            serde_json::json!({
                "type": "presentation-create-element",
                "containerKind": "layout",
                "containerId": "layout-1",
                "element": presentation_scene_element(
                    "element-layout-mcp",
                    "MCP layout object",
                    "text"
                ),
                "afterElementId": null,
            }),
        ),
        TIMEOUT,
    )
    .await;
    assert_ne!(
        created_element["result"]["isError"], true,
        "{created_element}"
    );
    assert_eq!(
        created_element["result"]["structuredContent"]["sequence"],
        2
    );

    let deleted = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_collaboration_mutate",
        mutation_arguments(
            &replica,
            "presentation-delete-element-mcp",
            serde_json::json!({
                "type": "presentation-delete-element",
                "containerKind": "slide",
                "containerId": "slide-2",
                "expectedElement": presentation_slide_body_element(),
            }),
        ),
        TIMEOUT,
    )
    .await;
    assert_ne!(deleted["result"]["isError"], true, "{deleted}");
    assert_eq!(deleted["result"]["structuredContent"]["sequence"], 3);

    let exported = call(
        &mut stdin,
        &mut stdout,
        6,
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
    let title = presentation_element(&update, "slides", "slide-1", "element-title").unwrap();
    assert_eq!(title["text"], "MCP shared title");
    assert_eq!(title["fill"], "#E0E7FF");
    assert_eq!(
        presentation_element(&update, "layouts", "layout-1", "element-layout-mcp").unwrap()["text"],
        "MCP layout object"
    );
    assert_eq!(
        presentation_element_order(&update, "layouts", "layout-1"),
        vec!["element-layout-mcp".to_owned()]
    );
    assert!(presentation_element_tombstoned(
        &update,
        "slides",
        "slide-2",
        "element-body"
    ));

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
        "actorId": "coding-agent-presentation-mcp",
        "mode": "edit",
        "artifactId": "fixture-presentation",
        "kind": "presentation",
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
