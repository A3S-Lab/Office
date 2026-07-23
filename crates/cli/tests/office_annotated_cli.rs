use std::path::Path;
use std::process::{Command, Output};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

fn execute(provider: &Path, args: &[&str]) -> Output {
    Command::new(binary())
        .args(args)
        .env("A3S_OFFICECLI_EXECUTABLE", provider)
        .output()
        .unwrap()
}

fn run(provider: &Path, args: &[&str]) -> serde_json::Value {
    let output = execute(provider, args);
    assert!(output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

#[test]
fn native_annotated_cli_covers_all_formats_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let word = temp.path().join("report.docx");
    let spreadsheet = temp.path().join("workbook.xlsx");
    let presentation = temp.path().join("deck.pptx");

    for document in [&word, &spreadsheet, &presentation] {
        run(&provider, &["create", document.to_str().unwrap(), "--json"]);
    }
    run(
        &provider,
        &[
            "set",
            word.to_str().unwrap(),
            "/body/p[1]",
            "--text",
            "Annotated Word",
            "--json",
        ],
    );
    run(
        &provider,
        &[
            "set",
            spreadsheet.to_str().unwrap(),
            "/Sheet1/A1",
            "--formula",
            "SUM(B1:B2)",
            "--json",
        ],
    );
    run(
        &provider,
        &[
            "add",
            presentation.to_str().unwrap(),
            "/",
            "--type",
            "slide",
            "--text",
            "Annotated Deck",
            "--json",
        ],
    );

    for (document, kind, expected_path) in [
        (&word, "word", "/body/p[1]"),
        (&spreadsheet, "spreadsheet", "/Sheet1/A1"),
        (&presentation, "presentation", "/slide[1]"),
    ] {
        let view = run(
            &provider,
            &[
                "view",
                document.to_str().unwrap(),
                "annotated",
                "--limit",
                "100",
                "--json",
            ],
        );
        assert_eq!(view["data"]["view"], "annotated");
        assert_eq!(view["data"]["result"]["kind"], kind);
        assert!(view["data"]["result"]["entries"]
            .as_array()
            .unwrap()
            .iter()
            .any(|entry| entry["path"] == expected_path));
        assert_eq!(view["data"]["result"]["limit"], 100);
    }
}

#[test]
fn native_annotated_cli_is_bounded_and_rejects_unrelated_options() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("bounded.docx");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);

    for args in [
        vec!["--limit", "0"],
        vec!["--limit", "1001"],
        vec!["--output", "annotated.txt"],
        vec!["--type", "content"],
        vec!["--timeout-ms", "1000"],
    ] {
        let mut command = vec!["view", document.to_str().unwrap(), "annotated"];
        command.extend(args);
        command.push("--json");
        let output = execute(&provider, &command);
        assert!(!output.status.success(), "{output:?}");
        let error: serde_json::Value = serde_json::from_slice(&output.stdout).unwrap();
        assert!(matches!(
            error["error"]["code"].as_str(),
            Some("use.office.annotated_limit_invalid" | "use.cli.invalid_usage")
        ));
    }
}

#[tokio::test]
async fn native_mcp_annotated_view_reads_unsaved_typed_session_state() {
    use std::process::Stdio;
    use std::time::Duration;

    use tokio::io::{AsyncReadExt, AsyncWriteExt, BufReader};

    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("session.docx");
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
                "clientInfo": { "name": "office-annotated-test", "version": "1" }
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
        .find(|tool| tool["name"] == "office_view")
        .unwrap()["inputSchema"]
        .to_string();
    assert!(schema.contains("annotated"));

    let created = call(
        &mut stdin,
        &mut stdout,
        3,
        "office_create",
        serde_json::json!({"session":"report","file":document}),
        TIMEOUT,
    )
    .await;
    assert_ne!(created["result"]["isError"], true);

    let applied = call(
        &mut stdin,
        &mut stdout,
        4,
        "office_apply_batch",
        serde_json::json!({
            "session":"report",
            "mutations":[{
                "operation":"set-text",
                "path":"/body/p[1]",
                "text":"Unsaved annotated state"
            }]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(applied["result"]["isError"], true, "{applied}");
    let disk = a3s_office::NativeOfficeDocument::open(&document)
        .await
        .unwrap();
    assert_eq!(disk.text_view().text, "");

    let annotated = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_view",
        serde_json::json!({"session":"report","view":"annotated","limit":10}),
        TIMEOUT,
    )
    .await;
    assert_ne!(annotated["result"]["isError"], true, "{annotated}");
    let result = &annotated["result"]["structuredContent"]["result"];
    assert_eq!(result["kind"], "word");
    assert_eq!(result["limit"], 10);
    assert!(result["entries"]
        .as_array()
        .unwrap()
        .iter()
        .any(|entry| entry["path"] == "/body/p[1]" && entry["text"] == "Unsaved annotated state"));

    let invalid = call(
        &mut stdin,
        &mut stdout,
        6,
        "office_view",
        serde_json::json!({"session":"report","view":"annotated","limit":1001}),
        TIMEOUT,
    )
    .await;
    assert_eq!(invalid["result"]["isError"], true);
    assert_eq!(
        invalid["result"]["structuredContent"]["code"],
        "use.office.annotated_limit_invalid"
    );

    let closed = call(
        &mut stdin,
        &mut stdout,
        7,
        "office_close",
        serde_json::json!({"session":"report","discard":true}),
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
    assert!(!provider.exists());
}

async fn call(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    id: u32,
    name: &str,
    arguments: serde_json::Value,
    timeout: std::time::Duration,
) -> serde_json::Value {
    request(
        stdin,
        stdout,
        serde_json::json!({
            "jsonrpc":"2.0",
            "id":id,
            "method":"tools/call",
            "params":{"name":name,"arguments":arguments}
        }),
        timeout,
    )
    .await
}

async fn request(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    value: serde_json::Value,
    timeout: std::time::Duration,
) -> serde_json::Value {
    use tokio::io::{AsyncBufReadExt, AsyncWriteExt};

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
