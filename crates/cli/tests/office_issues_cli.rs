use std::path::Path;
use std::process::{Command, Output};

const PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
];

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

fn failure(provider: &Path, args: &[&str]) -> serde_json::Value {
    let output = execute(provider, args);
    assert!(!output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

#[test]
fn native_cli_reports_bounded_cross_format_issues_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let image = temp.path().join("pixel.png");
    let word = temp.path().join("report.docx");
    let spreadsheet = temp.path().join("workbook.xlsx");
    let presentation = temp.path().join("deck.pptx");
    std::fs::write(&image, PNG_1X1).unwrap();

    for document in [&word, &spreadsheet, &presentation] {
        run(&provider, &["create", document.to_str().unwrap(), "--json"]);
    }
    run(
        &provider,
        &[
            "add",
            word.to_str().unwrap(),
            "/body",
            "--type",
            "picture",
            "--input",
            image.to_str().unwrap(),
            "--name",
            "Chart",
            "--json",
        ],
    );
    for cell in ["/Sheet1/A1", "/Sheet1/A2"] {
        run(
            &provider,
            &[
                "set",
                spreadsheet.to_str().unwrap(),
                cell,
                "--formula",
                "SUM(B1:B2)",
                "--json",
            ],
        );
    }
    run(
        &provider,
        &[
            "add",
            presentation.to_str().unwrap(),
            "/",
            "--type",
            "slide",
            "--text",
            "Issues",
            "--json",
        ],
    );
    run(
        &provider,
        &[
            "add",
            presentation.to_str().unwrap(),
            "/slide[1]",
            "--type",
            "picture",
            "--input",
            image.to_str().unwrap(),
            "--json",
        ],
    );

    for (document, kind) in [(&word, "word"), (&presentation, "presentation")] {
        let report = run(
            &provider,
            &[
                "view",
                document.to_str().unwrap(),
                "issues",
                "--type",
                "missing_alt_text",
                "--json",
            ],
        );
        let result = &report["data"]["result"];
        assert_eq!(report["data"]["view"], "issues");
        assert_eq!(result["kind"], kind);
        assert_eq!(result["filter"], "missing_alt_text");
        assert_eq!(result["count"], 1);
        assert_eq!(result["returned"], 1);
        assert_eq!(result["truncated"], false);
        assert_eq!(result["issues"][0]["type"], "content");
        assert_eq!(result["issues"][0]["subtype"], "missing_alt_text");
        assert!(result["issues"][0]["id"]
            .as_str()
            .unwrap()
            .starts_with("missing_alt_text:/"));
    }

    let formulas = run(
        &provider,
        &[
            "view",
            spreadsheet.to_str().unwrap(),
            "issues",
            "--type",
            "content",
            "--limit",
            "1",
            "--json",
        ],
    );
    let result = &formulas["data"]["result"];
    assert_eq!(result["kind"], "spreadsheet");
    assert_eq!(result["count"], 2);
    assert_eq!(result["returned"], 1);
    assert_eq!(result["truncated"], true);
    assert_eq!(result["issues"][0]["subtype"], "formula_not_evaluated");

    assert!(!provider.exists());
}

#[test]
fn native_cli_issue_options_fail_closed_without_starting_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("report.docx");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);

    let invalid_filter = failure(
        &provider,
        &[
            "view",
            document.to_str().unwrap(),
            "issues",
            "--type",
            "made_up",
            "--json",
        ],
    );
    assert_eq!(
        invalid_filter["error"]["code"],
        "use.office.issue_filter_invalid"
    );

    let invalid_limit = failure(
        &provider,
        &[
            "view",
            document.to_str().unwrap(),
            "issues",
            "--limit",
            "0",
            "--json",
        ],
    );
    assert_eq!(
        invalid_limit["error"]["code"],
        "use.office.issue_limit_invalid"
    );

    let wrong_view = failure(
        &provider,
        &[
            "view",
            document.to_str().unwrap(),
            "text",
            "--limit",
            "1",
            "--json",
        ],
    );
    assert_eq!(wrong_view["error"]["code"], "use.cli.invalid_usage");
    assert!(!provider.exists());
}

#[tokio::test]
async fn native_mcp_exposes_and_runs_typed_issue_views_without_officecli() {
    use std::process::Stdio;
    use std::time::Duration;

    use tokio::io::{AsyncReadExt, AsyncWriteExt, BufReader};

    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("mcp-book.xlsx");
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
                "clientInfo": { "name": "office-issues-test", "version": "1" }
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
    let view = tools["result"]["tools"]
        .as_array()
        .unwrap()
        .iter()
        .find(|tool| tool["name"] == "office_view")
        .unwrap()["inputSchema"]
        .to_string();
    assert!(view.contains("issues"));
    assert!(view.contains("issueType"));
    assert!(view.contains("formula_not_evaluated"));

    let created = call(
        &mut stdin,
        &mut stdout,
        3,
        "office_create",
        serde_json::json!({"session":"book","file":document}),
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
            "session":"book",
            "mutations":[{
                "operation":"set-cell-value",
                "path":"/Sheet1/A1",
                "value":{"type":"formula","expression":"SUM(B1:B2)"}
            }]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(applied["result"]["isError"], true);

    let issues = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_view",
        serde_json::json!({
            "session":"book",
            "view":"issues",
            "issueType":"formula_not_evaluated",
            "limit":10
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(issues["result"]["isError"], true, "{issues}");
    let report = &issues["result"]["structuredContent"]["result"];
    assert_eq!(report["count"], 1);
    assert_eq!(report["issues"][0]["path"], "/Sheet1/A1");
    assert_eq!(report["issues"][0]["subtype"], "formula_not_evaluated");

    let closed = call(
        &mut stdin,
        &mut stdout,
        6,
        "office_close",
        serde_json::json!({"session":"book","discard":true}),
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
