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

fn run_failure(provider: &Path, args: &[&str]) -> serde_json::Value {
    let output = execute(provider, args);
    assert!(!output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

#[test]
fn native_cli_renders_all_html_and_svg_formats_without_officecli() {
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
            "Word <semantic>",
            "--json",
        ],
    );
    run(
        &provider,
        &[
            "set",
            spreadsheet.to_str().unwrap(),
            "/Sheet1/XFD1048576",
            "--text",
            "Spreadsheet & sparse",
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
            "Presentation > preview",
            "--json",
        ],
    );

    for (document, expected) in [
        (&word, "Word &lt;semantic&gt;"),
        (&spreadsheet, "Spreadsheet &amp; sparse"),
        (&presentation, "Presentation &gt; preview"),
    ] {
        let rendered = run(
            &provider,
            &["view", document.to_str().unwrap(), "html", "--json"],
        );
        assert_eq!(rendered["data"]["view"], "html");
        let content = rendered["data"]["result"]["content"].as_str().unwrap();
        assert!(content.starts_with("<!doctype html>"));
        assert!(content.contains(expected));
        assert!(content.contains("Content-Security-Policy"));
        assert!(!content.contains("<script"));
        assert_ne!(rendered["data"]["result"]["sha256"], "");
    }

    for (document, kind, expected) in [
        (&word, "word", "Word &lt;semantic&gt;"),
        (&spreadsheet, "spreadsheet", "Spreadsheet &amp; sparse"),
        (&presentation, "presentation", "Presentation &gt; preview"),
    ] {
        let svg = run(
            &provider,
            &["view", document.to_str().unwrap(), "svg", "--json"],
        );
        assert_eq!(svg["data"]["view"], "svg");
        let content = svg["data"]["result"]["content"].as_str().unwrap();
        assert!(content.starts_with("<?xml version=\"1.0\""));
        assert!(content.contains(&format!("data-document-kind=\"{kind}\"")));
        assert!(content.contains(expected));
        assert!(!content.contains("<script"));
        assert!(!content.contains("href=\"http"));
        assert_eq!(svg["data"]["result"]["mediaType"], "image/svg+xml");
        assert_eq!(svg["data"]["result"]["sha256"].as_str().unwrap().len(), 64);
    }
}

#[test]
fn native_cli_render_files_are_atomic_and_no_clobber() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("report.docx");
    let html = temp.path().join("preview.html");
    let svg = temp.path().join("preview.svg");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);
    run(
        &provider,
        &[
            "set",
            document.to_str().unwrap(),
            "/body/p[1]",
            "--text",
            "Artifact",
            "--json",
        ],
    );

    let written = run(
        &provider,
        &[
            "view",
            document.to_str().unwrap(),
            "html",
            "--output",
            html.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(
        written["data"]["result"]["outputPath"],
        html.to_str().unwrap()
    );
    assert!(written["data"]["result"].get("content").is_none());
    let original = std::fs::read(&html).unwrap();
    assert!(original.starts_with(b"<!doctype html>"));

    let refused = run_failure(
        &provider,
        &[
            "view",
            document.to_str().unwrap(),
            "html",
            "--output",
            html.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(refused["error"]["code"], "use.office.render_output_exists");
    assert_eq!(std::fs::read(&html).unwrap(), original);

    let written = run(
        &provider,
        &[
            "view",
            document.to_str().unwrap(),
            "svg",
            "--output",
            svg.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(
        written["data"]["result"]["outputPath"],
        svg.to_str().unwrap()
    );
    assert!(written["data"]["result"].get("content").is_none());
    assert!(std::fs::read(&svg)
        .unwrap()
        .starts_with(b"<?xml version=\"1.0\""));
}

#[tokio::test]
async fn native_mcp_renders_all_html_and_svg_formats_without_officecli() {
    use std::process::Stdio;
    use std::time::Duration;

    use tokio::io::{AsyncReadExt, AsyncWriteExt, BufReader};

    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let word = temp.path().join("mcp-report.docx");
    let spreadsheet = temp.path().join("mcp-workbook.xlsx");
    let presentation = temp.path().join("mcp-deck.pptx");
    let mut child = tokio::process::Command::new(binary())
        .args(["mcp"])
        .env(
            "A3S_OFFICECLI_EXECUTABLE",
            temp.path().join("must-not-be-invoked"),
        )
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
                "clientInfo": { "name": "render-test", "version": "1" }
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
    let view_schema = tools["result"]["tools"]
        .as_array()
        .unwrap()
        .iter()
        .find(|tool| tool["name"] == "office_view")
        .unwrap();
    let view_schema = view_schema["inputSchema"].to_string();
    assert!(view_schema.contains("html"));
    assert!(view_schema.contains("svg"));
    assert!(view_schema.contains("screenshot"));
    assert!(view_schema.contains("output"));
    assert!(view_schema.contains("timeoutMs"));

    render_mcp_session(
        &mut stdin,
        &mut stdout,
        3,
        "word",
        &word,
        serde_json::json!([{
            "operation":"set-text",
            "path":"/body/p[1]",
            "text":"MCP Word <render>"
        }]),
        "word",
        "MCP Word &lt;render&gt;",
        TIMEOUT,
    )
    .await;
    render_mcp_session(
        &mut stdin,
        &mut stdout,
        8,
        "spreadsheet",
        &spreadsheet,
        serde_json::json!([{
            "operation":"set-cell-value",
            "path":"/Sheet1/XFD1048576",
            "value":{"type":"text","value":"MCP Spreadsheet & render"}
        }]),
        "spreadsheet",
        "MCP Spreadsheet &amp; render",
        TIMEOUT,
    )
    .await;
    render_mcp_session(
        &mut stdin,
        &mut stdout,
        13,
        "presentation",
        &presentation,
        serde_json::json!([{
            "operation":"add-slide",
            "parent":"/",
            "title":"MCP Presentation > render"
        }]),
        "presentation",
        "MCP Presentation &gt; render",
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
}

#[allow(clippy::too_many_arguments)]
async fn render_mcp_session(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    first_id: u32,
    session: &str,
    file: &Path,
    mutations: serde_json::Value,
    kind: &str,
    expected: &str,
    timeout: std::time::Duration,
) {
    let created = call(
        stdin,
        stdout,
        first_id,
        "office_create",
        serde_json::json!({"session":session,"file":file}),
        timeout,
    )
    .await;
    assert_ne!(created["result"]["isError"], true);

    let applied = call(
        stdin,
        stdout,
        first_id + 1,
        "office_apply_batch",
        serde_json::json!({"session":session,"mutations":mutations}),
        timeout,
    )
    .await;
    assert_ne!(applied["result"]["isError"], true, "{applied}");

    for (offset, mode, marker) in [
        (2, "html", "<!doctype html>"),
        (3, "svg", "<?xml version=\"1.0\""),
    ] {
        let rendered = call(
            stdin,
            stdout,
            first_id + offset,
            "office_view",
            serde_json::json!({"session":session,"view":mode}),
            timeout,
        )
        .await;
        assert_ne!(rendered["result"]["isError"], true, "{rendered}");
        let content = rendered["result"]["structuredContent"]["result"]["content"]
            .as_str()
            .unwrap();
        assert!(content.starts_with(marker));
        assert!(content.contains(expected));
        assert!(content.contains(&format!("data-document-kind=\"{kind}\"")));
    }

    let closed = call(
        stdin,
        stdout,
        first_id + 4,
        "office_close",
        serde_json::json!({"session":session,"discard":true}),
        timeout,
    )
    .await;
    assert_eq!(closed["result"]["structuredContent"]["closed"], true);
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
