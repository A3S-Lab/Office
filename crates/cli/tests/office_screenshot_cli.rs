use std::path::Path;
use std::process::{Command, Output};

#[cfg(feature = "browser")]
use sha2::{Digest, Sha256};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

fn command(provider: &Path) -> Command {
    let mut command = Command::new(binary());
    command.env("A3S_OFFICECLI_EXECUTABLE", provider);
    command
}

fn execute(provider: &Path, args: &[&str]) -> Output {
    command(provider).args(args).output().unwrap()
}

#[cfg(feature = "browser")]
fn execute_with_browser(provider: &Path, browser: &Path, args: &[&str]) -> Output {
    command(provider)
        .env("A3S_BROWSER_EXECUTABLE", browser)
        .args(args)
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

#[cfg(feature = "browser")]
#[test]
fn native_cli_screenshot_validates_options_before_starting_browser() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("report.docx");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);

    let missing = failure(
        &provider,
        &["view", document.to_str().unwrap(), "screenshot", "--json"],
    );
    assert_eq!(missing["error"]["code"], "use.cli.invalid_usage");

    let bad_extension = temp.path().join("preview.jpg");
    let invalid_output = failure(
        &provider,
        &[
            "view",
            document.to_str().unwrap(),
            "screenshot",
            "--output",
            bad_extension.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(
        invalid_output["error"]["code"],
        "use.office.screenshot_output_invalid"
    );
    assert!(!bad_extension.exists());

    let zero_timeout = temp.path().join("zero.png");
    let invalid_timeout = failure(
        &provider,
        &[
            "view",
            document.to_str().unwrap(),
            "screenshot",
            "--output",
            zero_timeout.to_str().unwrap(),
            "--timeout-ms",
            "0",
            "--json",
        ],
    );
    assert_eq!(
        invalid_timeout["error"]["code"],
        "use.office.screenshot_timeout_invalid"
    );
    assert!(!zero_timeout.exists());

    let excessive_timeout = temp.path().join("excessive.png");
    let invalid_timeout = failure(
        &provider,
        &[
            "view",
            document.to_str().unwrap(),
            "png",
            "--output",
            excessive_timeout.to_str().unwrap(),
            "--timeout-ms",
            "120001",
            "--json",
        ],
    );
    assert_eq!(
        invalid_timeout["error"]["code"],
        "use.office.screenshot_timeout_invalid"
    );
    assert!(!excessive_timeout.exists());
}

#[cfg(feature = "browser")]
#[test]
fn native_cli_screenshot_never_clobbers_an_existing_output() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("report.docx");
    let output = temp.path().join("preview.png");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);
    std::fs::write(&output, b"preserve this output").unwrap();

    let refused = failure(
        &provider,
        &[
            "view",
            document.to_str().unwrap(),
            "screenshot",
            "--output",
            output.to_str().unwrap(),
            "--json",
        ],
    );

    assert_eq!(
        refused["error"]["code"],
        "use.office.screenshot_output_exists"
    );
    assert_eq!(std::fs::read(&output).unwrap(), b"preserve this output");
}

#[cfg(not(feature = "browser"))]
#[test]
fn native_cli_screenshot_has_a_stable_browser_disabled_error() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("report.docx");
    let output = temp.path().join("preview.png");
    run(&provider, &["create", document.to_str().unwrap(), "--json"]);

    let disabled = failure(
        &provider,
        &[
            "view",
            document.to_str().unwrap(),
            "screenshot",
            "--output",
            output.to_str().unwrap(),
            "--json",
        ],
    );

    assert_eq!(disabled["error"]["code"], "use.browser.disabled");
    assert!(!output.exists());
}

#[cfg(feature = "browser")]
#[cfg_attr(
    windows,
    ignore = "Windows native Office screenshots are roadmap; macOS and Linux are the current supported platforms"
)]
#[test]
fn native_cli_screenshots_all_office_formats_with_discovered_chrome() {
    let Some(browser) = a3s_use_browser::detect_chrome() else {
        return;
    };
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");

    for (extension, kind) in [
        ("docx", "word"),
        ("xlsx", "spreadsheet"),
        ("pptx", "presentation"),
    ] {
        let document = temp.path().join(format!("source.{extension}"));
        let output = temp.path().join(format!("preview-{kind}.png"));
        let created = execute_with_browser(
            &provider,
            &browser,
            &["create", document.to_str().unwrap(), "--json"],
        );
        assert!(created.status.success(), "{created:?}");

        let rendered = execute_with_browser(
            &provider,
            &browser,
            &[
                "view",
                document.to_str().unwrap(),
                "screenshot",
                "--output",
                output.to_str().unwrap(),
                "--timeout-ms",
                "20000",
                "--json",
            ],
        );
        assert!(rendered.status.success(), "{rendered:?}");
        let receipt: serde_json::Value = serde_json::from_slice(&rendered.stdout).unwrap();
        let result = &receipt["data"]["result"];
        let bytes = std::fs::read(&output).unwrap();
        let sha256 = format!("{:x}", Sha256::digest(&bytes));

        assert_eq!(receipt["data"]["view"], "screenshot");
        assert_eq!(result["kind"], kind);
        assert_eq!(result["outputPath"], output.to_str().unwrap());
        assert_eq!(result["mediaType"], "image/png");
        assert!(result["widthPx"].as_u64().unwrap() > 0);
        assert!(result["heightPx"].as_u64().unwrap() > 0);
        assert_eq!(result["byteLength"], bytes.len() as u64);
        assert_eq!(result["sha256"], sha256);
        assert_eq!(result["sourceHtmlSha256"].as_str().unwrap().len(), 64);
        assert!(bytes.starts_with(b"\x89PNG\r\n\x1a\n"));
    }

    assert!(!provider.exists());
}

#[cfg(feature = "browser")]
#[cfg_attr(
    windows,
    ignore = "Windows native Office screenshots are roadmap; macOS and Linux are the current supported platforms"
)]
#[tokio::test(flavor = "multi_thread", worker_threads = 2)]
async fn native_mcp_screenshot_uses_the_browser_contract_when_chrome_is_available() {
    use std::process::Stdio;
    use std::time::Duration;

    use tokio::io::{AsyncReadExt, AsyncWriteExt, BufReader};

    const TIMEOUT: Duration = Duration::from_secs(30);

    let Some(browser) = a3s_use_browser::detect_chrome() else {
        return;
    };
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("mcp-report.docx");
    let output = temp.path().join("mcp-preview.png");
    let mut child = tokio::process::Command::new(binary())
        .args(["mcp"])
        .env("A3S_OFFICECLI_EXECUTABLE", &provider)
        .env("A3S_BROWSER_EXECUTABLE", browser)
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
                "clientInfo": { "name": "office-screenshot-test", "version": "1" }
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
        "office_create",
        serde_json::json!({"session":"report","file":document}),
        TIMEOUT,
    )
    .await;
    assert_ne!(created["result"]["isError"], true);

    let rendered = call(
        &mut stdin,
        &mut stdout,
        3,
        "office_view",
        serde_json::json!({
            "session":"report",
            "view":"screenshot",
            "output":output,
            "timeoutMs":20000
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(rendered["result"]["isError"], true, "{rendered}");
    let result = &rendered["result"]["structuredContent"]["result"];
    let bytes = std::fs::read(&output).unwrap();
    assert_eq!(result["kind"], "word");
    assert_eq!(result["mediaType"], "image/png");
    assert_eq!(result["byteLength"], bytes.len() as u64);
    assert_eq!(result["sha256"], format!("{:x}", Sha256::digest(&bytes)));
    assert!(bytes.starts_with(b"\x89PNG\r\n\x1a\n"));

    let closed = call(
        &mut stdin,
        &mut stdout,
        4,
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

#[cfg(feature = "browser")]
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

#[cfg(feature = "browser")]
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
