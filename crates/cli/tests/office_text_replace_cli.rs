use std::path::Path;
use std::process::{Command, Output};

use std::process::Stdio;
use std::time::Duration;
use tokio::io::{AsyncReadExt, AsyncWriteExt, BufReader};

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

fn create(provider: &Path, document: &Path) {
    run(provider, &["create", document.to_str().unwrap(), "--json"]);
}

#[test]
fn native_cli_replaces_bounded_text_in_all_formats_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");

    let word = temp.path().join("replace.docx");
    create(&provider, &word);
    run(
        &provider,
        &[
            "set",
            word.to_str().unwrap(),
            "/body/p[1]",
            "--text",
            "Q1 2025 and Q2 2025",
            "--json",
        ],
    );
    let replaced = run(
        &provider,
        &[
            "set",
            word.to_str().unwrap(),
            "/body",
            "--find",
            r"Q([1-4]) 2025",
            "--replace",
            "Q$1 2026",
            "--regex",
            "--json",
        ],
    );
    assert_eq!(replaced["data"]["operation"], "replace-text");
    assert_eq!(replaced["data"]["matches"], 2);
    assert_eq!(replaced["data"]["changed"], true);
    assert_eq!(replaced["data"]["result"]["mode"], "regex");
    assert_eq!(
        replaced["data"]["result"]["changedParts"],
        serde_json::json!(["/word/document.xml"])
    );
    assert_eq!(
        run(
            &provider,
            &["get", word.to_str().unwrap(), "/body/p[1]", "--json",],
        )["data"]["node"]["text"],
        "Q1 2026 and Q2 2026"
    );
    let before_zero = std::fs::read(&word).unwrap();
    #[cfg(unix)]
    let inode_before_zero = {
        use std::os::unix::fs::MetadataExt;
        std::fs::metadata(&word).unwrap().ino()
    };
    let zero = run(
        &provider,
        &[
            "set",
            word.to_str().unwrap(),
            "/body",
            "--find",
            "missing",
            "--replace",
            "unused",
            "--json",
        ],
    );
    assert_eq!(zero["data"]["matches"], 0);
    assert_eq!(zero["data"]["changed"], false);
    assert_eq!(std::fs::read(&word).unwrap(), before_zero);
    #[cfg(unix)]
    {
        use std::os::unix::fs::MetadataExt;
        assert_eq!(std::fs::metadata(&word).unwrap().ino(), inode_before_zero);
    }

    let spreadsheet = temp.path().join("replace.xlsx");
    create(&provider, &spreadsheet);
    for (cell, text) in [("/Sheet1/A1", "alpha beta"), ("/Sheet1/A2", "alpha beta")] {
        run(
            &provider,
            &[
                "set",
                spreadsheet.to_str().unwrap(),
                cell,
                "--text",
                text,
                "--json",
            ],
        );
    }
    let cell = run(
        &provider,
        &[
            "set",
            spreadsheet.to_str().unwrap(),
            "/Sheet1/A1",
            "--find",
            "alpha beta",
            "--replace",
            "selected",
            "--json",
        ],
    );
    assert_eq!(cell["data"]["matches"], 1);
    for (cell, expected) in [("/Sheet1/A1", "selected"), ("/Sheet1/A2", "alpha beta")] {
        assert_eq!(
            run(
                &provider,
                &["get", spreadsheet.to_str().unwrap(), cell, "--json",],
            )["data"]["node"]["text"],
            expected
        );
    }

    let presentation = temp.path().join("replace.pptx");
    create(&provider, &presentation);
    run(
        &provider,
        &[
            "add",
            presentation.to_str().unwrap(),
            "/",
            "--type",
            "slide",
            "--text",
            "Alpha roadmap",
            "--json",
        ],
    );
    let slide = run(
        &provider,
        &[
            "set",
            presentation.to_str().unwrap(),
            "/slide[1]",
            "--find",
            "Alpha",
            "--replace",
            "A3S",
            "--json",
        ],
    );
    assert_eq!(slide["data"]["matches"], 1);
    assert_eq!(
        run(
            &provider,
            &[
                "get",
                presentation.to_str().unwrap(),
                "/slide[1]/shape[1]",
                "--json",
            ],
        )["data"]["node"]["text"],
        "A3S roadmap"
    );
    assert!(!provider.exists());
}

#[test]
fn native_cli_rejects_ambiguous_replacement_options_without_writing() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let word = temp.path().join("invalid.docx");
    create(&provider, &word);
    let before = std::fs::read(&word).unwrap();
    let error = run_failure(
        &provider,
        &[
            "set",
            word.to_str().unwrap(),
            "/body",
            "--find",
            "alpha",
            "--replace",
            "beta",
            "--text",
            "ambiguous",
            "--json",
        ],
    );
    assert_eq!(error["error"]["code"], "use.cli.invalid_usage");
    assert_eq!(std::fs::read(&word).unwrap(), before);

    let missing = run_failure(
        &provider,
        &[
            "set",
            word.to_str().unwrap(),
            "/body",
            "--find",
            "alpha",
            "--json",
        ],
    );
    assert_eq!(missing["error"]["code"], "use.cli.invalid_usage");
    assert_eq!(std::fs::read(&word).unwrap(), before);
    assert!(!provider.exists());
}

#[tokio::test]
async fn native_standard_mcp_applies_and_persists_typed_text_replacement() {
    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("mcp-replace.docx");
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
            "jsonrpc":"2.0",
            "id":1,
            "method":"initialize",
            "params":{
                "protocolVersion":"2025-06-18",
                "capabilities":{},
                "clientInfo":{"name":"office-replace-test","version":"1"}
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
    let batch_schema = tools["result"]["tools"]
        .as_array()
        .unwrap()
        .iter()
        .find(|tool| tool["name"] == "office_apply_batch")
        .unwrap()["inputSchema"]
        .to_string();
    for expected in ["replace-text", "find", "replace", "literal", "regex"] {
        assert!(batch_schema.contains(expected), "missing {expected}");
    }

    let created = call(
        &mut stdin,
        &mut stdout,
        3,
        "office_create",
        serde_json::json!({"session":"replace","file":document}),
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
            "session":"replace",
            "mutations":[
                {
                    "operation":"set-text",
                    "path":"/body/p[1]",
                    "text":"Q1 2025 and Q2 2025"
                },
                {
                    "operation":"replace-text",
                    "path":"/body",
                    "replacement":{
                        "find":"Q([1-4]) 2025",
                        "replace":"Q$1 2026",
                        "mode":"regex"
                    }
                }
            ]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(applied["result"]["isError"], true, "{applied}");
    let result = &applied["result"]["structuredContent"]["result"];
    assert_eq!(result["applied"], 2);
    assert_eq!(result["textReplacements"][0]["matchCount"], 2);
    assert_eq!(result["textReplacements"][0]["changed"], true);

    let unsaved = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_get",
        serde_json::json!({
            "session":"replace",
            "path":"/body/p[1]",
            "depth":0
        }),
        TIMEOUT,
    )
    .await;
    assert_eq!(
        unsaved["result"]["structuredContent"]["node"]["text"],
        "Q1 2026 and Q2 2026"
    );
    let saved = call(
        &mut stdin,
        &mut stdout,
        6,
        "office_save",
        serde_json::json!({"session":"replace"}),
        TIMEOUT,
    )
    .await;
    assert_ne!(saved["result"]["isError"], true, "{saved}");
    let closed = call(
        &mut stdin,
        &mut stdout,
        7,
        "office_close",
        serde_json::json!({"session":"replace"}),
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
    assert_eq!(
        run(
            &provider,
            &["get", document.to_str().unwrap(), "/body/p[1]", "--json",],
        )["data"]["node"]["text"],
        "Q1 2026 and Q2 2026"
    );
    assert!(!provider.exists());
}

async fn call(
    stdin: &mut tokio::process::ChildStdin,
    stdout: &mut tokio::io::BufReader<tokio::process::ChildStdout>,
    id: u32,
    name: &str,
    arguments: serde_json::Value,
    timeout: Duration,
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
    timeout: Duration,
) -> serde_json::Value {
    use tokio::io::AsyncBufReadExt;

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
