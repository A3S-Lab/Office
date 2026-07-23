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
fn native_cli_manages_typed_comments_in_all_formats_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");

    let word = temp.path().join("comments.docx");
    create(&provider, &word);
    run(
        &provider,
        &[
            "set",
            word.to_str().unwrap(),
            "/body/p[1]",
            "--text",
            "Review this paragraph",
            "--json",
        ],
    );
    let word_comment = run(
        &provider,
        &[
            "add",
            word.to_str().unwrap(),
            "/body/p[1]",
            "--type",
            "comment",
            "--author",
            "Alice",
            "--initials",
            "AL",
            "--text",
            "Please reword this",
            "--json",
        ],
    );
    assert_eq!(word_comment["data"]["operation"], "add-comment");
    assert_eq!(word_comment["data"]["path"], "/comments/comment[1]");
    assert_eq!(word_comment["data"]["node"]["type"], "comment");
    assert_eq!(word_comment["data"]["node"]["format"]["author"], "Alice");
    let word_updated = run(
        &provider,
        &[
            "set",
            word.to_str().unwrap(),
            "/comments/comment[1]",
            "--author",
            "Bob",
            "--text",
            "Updated review",
            "--json",
        ],
    );
    assert_eq!(word_updated["data"]["operation"], "set-comment");
    assert_eq!(word_updated["data"]["node"]["text"], "Updated review");

    let spreadsheet = temp.path().join("comments.xlsx");
    create(&provider, &spreadsheet);
    let cell_comment = run(
        &provider,
        &[
            "add",
            spreadsheet.to_str().unwrap(),
            "/Sheet1/B2",
            "--type",
            "comment",
            "--author",
            "Alice",
            "--text",
            "Check formula",
            "--json",
        ],
    );
    assert_eq!(cell_comment["data"]["path"], "/Sheet1/B2/comment");
    assert_eq!(cell_comment["data"]["node"]["format"]["ref"], "B2");

    let presentation = temp.path().join("comments.pptx");
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
            "Review",
            "--json",
        ],
    );
    let slide_comment = run(
        &provider,
        &[
            "add",
            presentation.to_str().unwrap(),
            "/slide[1]",
            "--type",
            "comment",
            "--author",
            "Alice",
            "--initials",
            "AL",
            "--text",
            "Reword this slide",
            "--x-emu",
            "914400",
            "--y-emu",
            "457200",
            "--json",
        ],
    );
    assert_eq!(slide_comment["data"]["path"], "/slide[1]/comment[1]");
    assert_eq!(slide_comment["data"]["node"]["format"]["xEmu"], "914400");

    for (document, comment) in [
        (&word, "/comments/comment[1]"),
        (&spreadsheet, "/Sheet1/B2/comment"),
        (&presentation, "/slide[1]/comment[1]"),
    ] {
        assert_eq!(
            run(
                &provider,
                &["query", document.to_str().unwrap(), "comment", "--json",],
            )["data"]["matches"],
            1
        );
        run(
            &provider,
            &["remove", document.to_str().unwrap(), comment, "--json"],
        );
    }
    assert!(!provider.exists());
}

#[test]
fn native_cli_rejects_ambiguous_or_format_specific_comment_options_without_writing() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let spreadsheet = temp.path().join("invalid.xlsx");
    create(&provider, &spreadsheet);
    let before = std::fs::read(&spreadsheet).unwrap();

    let half_position = run_failure(
        &provider,
        &[
            "add",
            spreadsheet.to_str().unwrap(),
            "/Sheet1/A1",
            "--type",
            "comment",
            "--author",
            "Alice",
            "--text",
            "Review",
            "--x-emu",
            "1",
            "--json",
        ],
    );
    assert_eq!(half_position["error"]["code"], "use.cli.invalid_usage");
    assert_eq!(std::fs::read(&spreadsheet).unwrap(), before);

    let initials = run_failure(
        &provider,
        &[
            "add",
            spreadsheet.to_str().unwrap(),
            "/Sheet1/A1",
            "--type",
            "comment",
            "--author",
            "Alice",
            "--initials",
            "AL",
            "--text",
            "Review",
            "--json",
        ],
    );
    assert_eq!(
        initials["error"]["code"],
        "use.office.comment_initials_unsupported"
    );
    assert_eq!(std::fs::read(&spreadsheet).unwrap(), before);
    assert!(!provider.exists());
}

#[tokio::test]
async fn native_standard_mcp_manages_unsaved_typed_comments() {
    const TIMEOUT: Duration = Duration::from_secs(15);

    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("mcp-comments.docx");
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
                "clientInfo":{"name":"office-comment-test","version":"1"}
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
    assert!(batch_schema.contains("add-comment"));
    assert!(batch_schema.contains("set-comment"));
    assert!(batch_schema.contains("author"));

    let created = call(
        &mut stdin,
        &mut stdout,
        3,
        "office_create",
        serde_json::json!({"session":"review","file":document}),
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
            "session":"review",
            "mutations":[
                {
                    "operation":"set-text",
                    "path":"/body/p[1]",
                    "text":"Review this paragraph"
                },
                {
                    "operation":"add-comment",
                    "parent":"/body/p[1]",
                    "comment":{
                        "author":"Alice",
                        "initials":"AL",
                        "text":"Please reword this"
                    }
                }
            ]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(applied["result"]["isError"], true, "{applied}");
    assert_eq!(
        applied["result"]["structuredContent"]["result"]["applied"],
        2
    );

    let added = call(
        &mut stdin,
        &mut stdout,
        5,
        "office_get",
        serde_json::json!({
            "session":"review",
            "path":"/comments/comment[1]",
            "depth":0
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(added["result"]["isError"], true, "{added}");
    let node = &added["result"]["structuredContent"]["node"];
    assert_eq!(node["type"], "comment");
    assert_eq!(node["text"], "Please reword this");
    assert_eq!(node["format"]["author"], "Alice");
    assert_eq!(node["format"]["initials"], "AL");
    assert_eq!(node["format"]["anchoredTo"], "/body/p[1]");

    let updated = call(
        &mut stdin,
        &mut stdout,
        6,
        "office_apply_batch",
        serde_json::json!({
            "session":"review",
            "mutations":[{
                "operation":"set-comment",
                "path":"/comments/comment[1]",
                "update":{
                    "author":"Bob",
                    "text":"Updated review"
                }
            }]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(updated["result"]["isError"], true, "{updated}");
    let queried = call(
        &mut stdin,
        &mut stdout,
        7,
        "office_query",
        serde_json::json!({"session":"review","selector":"comment"}),
        TIMEOUT,
    )
    .await;
    assert_ne!(queried["result"]["isError"], true, "{queried}");
    assert_eq!(queried["result"]["structuredContent"]["matches"], 1);
    let updated_node = &queried["result"]["structuredContent"]["results"][0];
    assert_eq!(updated_node["text"], "Updated review");
    assert_eq!(updated_node["format"]["author"], "Bob");
    assert_eq!(updated_node["format"]["initials"], "AL");

    let saved = call(
        &mut stdin,
        &mut stdout,
        8,
        "office_save",
        serde_json::json!({"session":"review"}),
        TIMEOUT,
    )
    .await;
    assert_ne!(saved["result"]["isError"], true, "{saved}");
    let closed = call(
        &mut stdin,
        &mut stdout,
        9,
        "office_close",
        serde_json::json!({"session":"review"}),
        TIMEOUT,
    )
    .await;
    assert_eq!(closed["result"]["structuredContent"]["closed"], true);

    let reopened = call(
        &mut stdin,
        &mut stdout,
        10,
        "office_open",
        serde_json::json!({"session":"cleanup","file":document}),
        TIMEOUT,
    )
    .await;
    assert_ne!(reopened["result"]["isError"], true, "{reopened}");
    let removed = call(
        &mut stdin,
        &mut stdout,
        11,
        "office_apply_batch",
        serde_json::json!({
            "session":"cleanup",
            "mutations":[{
                "operation":"remove",
                "path":"/comments/comment[1]"
            }]
        }),
        TIMEOUT,
    )
    .await;
    assert_ne!(removed["result"]["isError"], true, "{removed}");
    let empty = call(
        &mut stdin,
        &mut stdout,
        12,
        "office_query",
        serde_json::json!({"session":"cleanup","selector":"comment"}),
        TIMEOUT,
    )
    .await;
    assert_ne!(empty["result"]["isError"], true, "{empty}");
    assert_eq!(empty["result"]["structuredContent"]["matches"], 0);
    let saved = call(
        &mut stdin,
        &mut stdout,
        13,
        "office_save",
        serde_json::json!({"session":"cleanup"}),
        TIMEOUT,
    )
    .await;
    assert_ne!(saved["result"]["isError"], true, "{saved}");
    let closed = call(
        &mut stdin,
        &mut stdout,
        14,
        "office_close",
        serde_json::json!({"session":"cleanup"}),
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
