use std::time::Duration;

use a3s_office::NativeOfficeEditor;
use a3s_office_cli::office_watch::{
    NativeOfficeWatchOptions, NativeOfficeWatchServer, NativeOfficeWatchStatus,
    MIN_NATIVE_OFFICE_WATCH_POLL_MS,
};
use futures_util::StreamExt;

const WAIT: Duration = Duration::from_secs(10);

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

#[tokio::test]
async fn native_watch_binds_every_ooxml_format_without_an_external_runtime() {
    let temp = tempfile::tempdir().unwrap();
    for (name, kind) in [
        ("document.docx", a3s_office::DocumentKind::Word),
        ("workbook.xlsx", a3s_office::DocumentKind::Spreadsheet),
        ("presentation.pptx", a3s_office::DocumentKind::Presentation),
    ] {
        let path = temp.path().join(name);
        NativeOfficeEditor::create(&path).await.unwrap();
        let server = NativeOfficeWatchServer::bind(&path, NativeOfficeWatchOptions::default())
            .await
            .unwrap();
        assert_eq!(server.ready().kind, kind);
        assert_eq!(server.ready().address.ip().to_string(), "127.0.0.1");
        assert_eq!(server.ready().url.matches("token=").count(), 1);
    }
}

#[tokio::test]
async fn native_watch_is_authenticated_read_only_and_recovers_across_revisions() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("secret-watch-source.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_text("/body/p[1]", "Initial <preview>").unwrap();
    editor.save().await.unwrap();
    let original = tokio::fs::read(&path).await.unwrap();

    let server = NativeOfficeWatchServer::bind(
        &path,
        NativeOfficeWatchOptions {
            port: 0,
            poll_interval_ms: MIN_NATIVE_OFFICE_WATCH_POLL_MS,
        },
    )
    .await
    .unwrap();
    let ready = server.ready().clone();
    assert_eq!(ready.version, 1);
    assert_eq!(ready.kind, a3s_office::DocumentKind::Word);
    assert_eq!(ready.url.matches("token=").count(), 1);
    assert_eq!(ready.render_sha256.len(), 64);

    let (stop_tx, stop_rx) = tokio::sync::oneshot::channel();
    let task = tokio::spawn(server.serve(async move {
        let _ = stop_rx.await;
    }));
    let client = reqwest::Client::new();
    let origin = format!("http://{}", ready.address);
    let token = ready.url.split("token=").nth(1).unwrap();

    let unauthorized = client.get(&origin).send().await.unwrap();
    assert_eq!(unauthorized.status(), reqwest::StatusCode::NOT_FOUND);
    let wrong_host = client
        .get(&ready.url)
        .header("host", "example.invalid")
        .send()
        .await
        .unwrap();
    assert_eq!(wrong_host.status(), reqwest::StatusCode::NOT_FOUND);

    let page = client.get(&ready.url).send().await.unwrap();
    assert_eq!(page.status(), reqwest::StatusCode::OK);
    assert!(page.headers().contains_key("content-security-policy"));
    assert_eq!(
        page.headers()["cross-origin-resource-policy"],
        "same-origin"
    );
    assert!(page.headers().contains_key("set-cookie"));
    let page = page.text().await.unwrap();
    assert!(page.contains("A3S Native Office semantic preview"));
    assert!(!page.contains("secret-watch-source.docx"));

    let preview_url = format!("{origin}/preview?token={token}");
    let preview = client.get(&preview_url).send().await.unwrap();
    assert_eq!(preview.status(), reqwest::StatusCode::OK);
    assert_eq!(
        preview.headers()["cache-control"],
        reqwest::header::HeaderValue::from_static("no-store")
    );
    let preview = preview.text().await.unwrap();
    assert!(preview.contains("Initial &lt;preview&gt;"));
    assert!(!preview.contains("secret-watch-source.docx"));

    let events_url = format!("{origin}/events?token={token}");
    let events = client.get(&events_url).send().await.unwrap();
    assert_eq!(events.status(), reqwest::StatusCode::OK);
    assert!(events.headers()["content-type"]
        .to_str()
        .unwrap()
        .starts_with("text/event-stream"));
    let mut events = events.bytes_stream();
    let initial_event = next_event(&mut events, "event: snapshot").await;
    assert!(initial_event.contains("\"version\":1"));

    tokio::fs::write(&path, b"not an OOXML package")
        .await
        .unwrap();
    let failed = wait_for_status(&client, &origin, token, |status| !status.healthy).await;
    assert_eq!(failed.version, 1);
    assert!(failed.error.is_some());
    let retained = client
        .get(&preview_url)
        .send()
        .await
        .unwrap()
        .text()
        .await
        .unwrap();
    assert!(retained.contains("Initial &lt;preview&gt;"));

    tokio::fs::write(&path, &original).await.unwrap();
    let recovered = wait_for_status(&client, &origin, token, |status| status.healthy).await;
    assert_eq!(recovered.version, 1);

    let mut editor = NativeOfficeEditor::open(&path).await.unwrap();
    editor.set_text("/body/p[1]", "Updated & live").unwrap();
    editor.save().await.unwrap();
    let updated = wait_for_status(&client, &origin, token, |status| status.version == 2).await;
    assert!(updated.healthy);
    assert_ne!(updated.revision, ready.revision);
    assert_ne!(updated.render_sha256, ready.render_sha256);
    let update_event = next_event(&mut events, "\"version\":2").await;
    assert!(update_event.contains("event: snapshot"));
    let preview = client
        .get(&preview_url)
        .send()
        .await
        .unwrap()
        .text()
        .await
        .unwrap();
    assert!(preview.contains("Updated &amp; live"));

    stop_tx.send(()).unwrap();
    tokio::time::timeout(WAIT, task)
        .await
        .expect("watch server did not stop")
        .unwrap()
        .unwrap();
}

#[tokio::test]
async fn native_watch_cli_refreshes_after_a_separate_native_command_without_officecli() {
    use std::process::Stdio;
    use tokio::io::AsyncBufReadExt;

    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("cli-watch.xlsx");
    let provider = temp.path().join("must-not-be-invoked");
    NativeOfficeEditor::create(&path).await.unwrap();

    let mut child = tokio::process::Command::new(binary())
        .args([
            "watch",
            path.to_str().unwrap(),
            "--port",
            "0",
            "--poll-ms",
            "50",
            "--timeout-ms",
            "10000",
            "--json",
        ])
        .env("A3S_OFFICECLI_EXECUTABLE", &provider)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    let stdout = child.stdout.take().unwrap();
    let mut stdout = tokio::io::BufReader::new(stdout);
    let mut line = String::new();
    tokio::time::timeout(WAIT, stdout.read_line(&mut line))
        .await
        .expect("watch startup timed out")
        .unwrap();
    let receipt: serde_json::Value = serde_json::from_str(&line).unwrap();
    assert_eq!(receipt["ok"], true);
    assert_eq!(receipt["data"]["operation"], "watch");
    assert_eq!(receipt["data"]["ready"], true);
    assert_eq!(receipt["data"]["server"]["kind"], "spreadsheet");
    let watch_url = receipt["data"]["server"]["url"].as_str().unwrap();
    let address = receipt["data"]["server"]["address"].as_str().unwrap();
    let token = watch_url.split("token=").nth(1).unwrap();
    let origin = format!("http://{address}");

    let mutation = tokio::process::Command::new(binary())
        .args([
            "set",
            path.to_str().unwrap(),
            "/Sheet1/XFD1048576",
            "--text",
            "CLI live <update>",
            "--json",
        ])
        .env("A3S_OFFICECLI_EXECUTABLE", &provider)
        .output()
        .await
        .unwrap();
    assert!(mutation.status.success(), "{mutation:?}");

    let client = reqwest::Client::new();
    let status = wait_for_status(&client, &origin, token, |status| status.version == 2).await;
    assert!(status.healthy);
    let preview = client
        .get(format!("{origin}/preview?token={token}"))
        .send()
        .await
        .unwrap()
        .text()
        .await
        .unwrap();
    assert!(preview.contains("CLI live &lt;update&gt;"));
    assert!(!provider.exists());

    child.start_kill().unwrap();
    tokio::time::timeout(WAIT, child.wait())
        .await
        .expect("watch process did not stop")
        .unwrap();
}

async fn wait_for_status(
    client: &reqwest::Client,
    origin: &str,
    token: &str,
    predicate: impl Fn(&NativeOfficeWatchStatus) -> bool,
) -> NativeOfficeWatchStatus {
    let deadline = tokio::time::Instant::now() + WAIT;
    loop {
        let value = client
            .get(format!("{origin}/status?token={token}"))
            .send()
            .await
            .unwrap()
            .text()
            .await
            .unwrap();
        let status: NativeOfficeWatchStatus = serde_json::from_str(&value).unwrap();
        if predicate(&status) {
            return status;
        }
        assert!(
            tokio::time::Instant::now() < deadline,
            "status timed out: {status:?}"
        );
        tokio::time::sleep(Duration::from_millis(25)).await;
    }
}

async fn next_event<S>(stream: &mut S, marker: &str) -> String
where
    S: futures_util::Stream<Item = Result<axum::body::Bytes, reqwest::Error>> + Unpin,
{
    let mut output = String::new();
    tokio::time::timeout(WAIT, async {
        while let Some(chunk) = stream.next().await {
            output.push_str(std::str::from_utf8(&chunk.unwrap()).unwrap());
            if output.contains(marker) {
                break;
            }
        }
    })
    .await
    .expect("SSE event timed out");
    output
}
