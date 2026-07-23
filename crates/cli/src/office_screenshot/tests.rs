use std::sync::atomic::{AtomicUsize, Ordering};

use a3s_office::NativeOfficeEditor;
use a3s_use_browser::RenderedPage;
use a3s_use_core::Artifact;
use async_trait::async_trait;

use super::*;

const PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
];

struct FixtureRenderer {
    calls: Arc<AtomicUsize>,
    corrupt: bool,
}

#[async_trait]
impl PageRenderer for FixtureRenderer {
    async fn render(&self, request: RenderRequest) -> UseResult<RenderedPage> {
        self.calls.fetch_add(1, Ordering::SeqCst);
        assert_eq!(request.url.scheme(), "file");
        assert_eq!(request.wait, WaitCondition::Load);
        let html_path = request.url.to_file_path().unwrap();
        let html = tokio::fs::read_to_string(html_path).await.unwrap();
        assert!(html.contains("Fixture &lt;Office&gt;"));
        assert!(html.contains("Content-Security-Policy"));
        let screenshot_path = request.screenshot_path.unwrap();
        let bytes = if self.corrupt {
            b"not a png".as_slice()
        } else {
            PNG_1X1
        };
        tokio::fs::write(&screenshot_path, bytes).await.unwrap();
        let sha256 = format!("{:x}", Sha256::digest(bytes));
        Ok(RenderedPage {
            requested_url: request.url.clone(),
            final_url: request.url,
            status: None,
            content_type: Some("text/html".to_string()),
            html,
            elapsed_ms: 7,
            artifacts: vec![Artifact {
                path: screenshot_path,
                media_type: "image/png".to_string(),
                size: bytes.len() as u64,
                sha256,
            }],
        })
    }
}

#[tokio::test]
async fn screenshot_renderer_is_injectable_validated_and_no_clobber() {
    let temp = tempfile::tempdir().unwrap();
    let document_path = temp.path().join("report.docx");
    let output = temp.path().join("report.png");
    let mut editor = NativeOfficeEditor::create(&document_path).await.unwrap();
    editor.set_text("/body/p[1]", "Fixture <Office>").unwrap();
    let document = editor.snapshot().unwrap();
    let calls = Arc::new(AtomicUsize::new(0));
    let renderer = NativeOfficeScreenshotRenderer::new(Arc::new(FixtureRenderer {
        calls: Arc::clone(&calls),
        corrupt: false,
    }));

    let screenshot = renderer
        .render(
            &document,
            NativeOfficeScreenshotRequest::new(&output).with_timeout_ms(2_000),
        )
        .await
        .unwrap();

    assert_eq!(screenshot.kind, DocumentKind::Word);
    assert_eq!(screenshot.output_path, output);
    assert_eq!(screenshot.media_type, "image/png");
    assert_eq!(screenshot.width_px, 1);
    assert_eq!(screenshot.height_px, 1);
    assert_eq!(screenshot.byte_length, PNG_1X1.len() as u64);
    assert_eq!(screenshot.renderer_elapsed_ms, 7);
    assert_eq!(std::fs::read(&output).unwrap(), PNG_1X1);
    assert_eq!(calls.load(Ordering::SeqCst), 1);

    let error = renderer
        .render(&document, NativeOfficeScreenshotRequest::new(&output))
        .await
        .unwrap_err();
    assert_eq!(error.code, "use.office.screenshot_output_exists");
    assert_eq!(calls.load(Ordering::SeqCst), 1);
    assert_eq!(std::fs::read(&output).unwrap(), PNG_1X1);
}

#[tokio::test]
async fn screenshot_renderer_rejects_invalid_provider_artifacts() {
    let temp = tempfile::tempdir().unwrap();
    let document_path = temp.path().join("report.docx");
    let output = temp.path().join("report.png");
    let mut editor = NativeOfficeEditor::create(&document_path).await.unwrap();
    editor.set_text("/body/p[1]", "Fixture <Office>").unwrap();
    let renderer = NativeOfficeScreenshotRenderer::new(Arc::new(FixtureRenderer {
        calls: Arc::new(AtomicUsize::new(0)),
        corrupt: true,
    }));

    let error = renderer
        .render(
            &editor.snapshot().unwrap(),
            NativeOfficeScreenshotRequest::new(&output),
        )
        .await
        .unwrap_err();

    assert_eq!(error.code, "use.office.screenshot_artifact_invalid");
    assert!(!output.exists());
}

#[test]
fn public_screenshot_contracts_are_send_and_sync() {
    fn assert_send_sync<T: Send + Sync>() {}

    assert_send_sync::<NativeOfficeScreenshotRequest>();
    assert_send_sync::<NativeOfficeScreenshot>();
    assert_send_sync::<NativeOfficeScreenshotRenderer>();
}

#[cfg(not(windows))]
#[tokio::test(flavor = "multi_thread", worker_threads = 2)]
async fn discovered_chrome_captures_office_semantic_html_when_available() {
    use a3s_use_browser::{BrowserPool, BrowserPoolConfig, BrowserProvider};

    let Some(executable) = a3s_use_browser::detect_chrome() else {
        return;
    };
    let temp = tempfile::tempdir().unwrap();
    let document_path = temp.path().join("report.docx");
    let output = temp.path().join("report.png");
    let mut editor = NativeOfficeEditor::create(&document_path).await.unwrap();
    editor.set_text("/body/p[1]", "Fixture <Office>").unwrap();
    let pool = Arc::new(BrowserPool::new(BrowserPoolConfig {
        provider: BrowserProvider::ChromeExecutable(executable),
        ..BrowserPoolConfig::default()
    }));
    let injected: Arc<dyn PageRenderer> = pool.clone();
    let renderer = NativeOfficeScreenshotRenderer::new(injected);

    let result = renderer
        .render(
            &editor.snapshot().unwrap(),
            NativeOfficeScreenshotRequest::new(&output),
        )
        .await;
    pool.shutdown().await;

    let screenshot = result.unwrap();
    assert!(screenshot.width_px > 0);
    assert!(screenshot.height_px > 0);
    assert!(screenshot.byte_length > PNG_1X1.len() as u64);
    assert!(output.exists());
}
