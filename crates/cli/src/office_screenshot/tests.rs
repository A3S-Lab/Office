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
    redirect: bool,
}

struct UnitFixtureRenderer {
    calls: Arc<AtomicUsize>,
}

#[async_trait]
impl PageRenderer for UnitFixtureRenderer {
    async fn render(&self, request: RenderRequest) -> UseResult<RenderedPage> {
        self.calls.fetch_add(1, Ordering::SeqCst);
        assert_eq!(request.url.scheme(), "file");
        assert_eq!(request.wait, WaitCondition::Load);
        let html_path = request.url.to_file_path().unwrap();
        let html = tokio::fs::read_to_string(html_path).await.unwrap();
        assert!(html.contains("SECOND-SHEET-MARKER"));
        assert!(!html.contains("FIRST-SHEET-MARKER"));
        assert!(html.contains(
            "data-document-kind=\"spreadsheet\" data-unit-path=\"/Data\" data-unit-ordinal=\"2\">"
        ));
        let screenshot_path = request.screenshot_path.unwrap();
        tokio::fs::write(&screenshot_path, PNG_1X1).await.unwrap();
        Ok(RenderedPage {
            requested_url: request.url.clone(),
            final_url: request.url,
            status: None,
            content_type: Some("text/html".to_string()),
            html,
            elapsed_ms: 11,
            artifacts: vec![Artifact {
                path: screenshot_path,
                media_type: "image/png".to_string(),
                size: PNG_1X1.len() as u64,
                sha256: format!("{:x}", Sha256::digest(PNG_1X1)),
            }],
        })
    }
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
            final_url: if self.redirect {
                Url::parse("https://example.invalid/redirected").unwrap()
            } else {
                request.url
            },
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
        redirect: false,
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
        redirect: false,
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

    let redirected_output = temp.path().join("redirected.png");
    let renderer = NativeOfficeScreenshotRenderer::new(Arc::new(FixtureRenderer {
        calls: Arc::new(AtomicUsize::new(0)),
        corrupt: false,
        redirect: true,
    }));
    let error = renderer
        .render(
            &editor.snapshot().unwrap(),
            NativeOfficeScreenshotRequest::new(&redirected_output),
        )
        .await
        .unwrap_err();
    assert_eq!(error.code, "use.office.screenshot_navigation_invalid");
    assert!(!redirected_output.exists());
}

#[tokio::test]
async fn unit_screenshot_renderer_captures_exact_semantic_unit_identity() {
    let temp = tempfile::tempdir().unwrap();
    let document_path = temp.path().join("workbook.xlsx");
    let output = temp.path().join("data.png");
    let mut editor = NativeOfficeEditor::create(&document_path).await.unwrap();
    editor.set_text("/Sheet1/A1", "FIRST-SHEET-MARKER").unwrap();
    editor.add_worksheet("Data").unwrap();
    editor.set_text("/Data/A1", "SECOND-SHEET-MARKER").unwrap();
    let document = editor.snapshot().unwrap();
    let locator = NativeOfficeUnitLocator::Worksheet {
        index: 2,
        name: "Data".to_string(),
    };
    let semantic = document
        .render_unit(
            &locator,
            NativeOfficeUnitRenderOptions {
                format: NativeOfficeRenderFormat::Html,
                max_output_bytes: MAX_NATIVE_OFFICE_RENDER_BYTES,
            },
        )
        .unwrap();
    let calls = Arc::new(AtomicUsize::new(0));
    let renderer = NativeOfficeScreenshotRenderer::new(Arc::new(UnitFixtureRenderer {
        calls: Arc::clone(&calls),
    }));

    let screenshot = renderer
        .render_unit(
            &document,
            &locator,
            NativeOfficeScreenshotRequest::new(&output),
        )
        .await
        .unwrap();

    assert_eq!(screenshot.kind, DocumentKind::Spreadsheet);
    assert_eq!(screenshot.unit, semantic.unit);
    assert_eq!(screenshot.document_unit_count, 2);
    assert_eq!(screenshot.output_path, output);
    assert_eq!(screenshot.source_html_byte_length, semantic.byte_length);
    assert_eq!(screenshot.source_html_sha256, semantic.sha256);
    assert_eq!(screenshot.renderer_elapsed_ms, 11);
    assert_eq!(std::fs::read(&output).unwrap(), PNG_1X1);
    assert_eq!(calls.load(Ordering::SeqCst), 1);

    let rejected_output = temp.path().join("rejected.png");
    let error = renderer
        .render_unit(
            &document,
            &NativeOfficeUnitLocator::Worksheet {
                index: 1,
                name: "Data".to_string(),
            },
            NativeOfficeScreenshotRequest::new(&rejected_output),
        )
        .await
        .unwrap_err();
    assert_eq!(error.code, "use.office.unit_identity_mismatch");
    assert_eq!(calls.load(Ordering::SeqCst), 1);
    assert!(!rejected_output.exists());
}

#[test]
fn public_screenshot_contracts_are_send_and_sync() {
    fn assert_send_sync<T: Send + Sync>() {}

    assert_send_sync::<NativeOfficeScreenshotRequest>();
    assert_send_sync::<NativeOfficeScreenshot>();
    assert_send_sync::<NativeOfficeUnitScreenshot>();
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
