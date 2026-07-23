//! Browser-injected screenshot rendering for native Office semantic HTML.

use std::fmt;
use std::io::Write as _;
use std::path::{Path, PathBuf};
use std::sync::Arc;

use a3s_office::{DocumentKind, NativeOfficeDocument, NativeOfficeImage, NativeOfficeImageFormat};
use a3s_use_browser::{BrowserPool, BrowserPoolConfig, PageRenderer, RenderRequest, WaitCondition};
use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use url::Url;

use crate::office_artifact::{self, OfficeArtifactKind};

pub const DEFAULT_NATIVE_OFFICE_SCREENSHOT_TIMEOUT_MS: u64 = 30_000;
pub const MAX_NATIVE_OFFICE_SCREENSHOT_TIMEOUT_MS: u64 = 120_000;
pub const MAX_NATIVE_OFFICE_SCREENSHOT_BYTES: u64 = 64 * 1024 * 1024;

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeScreenshotRequest {
    pub output: PathBuf,
    pub timeout_ms: u64,
}

impl NativeOfficeScreenshotRequest {
    pub fn new(output: impl Into<PathBuf>) -> Self {
        Self {
            output: output.into(),
            timeout_ms: DEFAULT_NATIVE_OFFICE_SCREENSHOT_TIMEOUT_MS,
        }
    }

    pub fn with_timeout_ms(mut self, timeout_ms: u64) -> Self {
        self.timeout_ms = timeout_ms;
        self
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeScreenshot {
    pub kind: DocumentKind,
    pub output_path: PathBuf,
    pub media_type: String,
    pub width_px: u32,
    pub height_px: u32,
    pub byte_length: u64,
    pub sha256: String,
    pub source_html_sha256: String,
    pub renderer_elapsed_ms: u64,
}

#[derive(Clone)]
pub struct NativeOfficeScreenshotRenderer {
    renderer: Arc<dyn PageRenderer>,
}

impl NativeOfficeScreenshotRenderer {
    pub fn new(renderer: Arc<dyn PageRenderer>) -> Self {
        Self { renderer }
    }

    pub async fn render(
        &self,
        document: &NativeOfficeDocument,
        request: NativeOfficeScreenshotRequest,
    ) -> UseResult<NativeOfficeScreenshot> {
        validate_request(&request).await?;
        let html = document.html_view()?;
        let kind = html.kind;
        let source_html_sha256 = html.sha256;
        let staging = ScreenshotStaging::prepare(html.content).await?;
        let url = Url::from_file_path(&staging.html_path).map_err(|_| {
            screenshot_error(
                "use.office.screenshot_staging_failed",
                "Failed to create a local URL for the native Office semantic preview.",
            )
        })?;
        let rendered = self
            .renderer
            .render(RenderRequest {
                url,
                timeout_ms: request.timeout_ms,
                wait: WaitCondition::Load,
                user_agent: Some("a3s-office-semantic-preview/1".to_string()),
                screenshot_path: Some(staging.screenshot_path.clone()),
            })
            .await?;
        let validated = validate_screenshot(&rendered, &staging.screenshot_path).await?;
        office_artifact::write_new(
            &request.output,
            validated.bytes,
            OfficeArtifactKind::Screenshot,
        )
        .await?;
        Ok(NativeOfficeScreenshot {
            kind,
            output_path: request.output,
            media_type: "image/png".to_string(),
            width_px: validated.width_px,
            height_px: validated.height_px,
            byte_length: validated.byte_length,
            sha256: validated.sha256,
            source_html_sha256,
            renderer_elapsed_ms: rendered.elapsed_ms,
        })
    }
}

/// Captures through the same discovered Browser provider used by `a3s use browser`.
pub async fn capture_native_office_screenshot(
    document: &NativeOfficeDocument,
    request: NativeOfficeScreenshotRequest,
) -> UseResult<NativeOfficeScreenshot> {
    let pool = Arc::new(BrowserPool::new(BrowserPoolConfig::default()));
    let injected: Arc<dyn PageRenderer> = pool.clone();
    let renderer = NativeOfficeScreenshotRenderer::new(injected);
    let result = renderer.render(document, request).await;
    pool.shutdown().await;
    result
}

impl fmt::Debug for NativeOfficeScreenshotRenderer {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("NativeOfficeScreenshotRenderer")
            .finish_non_exhaustive()
    }
}

async fn validate_request(request: &NativeOfficeScreenshotRequest) -> UseResult<()> {
    if request.output.as_os_str().is_empty()
        || request.output == Path::new("-")
        || !request
            .output
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("png"))
    {
        return Err(screenshot_error(
            "use.office.screenshot_output_invalid",
            "Native Office screenshot output must be a local path ending in '.png'.",
        ));
    }
    if !(1..=MAX_NATIVE_OFFICE_SCREENSHOT_TIMEOUT_MS).contains(&request.timeout_ms) {
        return Err(screenshot_error(
            "use.office.screenshot_timeout_invalid",
            format!(
                "Native Office screenshot timeout must be between 1 and {MAX_NATIVE_OFFICE_SCREENSHOT_TIMEOUT_MS} ms."
            ),
        ));
    }
    match tokio::fs::symlink_metadata(&request.output).await {
        Ok(_) => Err(screenshot_error(
            "use.office.screenshot_output_exists",
            format!(
                "Native Office screenshot output '{}' already exists; refusing to overwrite it.",
                request.output.display()
            ),
        )
        .with_suggestion("Choose a new PNG output path.")),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(error) => Err(screenshot_error(
            "use.office.screenshot_output_failed",
            format!(
                "Failed to inspect native Office screenshot output '{}': {error}",
                request.output.display()
            ),
        )),
    }
}

struct ScreenshotStaging {
    _directory: tempfile::TempDir,
    html_path: PathBuf,
    screenshot_path: PathBuf,
}

impl ScreenshotStaging {
    async fn prepare(html: String) -> UseResult<Self> {
        tokio::task::spawn_blocking(move || {
            let directory = tempfile::Builder::new()
                .prefix("a3s-office-screenshot-")
                .tempdir()
                .map_err(|error| {
                    screenshot_error(
                        "use.office.screenshot_staging_failed",
                        format!("Failed to create native Office screenshot staging: {error}"),
                    )
                })?;
            let html_path = directory.path().join("semantic-preview.html");
            let screenshot_path = directory.path().join("semantic-preview.png");
            let mut file = std::fs::OpenOptions::new()
                .write(true)
                .create_new(true)
                .open(&html_path)
                .map_err(|error| {
                    screenshot_error(
                        "use.office.screenshot_staging_failed",
                        format!("Failed to create native Office screenshot HTML: {error}"),
                    )
                })?;
            file.write_all(html.as_bytes()).map_err(|error| {
                screenshot_error(
                    "use.office.screenshot_staging_failed",
                    format!("Failed to stage native Office screenshot HTML: {error}"),
                )
            })?;
            file.sync_all().map_err(|error| {
                screenshot_error(
                    "use.office.screenshot_staging_failed",
                    format!("Failed to sync native Office screenshot HTML: {error}"),
                )
            })?;
            Ok(Self {
                _directory: directory,
                html_path,
                screenshot_path,
            })
        })
        .await
        .map_err(|error| {
            screenshot_error(
                "use.office.screenshot_staging_failed",
                format!("Native Office screenshot staging task failed: {error}"),
            )
        })?
    }
}

struct ValidatedScreenshot {
    bytes: Vec<u8>,
    width_px: u32,
    height_px: u32,
    byte_length: u64,
    sha256: String,
}

async fn validate_screenshot(
    rendered: &a3s_use_browser::RenderedPage,
    expected_path: &Path,
) -> UseResult<ValidatedScreenshot> {
    if rendered.artifacts.len() != 1 {
        return Err(screenshot_error(
            "use.office.screenshot_artifact_invalid",
            "The Browser renderer must return exactly one Office screenshot artifact.",
        ));
    }
    let artifact = &rendered.artifacts[0];
    if artifact.path != expected_path || artifact.media_type != "image/png" {
        return Err(screenshot_error(
            "use.office.screenshot_artifact_invalid",
            "The Browser renderer returned an unexpected Office screenshot artifact.",
        ));
    }
    let metadata = tokio::fs::symlink_metadata(expected_path)
        .await
        .map_err(|error| {
            screenshot_error(
                "use.office.screenshot_artifact_missing",
                format!("The Browser renderer did not create its PNG artifact: {error}"),
            )
        })?;
    if !metadata.is_file() || metadata.file_type().is_symlink() {
        return Err(screenshot_error(
            "use.office.screenshot_artifact_invalid",
            "The Browser renderer screenshot is not a regular, non-symlink PNG file.",
        ));
    }
    if metadata.len() == 0 || metadata.len() > MAX_NATIVE_OFFICE_SCREENSHOT_BYTES {
        return Err(screenshot_error(
            "use.office.screenshot_artifact_too_large",
            format!(
                "Native Office screenshot is {} bytes; the limit is {MAX_NATIVE_OFFICE_SCREENSHOT_BYTES}.",
                metadata.len()
            ),
        ));
    }
    let bytes = tokio::fs::read(expected_path).await.map_err(|error| {
        screenshot_error(
            "use.office.screenshot_artifact_invalid",
            format!("Failed to read the Browser renderer screenshot: {error}"),
        )
    })?;
    let byte_length = u64::try_from(bytes.len()).unwrap_or(u64::MAX);
    if byte_length == 0 || byte_length > MAX_NATIVE_OFFICE_SCREENSHOT_BYTES {
        return Err(screenshot_error(
            "use.office.screenshot_artifact_too_large",
            format!(
                "Native Office screenshot is {byte_length} bytes; the limit is {MAX_NATIVE_OFFICE_SCREENSHOT_BYTES}."
            ),
        ));
    }
    let image = NativeOfficeImage::inspect_bytes(&bytes).map_err(|error| {
        screenshot_error(
            "use.office.screenshot_artifact_invalid",
            format!("The Browser renderer screenshot is not a valid PNG: {error}"),
        )
    })?;
    if image.format != NativeOfficeImageFormat::Png {
        return Err(screenshot_error(
            "use.office.screenshot_artifact_invalid",
            "The Browser renderer screenshot is not PNG data.",
        ));
    }
    let sha256 = format!("{:x}", Sha256::digest(&bytes));
    if artifact.size != byte_length || !artifact.sha256.eq_ignore_ascii_case(&sha256) {
        return Err(screenshot_error(
            "use.office.screenshot_artifact_invalid",
            "The Browser renderer screenshot receipt does not match the PNG bytes.",
        ));
    }
    Ok(ValidatedScreenshot {
        bytes,
        width_px: image.width_px,
        height_px: image.height_px,
        byte_length,
        sha256,
    })
}

fn screenshot_error(code: &'static str, message: impl Into<String>) -> UseError {
    UseError::new(code, message)
}

#[cfg(test)]
mod tests;
