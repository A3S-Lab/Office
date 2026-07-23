use std::io::{ErrorKind, Write as _};
use std::path::Path;

use a3s_use_core::{UseError, UseResult};

#[derive(Debug, Clone, Copy)]
pub(crate) enum OfficeArtifactKind {
    SemanticRender,
    #[cfg(feature = "browser")]
    Screenshot,
}

impl OfficeArtifactKind {
    fn label(self) -> &'static str {
        match self {
            Self::SemanticRender => "render",
            #[cfg(feature = "browser")]
            Self::Screenshot => "screenshot",
        }
    }

    fn error_code(self, exists: bool) -> &'static str {
        match (self, exists) {
            (Self::SemanticRender, true) => "use.office.render_output_exists",
            (Self::SemanticRender, false) => "use.office.render_output_failed",
            #[cfg(feature = "browser")]
            (Self::Screenshot, true) => "use.office.screenshot_output_exists",
            #[cfg(feature = "browser")]
            (Self::Screenshot, false) => "use.office.screenshot_output_failed",
        }
    }
}

pub(crate) async fn write_new(
    path: &Path,
    bytes: Vec<u8>,
    kind: OfficeArtifactKind,
) -> UseResult<()> {
    let path = path.to_path_buf();
    let task_path = path.clone();
    tokio::task::spawn_blocking(move || write_new_blocking(&task_path, &bytes, kind))
        .await
        .map_err(|error| {
            artifact_error(
                kind.error_code(false),
                &path,
                format!("Native Office {} output task failed: {error}", kind.label()),
            )
        })?
}

fn write_new_blocking(path: &Path, bytes: &[u8], kind: OfficeArtifactKind) -> UseResult<()> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut temporary = tempfile::NamedTempFile::new_in(parent).map_err(|error| {
        artifact_error(
            kind.error_code(false),
            path,
            format!(
                "Failed to stage native Office {} output: {error}",
                kind.label()
            ),
        )
    })?;
    temporary.write_all(bytes).map_err(|error| {
        artifact_error(
            kind.error_code(false),
            path,
            format!(
                "Failed to write staged native Office {} output: {error}",
                kind.label()
            ),
        )
    })?;
    temporary.as_file().sync_all().map_err(|error| {
        artifact_error(
            kind.error_code(false),
            path,
            format!(
                "Failed to sync staged native Office {} output: {error}",
                kind.label()
            ),
        )
    })?;
    temporary.persist_noclobber(path).map_err(|error| {
        if error.error.kind() == ErrorKind::AlreadyExists {
            artifact_error(
                kind.error_code(true),
                path,
                format!(
                    "Native Office {} output '{}' already exists; refusing to overwrite it.",
                    kind.label(),
                    path.display()
                ),
            )
            .with_suggestion(format!(
                "Choose a new native Office {} output path.",
                kind.label()
            ))
        } else {
            artifact_error(
                kind.error_code(false),
                path,
                format!(
                    "Failed to publish native Office {} output: {}",
                    kind.label(),
                    error.error
                ),
            )
        }
    })?;
    Ok(())
}

fn artifact_error(code: &'static str, path: &Path, message: impl Into<String>) -> UseError {
    UseError::new(code, message).with_detail("output", path.display().to_string())
}
