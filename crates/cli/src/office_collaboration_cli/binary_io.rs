use std::path::{Path, PathBuf};

use a3s_use_core::UseResult;
use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;

use super::arguments::{collaboration_cli_error, usage_error};

pub(super) async fn read_optional_encoded_input(
    inline: Option<&str>,
    path: Option<&str>,
    max_bytes: usize,
    label: &str,
) -> UseResult<Option<Vec<u8>>> {
    match (inline, path) {
        (Some(_), Some(_)) => Err(usage_error(format!(
            "Specify either inline --{label} base64 or its input-file option, not both."
        ))),
        (Some(value), None) => {
            let max_encoded_bytes = max_bytes.saturating_add(2) / 3 * 4;
            if value.len() > max_encoded_bytes {
                return Err(collaboration_cli_error(
                    "office.collaboration.input_too_large",
                    format!(
                        "The inline collaboration {label} exceeds the encoded limit of {max_encoded_bytes} bytes."
                    ),
                ));
            }
            let bytes = STANDARD.decode(value).map_err(|error| {
                usage_error(format!("The inline {label} is not valid base64: {error}"))
            })?;
            if bytes.len() > max_bytes {
                return Err(collaboration_cli_error(
                    "office.collaboration.input_too_large",
                    format!(
                        "The inline collaboration {label} is {} bytes; the limit is {} bytes.",
                        bytes.len(),
                        max_bytes
                    ),
                ));
            }
            Ok(Some(bytes))
        }
        (None, Some(path)) => read_binary_input(&PathBuf::from(path), max_bytes, label)
            .await
            .map(Some),
        (None, None) => Ok(None),
    }
}

pub(super) async fn read_binary_input(
    path: &Path,
    max_bytes: usize,
    label: &str,
) -> UseResult<Vec<u8>> {
    let path = path.to_path_buf();
    let label = label.to_owned();
    spawn_blocking(move || {
        use std::io::Read as _;

        let metadata = std::fs::symlink_metadata(&path).map_err(|error| {
            collaboration_cli_error(
                "office.collaboration.input_unavailable",
                format!(
                    "Collaboration {label} '{}' is unavailable: {error}",
                    path.display()
                ),
            )
            .with_detail("input", path.display().to_string())
        })?;
        if !metadata.is_file() || metadata.file_type().is_symlink() {
            return Err(collaboration_cli_error(
                "office.collaboration.input_invalid",
                format!(
                    "Collaboration {label} '{}' must be a real file.",
                    path.display()
                ),
            ));
        }
        if metadata.len() > max_bytes as u64 {
            return Err(collaboration_cli_error(
                "office.collaboration.input_too_large",
                format!(
                    "Collaboration {label} '{}' is {} bytes; the limit is {} bytes.",
                    path.display(),
                    metadata.len(),
                    max_bytes
                ),
            ));
        }
        let file = std::fs::File::open(&path).map_err(|error| {
            collaboration_cli_error(
                "office.collaboration.input_failed",
                format!(
                    "Failed to open collaboration {label} '{}': {error}",
                    path.display()
                ),
            )
        })?;
        let mut bytes = Vec::new();
        file.take(max_bytes.saturating_add(1) as u64)
            .read_to_end(&mut bytes)
            .map_err(|error| {
                collaboration_cli_error(
                    "office.collaboration.input_failed",
                    format!(
                        "Failed to read collaboration {label} '{}': {error}",
                        path.display()
                    ),
                )
            })?;
        if bytes.len() > max_bytes {
            return Err(collaboration_cli_error(
                "office.collaboration.input_too_large",
                format!(
                    "Collaboration {label} '{}' grew beyond the limit of {} bytes while being read.",
                    path.display(),
                    max_bytes
                ),
            ));
        }
        Ok(bytes)
    })
    .await
}

pub(super) async fn write_new_binary(
    path: PathBuf,
    bytes: Vec<u8>,
    label: &'static str,
) -> UseResult<()> {
    spawn_blocking(move || {
        use std::io::{ErrorKind, Write as _};

        let parent = path
            .parent()
            .filter(|parent| !parent.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        let mut temporary = tempfile::NamedTempFile::new_in(parent).map_err(|error| {
            collaboration_cli_error(
                "office.collaboration.output_failed",
                format!("Failed to stage {label} '{}': {error}", path.display()),
            )
        })?;
        temporary.write_all(&bytes).map_err(|error| {
            collaboration_cli_error(
                "office.collaboration.output_failed",
                format!(
                    "Failed to write staged {label} '{}': {error}",
                    path.display()
                ),
            )
        })?;
        temporary.as_file().sync_all().map_err(|error| {
            collaboration_cli_error(
                "office.collaboration.output_failed",
                format!(
                    "Failed to sync staged {label} '{}': {error}",
                    path.display()
                ),
            )
        })?;
        temporary.persist_noclobber(&path).map_err(|error| {
            if error.error.kind() == ErrorKind::AlreadyExists {
                collaboration_cli_error(
                    "office.collaboration.output_exists",
                    format!(
                        "{label} '{}' already exists; refusing to overwrite it.",
                        path.display()
                    ),
                )
            } else {
                collaboration_cli_error(
                    "office.collaboration.output_failed",
                    format!(
                        "Failed to publish {label} '{}': {}",
                        path.display(),
                        error.error
                    ),
                )
            }
        })?;
        Ok(())
    })
    .await
}

pub(super) async fn spawn_blocking<T, F>(operation: F) -> UseResult<T>
where
    T: Send + 'static,
    F: FnOnce() -> UseResult<T> + Send + 'static,
{
    tokio::task::spawn_blocking(operation)
        .await
        .map_err(|error| {
            collaboration_cli_error(
                "office.collaboration.task_failed",
                format!("The native collaboration task failed: {error}"),
            )
        })?
}
