use std::path::{Path, PathBuf};

use a3s_office::MAX_NATIVE_OFFICE_REPLAY_MUTATIONS;
use a3s_use_core::{UseError, UseResult, UseSessionId};
use rmcp::model::CallToolResult;

use super::input::OfficeMutation;
use super::session::NativeOfficeSession;

pub(super) const MAX_BATCH_BYTES: usize = 8 * 1024 * 1024;
pub(super) const MAX_RAW_XML_BYTES: u64 = 1024 * 1024;
const MAX_RESULT_BYTES: usize = 8 * 1024 * 1024;

pub(super) fn session_value(
    session: &UseSessionId,
    state: &NativeOfficeSession,
) -> serde_json::Value {
    serde_json::json!({
        "session": session,
        "path": state.editor.package().path(),
        "kind": state.editor.package().kind(),
        "readOnly": state.read_only,
        "dirty": state.editor.is_dirty(),
        "revision": state.editor.package().source_revision(),
        "contentSha256": state.editor.package().content_sha256()
    })
}

pub(super) fn validate_batch(mutations: &[OfficeMutation]) -> UseResult<()> {
    if mutations.is_empty() {
        return Err(UseError::new(
            "use.office.batch_empty",
            "Native Office MCP mutation batches cannot be empty.",
        ));
    }
    if mutations.len() > MAX_NATIVE_OFFICE_REPLAY_MUTATIONS {
        return Err(UseError::new(
            "use.office.batch_mutation_limit",
            format!(
                "Native Office MCP batch contains {} mutations; the limit is {MAX_NATIVE_OFFICE_REPLAY_MUTATIONS}.",
                mutations.len()
            ),
        ));
    }
    validate_json_bytes(mutations, MAX_BATCH_BYTES, "mutation batch")
}

pub(super) fn validate_json_bytes<T: serde::Serialize + ?Sized>(
    value: &T,
    limit: usize,
    label: &str,
) -> UseResult<()> {
    let bytes = serde_json::to_vec(value).map_err(output_encoding_error)?;
    if bytes.len() > limit {
        return Err(UseError::new(
            "use.office.mcp_input_too_large",
            format!("Native Office MCP {label} exceeds the {limit}-byte limit."),
        )
        .with_detail("bytes", bytes.len()));
    }
    Ok(())
}

pub(super) fn tool_result(result: UseResult<serde_json::Value>) -> CallToolResult {
    match result {
        Ok(value) => match serde_json::to_vec(&value) {
            Ok(bytes) if bytes.len() <= MAX_RESULT_BYTES => CallToolResult::structured(value),
            Ok(bytes) => tool_error(
                UseError::new(
                    "use.office.mcp_result_too_large",
                    format!(
                        "Native Office MCP result is {} bytes; the limit is {MAX_RESULT_BYTES}.",
                        bytes.len()
                    ),
                )
                .with_suggestion("Narrow the semantic path, query selector, or requested view."),
            ),
            Err(error) => tool_error(output_encoding_error(error)),
        },
        Err(error) => tool_error(error),
    }
}

fn tool_error(error: UseError) -> CallToolResult {
    CallToolResult::structured_error(serde_json::to_value(error).unwrap_or_else(|_| {
        serde_json::json!({
            "code": "use.error_encoding_failed",
            "message": "Failed to encode A3S Use error."
        })
    }))
}

pub(super) fn output_encoding_error(error: serde_json::Error) -> UseError {
    UseError::new(
        "use.office.output_invalid",
        format!("Failed to encode native Office output: {error}"),
    )
}

pub(super) async fn reject_same_file(template: &Path, output: &Path) -> UseResult<()> {
    let template_path = canonical_or_absolute(template).await?;
    let output_path = canonical_or_absolute(output).await?;
    if template_path == output_path || existing_file_identity_matches(template, output).await? {
        return Err(UseError::new(
            "use.office.template_output_same_file",
            "Native Office template and output must be different files.",
        )
        .with_suggestion("Choose a separate output path; the template is never modified in place.")
        .with_detail("template", template_path.display().to_string())
        .with_detail("output", output_path.display().to_string()));
    }
    Ok(())
}

async fn canonical_or_absolute(path: &Path) -> UseResult<PathBuf> {
    if let Ok(canonical) = tokio::fs::canonicalize(path).await {
        return Ok(canonical);
    }
    let absolute = if path.is_absolute() {
        path.to_path_buf()
    } else {
        std::env::current_dir()
            .map_err(|error| {
                UseError::new(
                    "use.office.template_path_invalid",
                    format!("Failed to resolve the current directory: {error}"),
                )
            })?
            .join(path)
    };
    if let (Some(parent), Some(name)) = (absolute.parent(), absolute.file_name()) {
        if let Ok(parent) = tokio::fs::canonicalize(parent).await {
            return Ok(parent.join(name));
        }
    }
    Ok(absolute)
}

#[cfg(unix)]
async fn existing_file_identity_matches(left: &Path, right: &Path) -> UseResult<bool> {
    use std::os::unix::fs::MetadataExt;

    let left = optional_metadata(left).await?;
    let right = optional_metadata(right).await?;
    Ok(
        matches!((left, right), (Some(left), Some(right)) if left.dev() == right.dev() && left.ino() == right.ino()),
    )
}

#[cfg(not(unix))]
async fn existing_file_identity_matches(_left: &Path, _right: &Path) -> UseResult<bool> {
    Ok(false)
}

#[cfg(unix)]
async fn optional_metadata(path: &Path) -> UseResult<Option<std::fs::Metadata>> {
    match tokio::fs::metadata(path).await {
        Ok(metadata) => Ok(Some(metadata)),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(None),
        Err(error) => Err(UseError::new(
            "use.office.template_path_invalid",
            format!(
                "Failed to inspect native Office merge path '{}': {error}",
                path.display()
            ),
        )
        .with_detail("path", path.display().to_string())),
    }
}
