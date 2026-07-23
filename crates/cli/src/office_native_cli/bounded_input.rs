use a3s_use_core::{UseError, UseResult};
use tokio::io::AsyncReadExt;

#[derive(Debug, Clone, Copy)]
pub(super) enum NativeInputKind {
    Batch,
    Image,
    RawXml,
    SpreadsheetImport,
    TemplateData,
}

impl NativeInputKind {
    fn code_prefix(self) -> &'static str {
        match self {
            Self::Batch => "use.office.batch_input",
            Self::Image => "use.office.image_input",
            Self::RawXml => "use.office.raw_input",
            Self::SpreadsheetImport => "use.office.spreadsheet_import_input",
            Self::TemplateData => "use.office.template_data_input",
        }
    }

    fn label(self) -> &'static str {
        match self {
            Self::Batch => "Native Office batch input",
            Self::Image => "Native Office image input",
            Self::RawXml => "Native Office raw XML input",
            Self::SpreadsheetImport => "Native Spreadsheet delimited import input",
            Self::TemplateData => "Native Office template data input",
        }
    }
}

pub(super) async fn read_bounded_stdin(limit: u64, kind: NativeInputKind) -> UseResult<Vec<u8>> {
    let mut bytes = Vec::new();
    let mut reader = tokio::io::stdin().take(limit + 1);
    reader.read_to_end(&mut bytes).await.map_err(|error| {
        input_error(
            kind,
            "read_failed",
            "<stdin>",
            format!("Failed to read {} from stdin: {error}", kind.label()),
        )
    })?;
    if bytes.len() as u64 > limit {
        return Err(input_too_large(kind, "<stdin>", limit));
    }
    Ok(bytes)
}

pub(super) async fn read_bounded_input(
    path: &str,
    limit: u64,
    kind: NativeInputKind,
) -> UseResult<Vec<u8>> {
    let path_metadata = tokio::fs::symlink_metadata(path).await.map_err(|error| {
        input_error(
            kind,
            "open_failed",
            path,
            format!("Failed to inspect {} '{path}': {error}", kind.label()),
        )
    })?;
    if path_metadata.file_type().is_symlink() || !path_metadata.is_file() {
        return Err(input_error(
            kind,
            "invalid",
            path,
            format!("{} must be a regular, non-symlink file.", kind.label()),
        ));
    }
    if path_metadata.len() > limit {
        return Err(input_too_large(kind, path, limit));
    }

    let file = tokio::fs::File::open(path).await.map_err(|error| {
        input_error(
            kind,
            "open_failed",
            path,
            format!("Failed to open {} '{path}': {error}", kind.label()),
        )
    })?;
    let metadata = file.metadata().await.map_err(|error| {
        input_error(
            kind,
            "open_failed",
            path,
            format!("Failed to inspect {} '{path}': {error}", kind.label()),
        )
    })?;
    if !metadata.is_file() {
        return Err(input_error(
            kind,
            "invalid",
            path,
            format!("{} changed and is no longer a regular file.", kind.label()),
        ));
    }
    if metadata.len() > limit {
        return Err(input_too_large(kind, path, limit));
    }

    let mut bytes = Vec::with_capacity(metadata.len() as usize);
    let mut reader = file.take(limit + 1);
    reader.read_to_end(&mut bytes).await.map_err(|error| {
        input_error(
            kind,
            "read_failed",
            path,
            format!("Failed to read {} '{path}': {error}", kind.label()),
        )
    })?;
    if bytes.len() as u64 > limit {
        return Err(input_too_large(kind, path, limit));
    }
    Ok(bytes)
}

pub(super) fn input_too_large(kind: NativeInputKind, path: &str, limit: u64) -> UseError {
    input_error(
        kind,
        "too_large",
        path,
        format!("{} exceeds the {limit}-byte limit.", kind.label()),
    )
}

pub(super) fn input_error(
    kind: NativeInputKind,
    suffix: &str,
    path: &str,
    message: impl Into<String>,
) -> UseError {
    UseError::new(format!("{}_{suffix}", kind.code_prefix()), message).with_detail("input", path)
}

pub(super) fn batch_input_error(code: &str, path: &str, message: impl Into<String>) -> UseError {
    UseError::new(code, message).with_detail("input", path)
}
