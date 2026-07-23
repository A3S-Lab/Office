use std::io::ErrorKind;

use a3s_office::{
    NativeOfficeDocument, NativeOfficeMutation, NativeOfficeReplayArtifact,
    MAX_NATIVE_OFFICE_REPLAY_MUTATIONS,
};
use a3s_use_core::{UseError, UseResult};
use serde::Deserialize;
use tokio::io::AsyncWriteExt;

use super::bounded_input::{batch_input_error, read_bounded_input, NativeInputKind};
use super::{usage_error, AllowedOptions, CommandOutput, ParsedArguments, MAX_BATCH_INPUT_BYTES};

const MAX_DUMP_OUTPUT_BYTES: usize = 8 * 1024 * 1024;
const MAX_INLINE_DUMP_OUTPUT_BYTES: usize = 1024 * 1024;

#[derive(Debug)]
pub(super) enum NativeBatchInput {
    Mutations(Vec<NativeOfficeMutation>),
    Replay(NativeOfficeReplayArtifact),
}

impl NativeBatchInput {
    pub(super) fn mutation_count(&self) -> usize {
        match self {
            Self::Mutations(mutations) => mutations.len(),
            Self::Replay(artifact) => artifact.mutations.len(),
        }
    }
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct NativeMutationBatch {
    schema_version: u32,
    mutations: Vec<NativeOfficeMutation>,
}

pub(super) async fn dump(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::DUMP)?;
    if !(1..=2).contains(&parsed.positionals.len()) {
        return Err(usage_error(
            "office native dump requires <file> and an optional [path]",
        ));
    }
    let source = &parsed.positionals[0];
    let scope = parsed.positionals.get(1).map_or("/", String::as_str);
    let document = NativeOfficeDocument::open(source).await?;
    let artifact = NativeOfficeReplayArtifact::dump(&document, scope)?;
    let encoded = serde_json::to_string_pretty(&artifact).map_err(|error| {
        UseError::new(
            "use.office.dump_encode_failed",
            format!("Failed to encode native Office replay artifact: {error}"),
        )
    })?;
    let byte_length = encoded.len().checked_add(1).ok_or_else(|| {
        UseError::new(
            "use.office.dump_output_too_large",
            "Native Office replay artifact length overflowed.",
        )
    })?;
    if byte_length > MAX_DUMP_OUTPUT_BYTES {
        return Err(UseError::new(
            "use.office.dump_output_too_large",
            format!(
                "Native Office replay artifact is {byte_length} bytes; the limit is {MAX_DUMP_OUTPUT_BYTES}."
            ),
        )
        .with_detail("bytes", byte_length));
    }

    if let Some(output) = parsed.output.as_deref().filter(|output| *output != "-") {
        write_new_output(output, format!("{encoded}\n").as_bytes()).await?;
        return Ok(CommandOutput::success(
            format!("Wrote replayable native Office batch to '{output}'."),
            serde_json::json!({
                "operation": "dump",
                "outputFile": output,
                "bytes": byte_length,
                "format": artifact.format,
                "artifactSchemaVersion": artifact.schema_version,
                "documentKind": artifact.document_kind,
                "scope": artifact.scope,
                "baseSha256": artifact.base_sha256,
                "resultSha256": artifact.result_sha256,
                "mutations": artifact.mutations.len()
            }),
        ));
    }

    if byte_length > MAX_INLINE_DUMP_OUTPUT_BYTES {
        return Err(UseError::new(
            "use.office.dump_inline_too_large",
            format!(
                "Native Office replay artifact is {byte_length} bytes; inline output is limited to {MAX_INLINE_DUMP_OUTPUT_BYTES}."
            ),
        )
        .with_suggestion("Pass --output <batch.json> to write the bounded artifact to a new file."));
    }
    Ok(CommandOutput::success(
        encoded,
        serde_json::json!({
            "operation": "dump",
            "bytes": byte_length,
            "artifact": artifact
        }),
    ))
}

pub(super) async fn read_batch_input(path: &str) -> UseResult<NativeBatchInput> {
    let bytes = read_bounded_input(path, MAX_BATCH_INPUT_BYTES, NativeInputKind::Batch).await?;
    let value: serde_json::Value = serde_json::from_slice(&bytes).map_err(|error| {
        batch_input_error(
            "use.office.batch_input_invalid",
            path,
            format!("Native Office batch input '{path}' is invalid JSON: {error}"),
        )
    })?;
    let is_replay = value
        .as_object()
        .is_some_and(|object| object.contains_key("format"));
    if is_replay {
        let artifact: NativeOfficeReplayArtifact =
            serde_json::from_value(value).map_err(|error| {
                batch_input_error(
                    "use.office.batch_input_invalid",
                    path,
                    format!("Native Office replay input '{path}' is invalid: {error}"),
                )
            })?;
        artifact
            .validate()
            .map_err(|error| error.with_detail("input", path))?;
        return Ok(NativeBatchInput::Replay(artifact));
    }

    let input: NativeMutationBatch = serde_json::from_value(value).map_err(|error| {
        batch_input_error(
            "use.office.batch_input_invalid",
            path,
            format!("Native Office batch input '{path}' is invalid: {error}"),
        )
    })?;
    if input.schema_version != 1 {
        return Err(batch_input_error(
            "use.office.batch_schema_unsupported",
            path,
            format!(
                "Native Office batch schema version {} is not supported; expected 1.",
                input.schema_version
            ),
        ));
    }
    if input.mutations.len() > MAX_NATIVE_OFFICE_REPLAY_MUTATIONS {
        return Err(batch_input_error(
            "use.office.batch_mutation_limit",
            path,
            format!(
                "Native Office batch contains {} mutations; the limit is {MAX_NATIVE_OFFICE_REPLAY_MUTATIONS}.",
                input.mutations.len()
            ),
        ));
    }
    Ok(NativeBatchInput::Mutations(input.mutations))
}

async fn write_new_output(path: &str, bytes: &[u8]) -> UseResult<()> {
    let mut output = match tokio::fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(path)
        .await
    {
        Ok(output) => output,
        Err(error) if error.kind() == ErrorKind::AlreadyExists => {
            return Err(UseError::new(
                "use.office.dump_output_exists",
                format!(
                    "Native Office dump output '{path}' already exists; refusing to overwrite it."
                ),
            )
            .with_detail("output", path));
        }
        Err(error) => {
            return Err(UseError::new(
                "use.office.dump_output_failed",
                format!("Failed to create native Office dump output '{path}': {error}"),
            )
            .with_detail("output", path));
        }
    };
    if let Err(error) = output.write_all(bytes).await {
        drop(output);
        let _ = tokio::fs::remove_file(path).await;
        return Err(UseError::new(
            "use.office.dump_output_failed",
            format!("Failed to write native Office dump output '{path}': {error}"),
        )
        .with_detail("output", path));
    }
    if let Err(error) = output.flush().await {
        drop(output);
        let _ = tokio::fs::remove_file(path).await;
        return Err(UseError::new(
            "use.office.dump_output_failed",
            format!("Failed to flush native Office dump output '{path}': {error}"),
        )
        .with_detail("output", path));
    }
    Ok(())
}
