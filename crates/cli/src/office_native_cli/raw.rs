use std::io::ErrorKind;

use a3s_office::{NativeOfficeEditor, NativeRawXmlPart};
use a3s_use_core::{UseError, UseResult};
use tokio::io::AsyncWriteExt;

use super::bounded_input::{input_error, read_bounded_input, NativeInputKind};
use super::{save_editor, usage_error, AllowedOptions, CommandOutput, ParsedArguments};

pub(super) const MAX_RAW_XML_INPUT_BYTES: u64 = 8 * 1024 * 1024;
const MAX_RAW_XML_INLINE_BYTES: u64 = 1024 * 1024;

pub(super) async fn inspect(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::RAW)?;
    if parsed.positionals.len() != 2 {
        return Err(usage_error("office native raw requires <file> and <part>"));
    }

    let editor = NativeOfficeEditor::open(&parsed.positionals[0]).await?;
    let part = editor.raw_xml_part(&parsed.positionals[1])?;
    let revision = editor.package().source_revision().clone();
    if let Some(output) = parsed.output.as_deref() {
        let bytes = editor.package().part(&part.part)?.to_vec();
        write_output(output, &bytes).await?;
        return Ok(CommandOutput::success(
            format!(
                "Exported {} to '{}' without modifying the Office document.",
                part.part, output
            ),
            serde_json::json!({
                "operation": "raw",
                "part": metadata(&part),
                "exported": true,
                "outputPath": output,
                "revision": revision
            }),
        ));
    }

    let inline_bytes = u64::try_from(part.xml.len()).unwrap_or(u64::MAX);
    if inline_bytes > MAX_RAW_XML_INLINE_BYTES {
        return Err(UseError::new(
            "use.office.raw_output_too_large",
            format!(
                "Normalized XML for '{}' exceeds the {MAX_RAW_XML_INLINE_BYTES}-byte inline output limit.",
                part.part
            ),
        )
        .with_suggestion("Use --output <xml-file> to export the original part bytes.")
        .with_detail("part", part.part)
        .with_detail("bytes", inline_bytes));
    }

    let human = part.xml.clone();
    Ok(CommandOutput::success(
        human,
        serde_json::json!({
            "operation": "raw",
            "part": metadata(&part),
            "xml": part.xml,
            "exported": false,
            "revision": revision
        }),
    ))
}

pub(super) async fn replace(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::RAW_SET)?;
    if parsed.positionals.len() != 2 {
        return Err(usage_error(
            "office native raw-set requires <file> and <part>",
        ));
    }
    let input_path = parsed
        .input
        .as_deref()
        .ok_or_else(|| usage_error("office native raw-set requires --input <xml-file>"))?;
    let xml = read_xml_input(input_path).await?;
    let source = &parsed.positionals[0];
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let part = editor.replace_xml_part(&parsed.positionals[1], xml)?;
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;

    Ok(CommandOutput::success(
        format!("Replaced {part} and saved '{}'.", output_path.display()),
        serde_json::json!({
            "operation": "raw-set",
            "changed": true,
            "part": part,
            "inputPath": input_path,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

fn metadata(part: &NativeRawXmlPart) -> serde_json::Value {
    serde_json::json!({
        "name": part.part,
        "byteLength": part.byte_length,
        "sha256": part.sha256,
        "encoding": part.encoding,
        "root": part.root
    })
}

async fn write_output(path: &str, bytes: &[u8]) -> UseResult<()> {
    let mut file = tokio::fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(path)
        .await
        .map_err(|error| {
            if error.kind() == ErrorKind::AlreadyExists {
                UseError::new(
                    "use.office.raw_output_exists",
                    format!("Raw XML output '{path}' already exists; refusing to overwrite it."),
                )
                .with_suggestion("Choose a new raw XML output path.")
                .with_detail("output", path)
            } else {
                UseError::new(
                    "use.office.raw_output_write_failed",
                    format!("Failed to create raw XML output '{path}': {error}"),
                )
                .with_detail("output", path)
            }
        })?;
    let result = async {
        file.write_all(bytes).await?;
        file.sync_all().await
    }
    .await;
    if let Err(error) = result {
        drop(file);
        let _ = tokio::fs::remove_file(path).await;
        return Err(UseError::new(
            "use.office.raw_output_write_failed",
            format!("Failed to write raw XML output '{path}': {error}"),
        )
        .with_detail("output", path));
    }
    Ok(())
}

pub(super) async fn read_xml_input(path: &str) -> UseResult<String> {
    let bytes = read_bounded_input(path, MAX_RAW_XML_INPUT_BYTES, NativeInputKind::RawXml).await?;
    String::from_utf8(bytes).map_err(|error| {
        input_error(
            NativeInputKind::RawXml,
            "invalid",
            path,
            format!("Native Office raw XML input '{path}' is not UTF-8: {error}"),
        )
    })
}
