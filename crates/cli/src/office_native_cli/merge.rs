use std::path::{Path, PathBuf};

use a3s_office::NativeOfficeEditor;
use a3s_use_core::{UseError, UseResult};

use super::bounded_input::{input_error, input_too_large, read_bounded_input, NativeInputKind};
use super::{usage_error, AllowedOptions, CommandOutput, ParsedArguments};

pub(super) const MAX_TEMPLATE_DATA_INPUT_BYTES: u64 = 8 * 1024 * 1024;

pub(super) async fn run(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::MERGE)?;
    if parsed.positionals.len() != 2 {
        return Err(usage_error(
            "office native merge requires <template> and <output>",
        ));
    }
    let data_argument = parsed
        .data
        .as_deref()
        .ok_or_else(|| usage_error("office native merge requires --data <json|@file.json>"))?;
    let template = Path::new(&parsed.positionals[0]);
    let output = Path::new(&parsed.positionals[1]);
    reject_same_file(template, output)?;
    let template_path = canonical_or_absolute(template)?;

    let (data, data_source) = read_template_data(data_argument).await?;
    let mut editor = NativeOfficeEditor::open(template).await?;
    let result = editor.merge_template(&data)?;
    if parsed.force {
        editor.save_as(output).await?;
    } else {
        editor.save_as_new(output).await?;
    }
    let output_path = editor.package().path().to_path_buf();
    let unresolved_count = result.unresolved_placeholders.len();
    let changed = !result.changed_parts.is_empty();
    let human = if unresolved_count == 0 {
        format!(
            "Merged {} replacement(s) into '{}'.",
            result.replaced_count,
            output_path.display()
        )
    } else {
        format!(
            "Merged {} replacement(s) into '{}'; {unresolved_count} placeholder(s) remain unresolved.",
            result.replaced_count,
            output_path.display()
        )
    };

    Ok(CommandOutput::success(
        human,
        serde_json::json!({
            "operation": "merge",
            "changed": changed,
            "atomic": true,
            "templatePath": template_path,
            "outputPath": output_path,
            "dataSource": data_source,
            "force": parsed.force,
            "kind": editor.package().kind(),
            "result": result,
            "revision": editor.package().source_revision()
        }),
    ))
}

async fn read_template_data(argument: &str) -> UseResult<(serde_json::Value, String)> {
    let explicit_file = argument.strip_prefix('@');
    if explicit_file == Some("") {
        return Err(usage_error("--data @file requires a non-empty file path"));
    }
    let implicit_file = explicit_file.is_none()
        && Path::new(argument)
            .extension()
            .and_then(|value| value.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("json"))
        && tokio::fs::symlink_metadata(argument).await.is_ok();

    let (bytes, source) = if let Some(path) = explicit_file {
        (
            read_bounded_input(
                path,
                MAX_TEMPLATE_DATA_INPUT_BYTES,
                NativeInputKind::TemplateData,
            )
            .await?,
            path.to_string(),
        )
    } else if implicit_file {
        (
            read_bounded_input(
                argument,
                MAX_TEMPLATE_DATA_INPUT_BYTES,
                NativeInputKind::TemplateData,
            )
            .await?,
            argument.to_string(),
        )
    } else {
        let bytes = argument.as_bytes();
        if u64::try_from(bytes.len()).unwrap_or(u64::MAX) > MAX_TEMPLATE_DATA_INPUT_BYTES {
            return Err(input_too_large(
                NativeInputKind::TemplateData,
                "<inline>",
                MAX_TEMPLATE_DATA_INPUT_BYTES,
            ));
        }
        (bytes.to_vec(), "inline".to_string())
    };

    let data = serde_json::from_slice(&bytes).map_err(|error| {
        input_error(
            NativeInputKind::TemplateData,
            "invalid",
            &source,
            format!("Native Office template data is invalid JSON: {error}"),
        )
    })?;
    Ok((data, source))
}

fn reject_same_file(template: &Path, output: &Path) -> UseResult<()> {
    let template_path = canonical_or_absolute(template)?;
    let output_path = canonical_or_absolute(output)?;
    let same_path = template_path == output_path;
    let same_identity = existing_file_identity_matches(template, output)?;
    if same_path || same_identity {
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

fn canonical_or_absolute(path: &Path) -> UseResult<PathBuf> {
    if let Ok(canonical) = std::fs::canonicalize(path) {
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
        if let Ok(parent) = std::fs::canonicalize(parent) {
            return Ok(parent.join(name));
        }
    }
    Ok(absolute)
}

#[cfg(unix)]
fn existing_file_identity_matches(left: &Path, right: &Path) -> UseResult<bool> {
    use std::os::unix::fs::MetadataExt;

    let left = match std::fs::metadata(left) {
        Ok(metadata) => metadata,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(false),
        Err(error) => return Err(identity_error(left, error)),
    };
    let right = match std::fs::metadata(right) {
        Ok(metadata) => metadata,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(false),
        Err(error) => return Err(identity_error(right, error)),
    };
    Ok(left.dev() == right.dev() && left.ino() == right.ino())
}

#[cfg(not(unix))]
fn existing_file_identity_matches(_left: &Path, _right: &Path) -> UseResult<bool> {
    Ok(false)
}

#[cfg(unix)]
fn identity_error(path: &Path, error: std::io::Error) -> UseError {
    UseError::new(
        "use.office.template_path_invalid",
        format!(
            "Failed to inspect native Office merge path '{}': {error}",
            path.display()
        ),
    )
    .with_detail("path", path.display().to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[tokio::test]
    async fn data_accepts_inline_at_file_and_existing_json_paths() {
        let temp = tempfile::tempdir().unwrap();
        let path = temp.path().join("data.json");
        std::fs::write(&path, br#"{"name":"file"}"#).unwrap();

        let (inline, source) = read_template_data(r#"{"name":"inline"}"#).await.unwrap();
        assert_eq!(inline["name"], "inline");
        assert_eq!(source, "inline");

        let (explicit, _) = read_template_data(&format!("@{}", path.display()))
            .await
            .unwrap();
        assert_eq!(explicit["name"], "file");
        let (implicit, _) = read_template_data(path.to_str().unwrap()).await.unwrap();
        assert_eq!(implicit["name"], "file");
    }

    #[tokio::test]
    async fn data_files_are_size_bounded_and_reject_symbolic_links() {
        let oversized = tempfile::NamedTempFile::new().unwrap();
        oversized
            .as_file()
            .set_len(MAX_TEMPLATE_DATA_INPUT_BYTES + 1)
            .unwrap();
        let error = read_template_data(&format!("@{}", oversized.path().display()))
            .await
            .unwrap_err();
        assert_eq!(error.code, "use.office.template_data_input_too_large");

        #[cfg(unix)]
        {
            let directory = tempfile::tempdir().unwrap();
            let link = directory.path().join("data.json");
            std::os::unix::fs::symlink(oversized.path(), &link).unwrap();
            let error = read_template_data(&format!("@{}", link.display()))
                .await
                .unwrap_err();
            assert_eq!(error.code, "use.office.template_data_input_invalid");
        }
    }
}
