use std::path::Path;

use a3s_office::{
    NativeOfficeEditor, NativeSpreadsheetDelimitedFormat, NativeSpreadsheetDelimitedImport,
    MAX_NATIVE_SPREADSHEET_IMPORT_BYTES,
};
use a3s_use_core::UseResult;

use super::bounded_input::{input_error, read_bounded_input, read_bounded_stdin, NativeInputKind};
use super::{save_editor, usage_error};
use crate::CommandOutput;

#[derive(Debug, Default, PartialEq, Eq)]
struct ImportArguments {
    positionals: Vec<String>,
    source_file: Option<String>,
    stdin: bool,
    format: Option<NativeSpreadsheetDelimitedFormat>,
    header: bool,
    start_cell: String,
    output: Option<String>,
}

pub(super) async fn run(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = parse(args)?;
    if !(2..=3).contains(&parsed.positionals.len()) {
        return Err(usage_error(
            "office native import requires <file.xlsx>, <sheet>, and one source file or --stdin",
        ));
    }
    let positional_source = parsed.positionals.get(2).map(String::as_str);
    let sources = usize::from(positional_source.is_some())
        + usize::from(parsed.source_file.is_some())
        + usize::from(parsed.stdin);
    if sources != 1 {
        return Err(usage_error(
            "office native import requires exactly one positional source file, --file <source>, or --stdin",
        ));
    }
    let source_file = parsed.source_file.as_deref().or(positional_source);
    let limit = u64::try_from(MAX_NATIVE_SPREADSHEET_IMPORT_BYTES).unwrap_or(u64::MAX);
    let bytes = if parsed.stdin {
        read_bounded_stdin(limit, NativeInputKind::SpreadsheetImport).await?
    } else {
        read_bounded_input(
            source_file.ok_or_else(|| usage_error("Spreadsheet import source is missing"))?,
            limit,
            NativeInputKind::SpreadsheetImport,
        )
        .await?
    };
    let input_label = source_file.unwrap_or("<stdin>");
    let content = String::from_utf8(bytes).map_err(|error| {
        input_error(
            NativeInputKind::SpreadsheetImport,
            "invalid",
            input_label,
            format!("Spreadsheet import input '{input_label}' is not valid UTF-8: {error}"),
        )
    })?;
    let format = parsed.format.unwrap_or_else(|| infer_format(source_file));
    let import = NativeSpreadsheetDelimitedImport::new(content, format)
        .with_header(parsed.header)
        .with_start_cell(parsed.start_cell);

    let source = &parsed.positionals[0];
    let sheet = &parsed.positionals[1];
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let result = editor.import_spreadsheet_delimited(sheet, import)?;
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;
    let human = if result.changed {
        format!(
            "Imported {} row(s) x {} column(s) into {} at {} and saved '{}'.",
            result.row_count,
            result.column_count,
            result.sheet,
            result.start_cell,
            output_path.display()
        )
    } else {
        format!(
            "No delimited rows were present; saved '{}'.",
            output_path.display()
        )
    };
    Ok(CommandOutput::success(
        human,
        serde_json::json!({
            "operation": "import-spreadsheet-delimited",
            "changed": result.changed,
            "source": if parsed.stdin { serde_json::Value::String("stdin".into()) } else { serde_json::Value::String(input_label.into()) },
            "result": result,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

fn parse(args: &[String]) -> UseResult<ImportArguments> {
    let mut parsed = ImportArguments {
        start_cell: "A1".into(),
        ..ImportArguments::default()
    };
    let mut index = 1;
    let mut header = false;
    let mut start_cell_seen = false;
    while index < args.len() {
        match args[index].as_str() {
            "--json" => index += 1,
            "--" => {
                parsed.positionals.extend_from_slice(&args[index + 1..]);
                break;
            }
            "--file" => {
                set_option(&mut parsed.source_file, args, index, "--file")?;
                index += 2;
            }
            "--stdin" => {
                if parsed.stdin {
                    return Err(usage_error("--stdin may be specified only once"));
                }
                parsed.stdin = true;
                index += 1;
            }
            "--format" => {
                if parsed.format.is_some() {
                    return Err(usage_error("--format may be specified only once"));
                }
                let value = option_value(args, index, "--format")?;
                parsed.format = Some(match value.to_ascii_lowercase().as_str() {
                    "csv" => NativeSpreadsheetDelimitedFormat::Csv,
                    "tsv" => NativeSpreadsheetDelimitedFormat::Tsv,
                    _ => return Err(usage_error("--format requires csv or tsv")),
                });
                index += 2;
            }
            "--header" => {
                if header {
                    return Err(usage_error("--header may be specified only once"));
                }
                header = true;
                parsed.header = true;
                index += 1;
            }
            "--start-cell" => {
                if start_cell_seen {
                    return Err(usage_error("--start-cell may be specified only once"));
                }
                parsed.start_cell = option_value(args, index, "--start-cell")?.into();
                start_cell_seen = true;
                index += 2;
            }
            "--output" => {
                set_option(&mut parsed.output, args, index, "--output")?;
                index += 2;
            }
            option if option.starts_with('-') => {
                return Err(usage_error(format!(
                    "unknown native Spreadsheet import option '{option}'"
                )));
            }
            positional => {
                parsed.positionals.push(positional.into());
                index += 1;
            }
        }
    }
    Ok(parsed)
}

fn infer_format(source: Option<&str>) -> NativeSpreadsheetDelimitedFormat {
    let extension = source
        .and_then(|source| Path::new(source).extension())
        .and_then(|extension| extension.to_str())
        .map(str::to_ascii_lowercase);
    if matches!(extension.as_deref(), Some("tsv" | "tab")) {
        NativeSpreadsheetDelimitedFormat::Tsv
    } else {
        NativeSpreadsheetDelimitedFormat::Csv
    }
}

fn set_option(
    target: &mut Option<String>,
    args: &[String],
    index: usize,
    option: &str,
) -> UseResult<()> {
    if target.is_some() {
        return Err(usage_error(format!("{option} may be specified only once")));
    }
    *target = Some(option_value(args, index, option)?.into());
    Ok(())
}

fn option_value<'a>(args: &'a [String], index: usize, option: &str) -> UseResult<&'a str> {
    args.get(index + 1)
        .filter(|value| !value.starts_with("--"))
        .map(String::as_str)
        .ok_or_else(|| usage_error(format!("{option} requires a value")))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_file_stdin_format_header_and_start_cell() {
        let parsed = parse(&[
            "import".into(),
            "book.xlsx".into(),
            "/Data".into(),
            "--file".into(),
            "source.tab".into(),
            "--format".into(),
            "tsv".into(),
            "--header".into(),
            "--start-cell".into(),
            "B2".into(),
            "--output".into(),
            "copy.xlsx".into(),
        ])
        .unwrap();
        assert_eq!(parsed.positionals, ["book.xlsx", "/Data"]);
        assert_eq!(parsed.source_file.as_deref(), Some("source.tab"));
        assert_eq!(parsed.format, Some(NativeSpreadsheetDelimitedFormat::Tsv));
        assert!(parsed.header);
        assert_eq!(parsed.start_cell, "B2");
        assert_eq!(parsed.output.as_deref(), Some("copy.xlsx"));
        assert_eq!(
            infer_format(Some("data.TAB")),
            NativeSpreadsheetDelimitedFormat::Tsv
        );
    }

    #[test]
    fn rejects_duplicate_and_unknown_options() {
        for args in [
            vec!["import".into(), "--stdin".into(), "--stdin".into()],
            vec!["import".into(), "--format".into(), "json".into()],
            vec!["import".into(), "--unknown".into()],
        ] {
            assert_eq!(parse(&args).unwrap_err().code, "use.cli.invalid_usage");
        }
    }
}
