use a3s_office::{
    NativeOfficeEditor, NativeSpreadsheetSort, NativeSpreadsheetSortDirection,
    NativeSpreadsheetSortKey,
};
use a3s_use_core::UseResult;

use super::{save_editor, usage_error};
use crate::CommandOutput;

#[derive(Debug, Default, PartialEq, Eq)]
struct SortArguments {
    positionals: Vec<String>,
    keys: Vec<NativeSpreadsheetSortKey>,
    header: bool,
    case_sensitive: bool,
    output: Option<String>,
}

pub(super) async fn run(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = parse(args)?;
    if parsed.positionals.len() != 2 {
        return Err(usage_error(
            "office native sort requires <file.xlsx> and </Sheet|/Sheet/A1:D100>",
        ));
    }
    if parsed.keys.is_empty() {
        return Err(usage_error(
            "office native sort requires at least one ordered --key <A:XFD[:asc|desc]>",
        ));
    }

    let source = &parsed.positionals[0];
    let range = &parsed.positionals[1];
    let request = NativeSpreadsheetSort::new(parsed.keys)
        .with_header(parsed.header)
        .with_case_sensitive(parsed.case_sensitive);
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let result_path = editor.sort_spreadsheet_range(range, request.clone())?;
    let node = editor.snapshot()?.get(&result_path, 1)?;
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;

    Ok(CommandOutput::success(
        format!(
            "Sorted {range} by {} key(s) and saved '{}'.",
            request.keys.len(),
            output_path.display()
        ),
        serde_json::json!({
            "operation": "sort-spreadsheet-range",
            "changed": true,
            "range": range,
            "path": result_path,
            "sort": request,
            "node": node,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

fn parse(args: &[String]) -> UseResult<SortArguments> {
    let mut parsed = SortArguments::default();
    let mut index = 1;
    let mut header = None;
    let mut case_sensitive = None;
    while index < args.len() {
        match args[index].as_str() {
            "--json" => index += 1,
            "--" => {
                parsed.positionals.extend_from_slice(&args[index + 1..]);
                break;
            }
            "--key" => {
                let value = option_value(args, index, "--key")?;
                parsed.keys.push(parse_key(value)?);
                index += 2;
            }
            "--header" => {
                if header.is_some() {
                    return Err(usage_error("--header may be specified only once"));
                }
                header = Some(parse_bool(
                    "--header",
                    option_value(args, index, "--header")?,
                )?);
                index += 2;
            }
            "--case-sensitive" => {
                if case_sensitive.is_some() {
                    return Err(usage_error("--case-sensitive may be specified only once"));
                }
                case_sensitive = Some(parse_bool(
                    "--case-sensitive",
                    option_value(args, index, "--case-sensitive")?,
                )?);
                index += 2;
            }
            "--output" => {
                if parsed.output.is_some() {
                    return Err(usage_error("--output may be specified only once"));
                }
                parsed.output = Some(option_value(args, index, "--output")?.to_string());
                index += 2;
            }
            option if option.starts_with('-') => {
                return Err(usage_error(format!(
                    "unknown native Spreadsheet sort option '{option}'"
                )));
            }
            positional => {
                parsed.positionals.push(positional.to_string());
                index += 1;
            }
        }
    }
    parsed.header = header.unwrap_or(false);
    parsed.case_sensitive = case_sensitive.unwrap_or(false);
    Ok(parsed)
}

fn parse_key(value: &str) -> UseResult<NativeSpreadsheetSortKey> {
    let mut segments = value.split(':');
    let column = segments.next().unwrap_or_default().trim();
    let direction = segments.next().map(str::trim);
    if column.is_empty() || segments.next().is_some() {
        return Err(key_error(value));
    }
    let direction = match direction.map(str::to_ascii_lowercase).as_deref() {
        None | Some("asc" | "ascending") => NativeSpreadsheetSortDirection::Ascending,
        Some("desc" | "descending") => NativeSpreadsheetSortDirection::Descending,
        Some(_) => return Err(key_error(value)),
    };
    Ok(NativeSpreadsheetSortKey {
        column: column.to_string(),
        direction,
    })
}

fn parse_bool(option: &str, value: &str) -> UseResult<bool> {
    match value.to_ascii_lowercase().as_str() {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(usage_error(format!(
            "{option} requires true or false, received '{value}'"
        ))),
    }
}

fn option_value<'a>(args: &'a [String], index: usize, option: &str) -> UseResult<&'a str> {
    args.get(index + 1)
        .filter(|value| !value.starts_with("--"))
        .map(String::as_str)
        .ok_or_else(|| usage_error(format!("{option} requires a value")))
}

fn key_error(value: &str) -> a3s_use_core::UseError {
    usage_error(format!(
        "Spreadsheet sort key '{value}' must be <A:XFD[:asc|desc]>"
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_ordered_sort_keys_and_explicit_flags() {
        let parsed = parse(&[
            "sort".into(),
            "book.xlsx".into(),
            "/Sheet1/A1:D10".into(),
            "--key".into(),
            "B:desc".into(),
            "--key".into(),
            "C".into(),
            "--header".into(),
            "true".into(),
            "--case-sensitive".into(),
            "false".into(),
            "--output".into(),
            "sorted.xlsx".into(),
        ])
        .unwrap();
        assert_eq!(parsed.positionals, ["book.xlsx", "/Sheet1/A1:D10"]);
        assert_eq!(
            parsed.keys,
            [
                NativeSpreadsheetSortKey::descending("B"),
                NativeSpreadsheetSortKey::ascending("C")
            ]
        );
        assert!(parsed.header);
        assert!(!parsed.case_sensitive);
        assert_eq!(parsed.output.as_deref(), Some("sorted.xlsx"));
    }

    #[test]
    fn rejects_ambiguous_sort_options() {
        for args in [
            vec!["sort".into(), "--key".into(), "A:sideways".into()],
            vec!["sort".into(), "--header".into(), "yes".into()],
            vec!["sort".into(), "--unknown".into()],
        ] {
            assert_eq!(parse(&args).unwrap_err().code, "use.cli.invalid_usage");
        }
    }
}
