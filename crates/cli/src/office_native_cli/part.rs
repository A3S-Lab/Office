use a3s_office::{NativeOfficeEditor, NativeOfficePartType};
use a3s_use_core::UseResult;

use super::{save_editor, usage_error, AllowedOptions, CommandOutput, ParsedArguments};

pub(super) async fn add(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::ADD_PART)?;
    if parsed.positionals.len() != 2 {
        return Err(usage_error(
            "office native add-part requires <file> and <parent>",
        ));
    }
    let part_type = parse_part_type(
        parsed
            .node_type
            .as_deref()
            .ok_or_else(|| usage_error("office native add-part requires --type <part-type>"))?,
    )?;
    let source = &parsed.positionals[0];
    let parent = &parsed.positionals[1];
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let created = editor.add_part(parent, part_type)?;
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;

    Ok(CommandOutput::success(
        format!(
            "Created {:?} part {} as {} with relationship {} and saved '{}'.",
            created.part_type,
            created.path,
            created.part,
            created.relationship_id,
            output_path.display()
        ),
        serde_json::json!({
            "operation": "add-part",
            "changed": true,
            "createdPart": created,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

fn parse_part_type(value: &str) -> UseResult<NativeOfficePartType> {
    match value.to_ascii_lowercase().as_str() {
        "chart" => Ok(NativeOfficePartType::Chart),
        "header" => Ok(NativeOfficePartType::Header),
        "footer" => Ok(NativeOfficePartType::Footer),
        _ => Err(usage_error(format!(
            "native Office part type '{value}' is not supported; expected chart, header, or footer"
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn part_types_are_explicit() {
        assert_eq!(
            parse_part_type("chart").unwrap(),
            NativeOfficePartType::Chart
        );
        assert_eq!(
            parse_part_type("HEADER").unwrap(),
            NativeOfficePartType::Header
        );
        assert_eq!(
            parse_part_type("binary").unwrap_err().code,
            "use.cli.invalid_usage"
        );
    }
}
