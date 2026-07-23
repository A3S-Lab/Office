use a3s_office::{NativeOfficeEditor, NativeOfficeInsertPosition};
use a3s_use_core::UseResult;

use super::{save_editor, usage_error};
use crate::CommandOutput;

use super::arguments::{AllowedOptions, ParsedArguments};

pub(super) async fn move_node(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::MOVE_NODE)?;
    if parsed.positionals.len() != 2 {
        return Err(usage_error("office native move requires <file> and <path>"));
    }
    let position = parse_insert_position(&parsed)?;
    let source = &parsed.positionals[0];
    let path = &parsed.positionals[1];
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let result_path = editor.move_node(path, parsed.target_parent.clone(), position.clone())?;
    let node = editor.snapshot()?.get(&result_path, 2)?;
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;

    Ok(CommandOutput::success(
        format!(
            "Moved {path} to {result_path} and saved '{}'.",
            output_path.display()
        ),
        serde_json::json!({
            "operation": "move",
            "changed": result_path != *path || position.is_some(),
            "path": path,
            "to": parsed.target_parent,
            "position": position,
            "resultPath": result_path,
            "node": node,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

pub(super) async fn copy_node(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::COPY_NODE)?;
    if parsed.positionals.len() != 2 {
        return Err(usage_error("office native copy requires <file> and <path>"));
    }
    let position = parse_insert_position(&parsed)?;
    let source = &parsed.positionals[0];
    let path = &parsed.positionals[1];
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let result_path = editor.copy_node(
        path,
        parsed.target_parent.clone(),
        position.clone(),
        parsed.name.clone(),
    )?;
    let node = editor.snapshot()?.get(&result_path, 2)?;
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;

    Ok(CommandOutput::success(
        format!(
            "Copied {path} to {result_path} and saved '{}'.",
            output_path.display()
        ),
        serde_json::json!({
            "operation": "copy",
            "changed": true,
            "path": path,
            "to": parsed.target_parent,
            "name": parsed.name,
            "position": position,
            "resultPath": result_path,
            "node": node,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

pub(super) async fn swap_nodes(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::MUTATE)?;
    if parsed.positionals.len() != 3 {
        return Err(usage_error(
            "office native swap requires <file>, <path>, and <with>",
        ));
    }
    let source = &parsed.positionals[0];
    let path = &parsed.positionals[1];
    let with = &parsed.positionals[2];
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let result = editor.swap_nodes(path, with)?;
    let snapshot = editor.snapshot()?;
    let first = snapshot.get(&result.first, 2)?;
    let second = snapshot.get(&result.second, 2)?;
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;

    Ok(CommandOutput::success(
        format!(
            "Swapped {path} with {with} and saved '{}'.",
            output_path.display()
        ),
        serde_json::json!({
            "operation": "swap",
            "changed": path != with,
            "path": path,
            "with": with,
            "result": result,
            "nodes": [first, second],
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

fn parse_insert_position(
    parsed: &ParsedArguments,
) -> UseResult<Option<NativeOfficeInsertPosition>> {
    let selected = [
        parsed.index.is_some(),
        parsed.before.is_some(),
        parsed.after.is_some(),
    ]
    .into_iter()
    .filter(|selected| *selected)
    .count();
    if selected > 1 {
        return Err(usage_error(
            "--index, --before, and --after are mutually exclusive",
        ));
    }
    Ok(if let Some(index) = parsed.index {
        Some(NativeOfficeInsertPosition::at_index(index))
    } else if let Some(path) = parsed.before.as_ref() {
        Some(NativeOfficeInsertPosition::before(path))
    } else {
        parsed.after.as_ref().map(NativeOfficeInsertPosition::after)
    })
}
