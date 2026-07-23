use a3s_office::NativeOfficeEditor;
use a3s_use_core::UseResult;

use super::arguments::{AllowedOptions, ParsedArguments};
use super::{save_editor, usage_error};
use crate::CommandOutput;

pub(super) async fn recalculate(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::MUTATE)?;
    if parsed.positionals.len() != 1 {
        return Err(usage_error(
            "office native recalculate requires <file.xlsx>",
        ));
    }
    let source = &parsed.positionals[0];
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let calculation = editor.recalculate_spreadsheet_formulas()?;
    let changed = editor.is_dirty();
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;
    Ok(CommandOutput::success(
        format!(
            "Recalculated {} formula(s) and {} spill cell(s), then saved '{}'.",
            calculation.formula_count,
            calculation.spill_cell_count,
            output_path.display()
        ),
        serde_json::json!({
            "operation": "recalculate-spreadsheet-formulas",
            "changed": changed,
            "result": calculation,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}
