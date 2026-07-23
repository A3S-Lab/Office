use a3s_office::{
    DocumentKind, DocumentNode, NativeOfficeCommentPosition, NativeOfficeCommentUpdate,
    NativeOfficeDocument, NativeOfficeEditor, NativeOfficeMutation, NativeOfficeTextReplacement,
    OfficeNodeType, SpreadsheetCellValue,
};
use a3s_use_core::{UseError, UseResult};

use crate::CommandOutput;

mod add;
mod arguments;
mod arrange;
mod bounded_input;
mod conditional_formatting;
mod data_validation;
mod format;
mod merge;
mod named_range;
mod part;
mod raw;
mod replay;
mod spreadsheet_filter;
mod spreadsheet_formula;
mod spreadsheet_import;
mod spreadsheet_sort;
mod spreadsheet_table;
mod view;
mod watch;

use arguments::{parse_boolean_option, AllowedOptions, ParsedArguments};
use format::{parse_cell_format, parse_hyperlink, parse_text_format};

const MAX_BATCH_INPUT_BYTES: u64 = 8 * 1024 * 1024;
const MAX_IMAGE_INPUT_BYTES: u64 = 64 * 1024 * 1024;

const HELP: &str = concat!(
    "a3s-office — dependency-free OOXML operations\n\n",
    "usage:\n",
    "  a3s-office get <file> [path] [--depth <n>] [--json]\n",
    "  a3s-office query <file> <selector> [--json]\n",
    "  a3s-office view <file> text|annotated|outline|stats|issues|html|svg|screenshot [--type <filter>] [--limit <n>] [--output <file>] [--timeout-ms <ms>] [--json]\n",
    "  a3s-office watch <file> [--port <0-65535>] [--poll-ms <50-10000>] [--timeout-ms <ms>] [--json]\n",
    "  a3s-office raw <file> <part> [--output <xml-file>] [--json]\n",
    "  a3s-office raw-set <file> <part> --input <xml-file> [--output <file>] [--json]\n",
    "  a3s-office dump <file> [path] [--output <batch.json>] [--json]\n",
    "  a3s-office merge <template> <output> --data <json|@file.json> [--force] [--json]\n",
    "  a3s-office validate <file> [--json]\n",
    "  a3s-office create <file.docx|file.xlsx|file.pptx> [--json]\n",
    "  a3s-office add <file> <parent> --type paragraph|table|row|cell|sheet|slide|shape|picture|hyperlink|comment|data-validation|conditional-format|named-range|auto-filter [--author <name>] [--initials <value>] [--x-emu <i32> --y-emu <i32>] [--url <http|https|mailto>|--location <internal>] [--display <text>] [--tooltip <text>] [--input <image>] [--name <name>] [--alt <text>] [--width <pixels>] [--height <pixels>] [--rows <n>] [--columns <n>] [--text <value>] [--output <file>] [--json]\n",
    "  a3s-office add <file.xlsx> <sheet> --type data-validation --validation-type list|whole|decimal|date|time|text-length|custom --range <A1-range> [--range <A1-range>] --formula1 <value> [--operator <comparison>] [--formula2 <value>] [--allow-blank <true|false>] [--show-input <true|false>] [--show-error <true|false>] [--prompt-title <text>] [--prompt <text>] [--error-title <text>] [--error-message <text>] [--error-style stop|warning|information] [--in-cell-dropdown <true|false>] [--output <file>] [--json]\n",
    "  a3s-office add <file.xlsx> <sheet> --type conditional-format --rule-type cell-is|formula|contains-text|not-contains-text|begins-with|ends-with|top|bottom|above-average|below-average|duplicate-values|unique-values|contains-blanks|not-contains-blanks|contains-errors|not-contains-errors|time-period|data-bar|color-scale|icon-set --range <A1-range> [rule-specific options] [--stop-if-true <true|false>] [--output <file>] [--json]\n",
    "  a3s-office add <file.xlsx> </|sheet> --type named-range --name <name> --ref <expression> [--scope workbook|worksheet:workbook|<sheet>] [--comment <text>] [--volatile <true|false>] [--output <file>] [--json]\n",
    "  a3s-office add <file.xlsx> <sheet> --type auto-filter --range <A1-range> [--filter <strict-json-column> ...] [--output <file>] [--json]\n",
    "  a3s-office add <file.xlsx> <sheet> --type table --name <name> --range <A1-range> --table-column <name> [--table-column <name> ...] [--filter <strict-json-column> ...] [--display-name <name>] [--header-row <true|false>] [--totals-row <true|false>] [--style none|light:<1-21>|medium:<1-28>|dark:<1-11>] [--show-first-column <true|false>] [--show-last-column <true|false>] [--show-row-stripes <true|false>] [--show-column-stripes <true|false>] [--output <file>] [--json]\n",
    "  a3s-office add-part <file> <parent> --type chart|header|footer [--output <file>] [--json]\n",
    "  a3s-office set <file> <path> [--find <text> --replace <text> [--regex]|--text <value>|--number <value>|--boolean <true|false>|--formula <expression>|--width-emu <n>] [--author <name>] [--initials <value>] [--x-emu <i32> --y-emu <i32>] [--bold <true|false>] [--italic <true|false>] [--underline <none|single|double>] [--script <baseline|superscript|subscript>] [--strikethrough <true|false>] [--double-strikethrough <true|false>] [--text-case <none|small-caps|all-caps>] [--highlight <color>] [--language <BCP-47>] [--font-family <name>] [--font-size <points>] [--text-color <RRGGBB>] [--align <left|center|right|justify>] [--number-format <code>] [--fill <none|RRGGBB>] [--border-all <style>] [--border-color <RRGGBB>] [--border-left|--border-right|--border-top|--border-bottom <style>] [--border-left-color|--border-right-color|--border-top-color|--border-bottom-color <RRGGBB>] [--border-diagonal <style>] [--border-diagonal-color <RRGGBB>] [--border-diagonal-up <true|false>] [--border-diagonal-down <true|false>] [--vertical-align <top|center|bottom|justify|distributed>] [--wrap-text <true|false>] [--text-rotation <0..180|255>] [--indent <0..255>] [--shrink-to-fit <true|false>] [--reading-order <context|ltr|rtl>] [--merge-cells <true|false>] [--url <http|https|mailto>|--location <internal>] [--display <text>] [--tooltip <text>] [--output <file>] [--json]\n",
    "  a3s-office set <file.xlsx> <sheet/dataValidation[N]> [data-validation options from add; unspecified fields are preserved, use none or an empty value to clear optional text/formula2] [--output <file>] [--json]\n",
    "  a3s-office set <file.xlsx> <sheet/cf[N]> [conditional-format options from add; unspecified fields are preserved] [--output <file>] [--json]\n",
    "  a3s-office set <file.xlsx> <namedrange-selector> [--name <name>] [--ref <expression>] [--scope workbook|worksheet:workbook|<sheet>] [--comment <text|none>] [--volatile <true|false>] [--output <file>] [--json]\n",
    "  a3s-office set <file.xlsx> <sheet/autofilter> [--range <A1-range>] [--filter <strict-json-column> ...|--clear-filters] [--output <file>] [--json]\n",
    "  a3s-office set <file.xlsx> <sheet/table[N]> [--name <name>] [--display-name <name|none>] [--range <A1-range>] [--table-column <name> ...] [--filter <strict-json-column> ...|--clear-filters] [--header-row <true|false>] [--totals-row <true|false>] [--style none|light:<1-21>|medium:<1-28>|dark:<1-11>] [--show-first-column <true|false>] [--show-last-column <true|false>] [--show-row-stripes <true|false>] [--show-column-stripes <true|false>] [--output <file>] [--json]\n",
    "  a3s-office sort <file.xlsx> </Sheet|/Sheet/A1:D100> --key <A:XFD[:asc|desc]> [--key <column[:direction]> ...] [--header <true|false>] [--case-sensitive <true|false>] [--output <file>] [--json]\n",
    "  a3s-office import <file.xlsx> <sheet> [source.csv|source.tsv] [--file <source>] [--stdin] [--format csv|tsv] [--header] [--start-cell <A1>] [--output <file>] [--json]\n",
    "  a3s-office recalculate <file.xlsx> [--output <file>] [--json]\n",
    "  a3s-office remove <file> <path> [--output <file>] [--json]\n",
    "  a3s-office move <file> <path> [--to <parent>] [--index <zero-based>|--before <path>|--after <path>] [--output <file>] [--json]\n",
    "  a3s-office copy <file> <path> [--to <parent>] [--name <worksheet-name>] [--index <zero-based>|--before <path>|--after <path>] [--output <file>] [--json]\n",
    "  a3s-office swap <file> <path> <with> [--output <file>] [--json]\n",
    "  a3s-office insert-rows|delete-rows <file> <sheet> <start> [--count <n>] [--output <file>] [--json]\n",
    "  a3s-office insert-columns|delete-columns <file> <sheet> <start> [--count <n>] [--output <file>] [--json]\n",
    "  a3s-office rename-sheet <file> <sheet> <new-name> [--output <file>] [--json]\n",
    "  a3s-office move-sheet <file> <sheet> <one-based-position> [--output <file>] [--json]\n",
    "  a3s-office copy-sheet <file> <sheet> <new-name> [--position <one-based-position>] [--output <file>] [--json]\n",
    "  a3s-office batch <file> --input <batch.json> [--output <file>] [--json]"
);

pub async fn run(args: &[String]) -> UseResult<CommandOutput> {
    match args.first().map(String::as_str) {
        None | Some("help" | "--help" | "-h") => Ok(help()),
        Some("get") => get(args).await,
        Some("query") => query(args).await,
        Some("view") => view::run(args).await,
        Some("watch") => watch::run(args).await,
        Some("raw") => raw::inspect(args).await,
        Some("raw-set") => raw::replace(args).await,
        Some("dump") => replay::dump(args).await,
        Some("merge") => merge::run(args).await,
        Some("validate") => validate(args).await,
        Some("create") => create(args).await,
        Some("add") => add::run(args).await,
        Some("add-part") => part::add(args).await,
        Some("set") => set(args).await,
        Some("sort") => spreadsheet_sort::run(args).await,
        Some("import") => spreadsheet_import::run(args).await,
        Some("recalculate") => spreadsheet_formula::recalculate(args).await,
        Some("remove") => remove(args).await,
        Some("move") => arrange::move_node(args).await,
        Some("copy") => arrange::copy_node(args).await,
        Some("swap") => arrange::swap_nodes(args).await,
        Some("insert-rows") => edit_structure(args, StructureOperation::InsertRows).await,
        Some("delete-rows") => edit_structure(args, StructureOperation::DeleteRows).await,
        Some("insert-columns") => edit_structure(args, StructureOperation::InsertColumns).await,
        Some("delete-columns") => edit_structure(args, StructureOperation::DeleteColumns).await,
        Some("rename-sheet") => rename_sheet(args).await,
        Some("move-sheet") => move_sheet(args).await,
        Some("copy-sheet") => copy_sheet(args).await,
        Some("batch") => batch(args).await,
        Some(command) => Err(usage_error(format!(
            "unknown native Office command '{command}'"
        ))),
    }
}

fn help() -> CommandOutput {
    CommandOutput::success(
        HELP,
        serde_json::json!({
            "commands": [
                "get", "query", "view", "watch", "raw", "raw-set", "dump", "merge", "validate", "create", "add", "add-part", "set", "sort", "import", "recalculate", "remove", "move", "copy", "swap",
                "insert-rows", "delete-rows", "insert-columns", "delete-columns",
                "rename-sheet", "move-sheet", "copy-sheet", "batch"
            ],
            "formats": ["docx", "xlsx", "pptx"],
            "textReplacementModes": ["literal", "regex"],
            "runtimeDependencies": [],
            "atomicBatch": true
        }),
    )
}

async fn get(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::GET)?;
    if !(1..=2).contains(&parsed.positionals.len()) {
        return Err(usage_error(
            "office native get requires <file> and an optional [path]",
        ));
    }
    let depth = parsed.depth.unwrap_or(1);
    if depth > 64 {
        return Err(usage_error("--depth cannot exceed 64"));
    }
    let document = NativeOfficeDocument::open(&parsed.positionals[0]).await?;
    let path = parsed.positionals.get(1).map_or("/", String::as_str);
    let node = document.get(path, depth)?;
    let human = format_node(&node, 0);
    Ok(CommandOutput::success(
        human,
        serde_json::json!({ "node": node }),
    ))
}

async fn query(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::NONE)?;
    if parsed.positionals.len() != 2 {
        return Err(usage_error(
            "office native query requires <file> and <selector>",
        ));
    }
    let document = NativeOfficeDocument::open(&parsed.positionals[0]).await?;
    let results = document.query(&parsed.positionals[1])?;
    let mut human = format!("Matches: {}", results.len());
    for node in &results {
        human.push_str(&format!("\n  {}: {}", node.path, single_line(&node.text)));
    }
    Ok(CommandOutput::success(
        human,
        serde_json::json!({
            "matches": results.len(),
            "results": results
        }),
    ))
}

async fn validate(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::NONE)?;
    if parsed.positionals.len() != 1 {
        return Err(usage_error("office native validate requires <file>"));
    }
    let document = NativeOfficeDocument::open(&parsed.positionals[0]).await?;
    Ok(CommandOutput::success(
        format!(
            "Valid native {:?} document: {}",
            document.kind(),
            parsed.positionals[0]
        ),
        serde_json::json!({
            "valid": true,
            "kind": document.kind(),
            "path": parsed.positionals[0]
        }),
    ))
}

async fn create(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::NONE)?;
    if parsed.positionals.len() != 1 {
        return Err(usage_error("office native create requires <file>"));
    }
    let editor = NativeOfficeEditor::create(&parsed.positionals[0]).await?;
    let output_path = editor.package().path().to_path_buf();
    Ok(CommandOutput::success(
        format!(
            "Created blank native Office document '{}'.",
            output_path.display()
        ),
        serde_json::json!({
            "operation": "create",
            "created": true,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "revision": editor.package().source_revision()
        }),
    ))
}

async fn set(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::SET)?;
    if parsed.positionals.len() != 2 {
        return Err(usage_error("office native set requires <file> and <path>"));
    }
    if parsed.find.is_some() || parsed.replacement.is_some() || parsed.regex {
        return replace_text(parsed).await;
    }
    let value_count = [
        parsed.text.is_some(),
        parsed.number.is_some(),
        parsed.boolean.is_some(),
        parsed.formula.is_some(),
        parsed.width_emu.is_some(),
    ]
    .into_iter()
    .filter(|present| *present)
    .count();
    let format = parse_text_format(&parsed)?;
    let cell_format = parse_cell_format(&parsed)?;
    let hyperlink = parse_hyperlink(&parsed, parsed.display.as_deref())?;
    let merge_cells = parsed
        .merge_cells
        .as_deref()
        .map(parse_merge_cells)
        .transpose()?;
    if value_count > 1 {
        return Err(usage_error(
            "office native set accepts at most one of --text, --number, --boolean, --formula, or --width-emu",
        ));
    }
    let typed_value = if let Some(value) = parsed.number.as_ref() {
        Some(SpreadsheetCellValue::Number {
            value: value.clone(),
        })
    } else if let Some(value) = parsed.boolean.as_deref() {
        Some(SpreadsheetCellValue::Boolean {
            value: parse_boolean_option(value)?,
        })
    } else {
        parsed
            .formula
            .as_ref()
            .map(|expression| SpreadsheetCellValue::Formula {
                expression: expression.clone(),
            })
    };
    let source = &parsed.positionals[0];
    let path = &parsed.positionals[1];
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let snapshot = editor.snapshot()?;
    let conditional_format_path = conditional_formatting::canonical_path(path);
    let validation_path = data_validation::canonical_path(path);
    let auto_filter_path = spreadsheet_filter::canonical_path(path);
    let named_range_path = named_range::is_path(path);
    let target_path = conditional_format_path
        .as_deref()
        .or(validation_path.as_deref())
        .or(auto_filter_path.as_deref())
        .unwrap_or(path);
    let target_node = snapshot.get(target_path, 0).ok();
    if (conditional_format_path.is_some()
        || validation_path.is_some()
        || auto_filter_path.is_some()
        || named_range_path)
        && target_node.is_none()
    {
        snapshot.get(target_path, 0)?;
    }
    let is_comment = target_node
        .as_ref()
        .is_some_and(|node| node.node_type == OfficeNodeType::Comment);
    let is_data_validation = target_node
        .as_ref()
        .is_some_and(|node| node.node_type == OfficeNodeType::DataValidation);
    let is_conditional_format = target_node
        .as_ref()
        .is_some_and(|node| node.node_type == OfficeNodeType::ConditionalFormatting);
    let is_named_range = target_node
        .as_ref()
        .is_some_and(|node| node.node_type == OfficeNodeType::NamedRange);
    let is_spreadsheet_auto_filter = target_node
        .as_ref()
        .is_some_and(|node| node.node_type == OfficeNodeType::AutoFilter);
    let is_spreadsheet_table = editor.package().kind() == DocumentKind::Spreadsheet
        && target_node
            .as_ref()
            .is_some_and(|node| node.node_type == OfficeNodeType::Table);
    let has_data_validation_options = parsed.has_data_validation_options();
    let has_conditional_format_options = parsed.has_conditional_format_options();
    let has_named_range_options = parsed.has_named_range_options();
    let has_spreadsheet_table_options = parsed.has_spreadsheet_table_options();
    let has_spreadsheet_filter_options = parsed.has_spreadsheet_filter_options();
    let has_comment_options = parsed.author.is_some()
        || parsed.initials.is_some()
        || parsed.x_emu.is_some()
        || parsed.y_emu.is_some();
    let mut result_path = path.clone();
    let operation = if is_spreadsheet_auto_filter
        || auto_filter_path.is_some()
        || (parsed.has_spreadsheet_filter_specific_options() && !is_spreadsheet_table)
    {
        if !is_spreadsheet_auto_filter {
            return Err(usage_error(
                "Spreadsheet filter options require an existing /Sheet/autofilter path or Spreadsheet table",
            ));
        }
        if !has_spreadsheet_filter_options {
            return Err(usage_error(
                "native Spreadsheet AutoFilter set requires --range, --filter, or --clear-filters",
            ));
        }
        if has_spreadsheet_filter_conflicts(&parsed) {
            return Err(usage_error(
                "native Spreadsheet AutoFilter set cannot be combined with content, formatting, hyperlink, merge, comment, validation, conditional-format, named-range, or table-only options",
            ));
        }
        let node = target_node.as_ref().ok_or_else(|| {
            usage_error("native Spreadsheet AutoFilter set requires an existing AutoFilter node")
        })?;
        let filter_node = snapshot.get(&node.path, 2)?;
        let filter = spreadsheet_filter::merge_existing(&filter_node, &parsed)?;
        result_path = editor.set_spreadsheet_auto_filter(&node.path, filter)?;
        "set-spreadsheet-auto-filter"
    } else if is_spreadsheet_table || parsed.has_spreadsheet_table_specific_options() {
        if !is_spreadsheet_table {
            return Err(usage_error(
                "Spreadsheet table options require an existing /Sheet/table[N] path",
            ));
        }
        if !has_spreadsheet_table_options {
            return Err(usage_error(
                "native Spreadsheet table set requires at least one table option",
            ));
        }
        if value_count > 0
            || format.is_some()
            || cell_format.is_some()
            || hyperlink.is_some()
            || merge_cells.is_some()
            || has_comment_options
            || parsed.has_data_validation_specific_options()
            || parsed.validation_operator.is_some()
            || parsed.validation_formula1.is_some()
            || parsed.validation_formula2.is_some()
            || has_conditional_format_options
            || parsed.named_range_ref.is_some()
            || parsed.named_range_scope.is_some()
            || parsed.named_range_comment.is_some()
            || parsed.named_range_volatile.is_some()
        {
            return Err(usage_error(
                "native Spreadsheet table set cannot be combined with content, formatting, hyperlink, merge, comment, validation, conditional-format, or named-range options",
            ));
        }
        let node = target_node.as_ref().ok_or_else(|| {
            usage_error("native Spreadsheet table set requires an existing table node")
        })?;
        let table_node = snapshot.get(&node.path, 2)?;
        let table = spreadsheet_table::merge_existing(&table_node, &parsed)?;
        result_path = editor.set_spreadsheet_table(&node.path, table)?;
        "set-spreadsheet-table"
    } else if is_named_range || named_range_path || has_named_range_options {
        if !is_named_range {
            return Err(usage_error(
                "named-range options require an existing /namedrange[NAME] path",
            ));
        }
        if !has_named_range_options {
            return Err(usage_error(
                "native named-range set requires at least one of --name, --ref, --scope, --comment, or --volatile",
            ));
        }
        if value_count > 0
            || format.is_some()
            || cell_format.is_some()
            || hyperlink.is_some()
            || merge_cells.is_some()
            || has_comment_options
            || has_data_validation_options
            || has_conditional_format_options
            || parsed.has_spreadsheet_filter_specific_options()
        {
            return Err(usage_error(
                "native named-range set cannot be combined with content, formatting, hyperlink, merge, width, comment-node, or data-validation options",
            ));
        }
        let node = target_node.as_ref().ok_or_else(|| {
            usage_error("native named-range set requires an existing named-range node")
        })?;
        let named_range = named_range::merge_existing(node, &parsed)?;
        result_path = editor.set_named_range(&node.path, named_range)?;
        "set-named-range"
    } else if is_conditional_format
        || conditional_format_path.is_some()
        || has_conditional_format_options
    {
        if !is_conditional_format {
            return Err(usage_error(
                "conditional-format options require an existing /Sheet/cf[N] path",
            ));
        }
        if !parsed.has_conditional_format_update_options() {
            return Err(usage_error(
                "native conditional-format set requires at least one conditional-format option",
            ));
        }
        if has_conditional_format_conflicts(&parsed) {
            return Err(usage_error(
                "native conditional-format set cannot be combined with cell values, regular text/cell formatting, hyperlinks, merges, comments, data-validation-only options, or named-range options",
            ));
        }
        let node = target_node.as_ref().ok_or_else(|| {
            usage_error(
                "native conditional-format set requires an existing conditional-format node",
            )
        })?;
        let conditional_format = conditional_formatting::merge_existing(node, &parsed)?;
        result_path = editor.set_conditional_format(&node.path, conditional_format)?;
        "set-conditional-format"
    } else if is_data_validation || validation_path.is_some() || has_data_validation_options {
        if !is_data_validation {
            return Err(usage_error(
                "data-validation options require an existing /Sheet/dataValidation[N] path",
            ));
        }
        if !has_data_validation_options {
            return Err(usage_error(
                "native data-validation set requires at least one data-validation option",
            ));
        }
        if value_count > 0
            || format.is_some()
            || cell_format.is_some()
            || hyperlink.is_some()
            || merge_cells.is_some()
            || has_comment_options
            || parsed.has_spreadsheet_filter_specific_options()
        {
            return Err(usage_error(
                "native data-validation set cannot be combined with content, formatting, hyperlink, merge, width, or comment options",
            ));
        }
        let node = target_node.as_ref().ok_or_else(|| {
            usage_error("native data-validation set requires an existing validation node")
        })?;
        let validation = data_validation::merge_existing(node, &parsed)?;
        result_path = editor.set_data_validation(&node.path, validation)?;
        "set-data-validation"
    } else if is_comment || has_comment_options {
        if !is_comment {
            return Err(usage_error(
                "--author, --initials, --x-emu, and --y-emu require an existing comment path",
            ));
        }
        if parsed.number.is_some()
            || parsed.boolean.is_some()
            || parsed.formula.is_some()
            || parsed.width_emu.is_some()
            || format.is_some()
            || cell_format.is_some()
            || hyperlink.is_some()
            || merge_cells.is_some()
        {
            return Err(usage_error(
                "native comment set accepts only --text, --author, --initials, --x-emu, and --y-emu",
            ));
        }
        let position = match (parsed.x_emu, parsed.y_emu) {
            (Some(x_emu), Some(y_emu)) => Some(NativeOfficeCommentPosition::new(x_emu, y_emu)),
            (None, None) => None,
            _ => {
                return Err(usage_error(
                    "native Presentation comment coordinates require both --x-emu and --y-emu",
                ))
            }
        };
        editor.set_comment(
            path,
            NativeOfficeCommentUpdate {
                author: parsed.author.clone(),
                text: parsed.text.clone(),
                initials: parsed.initials.clone(),
                position,
            },
        )?;
        "set-comment"
    } else if let Some(width_emu) = parsed.width_emu {
        if format.is_some() || cell_format.is_some() || hyperlink.is_some() || merge_cells.is_some()
        {
            return Err(usage_error(
                "--width-emu cannot be combined with formatting or hyperlink options",
            ));
        }
        editor.set_table_column_width(path, width_emu)?;
        "set-table-column-width"
    } else {
        if value_count == 0
            && format.is_none()
            && cell_format.is_none()
            && hyperlink.is_none()
            && merge_cells.is_none()
        {
            return Err(usage_error(
                "office native set requires content, width, typed formatting, a merged-cell state, a hyperlink target, or comment properties",
            ));
        }
        let mut mutations = Vec::with_capacity(5);
        if let Some(value) = typed_value {
            mutations.push(NativeOfficeMutation::SetCellValue {
                path: path.clone(),
                value,
            });
        } else if let Some(text) = &parsed.text {
            mutations.push(NativeOfficeMutation::SetText {
                path: path.clone(),
                text: text.clone(),
            });
        }
        if let Some(format) = format {
            mutations.push(NativeOfficeMutation::SetTextFormat {
                path: path.clone(),
                format,
            });
        }
        if let Some(format) = cell_format {
            mutations.push(NativeOfficeMutation::SetCellFormat {
                path: path.clone(),
                format,
            });
        }
        if let Some(hyperlink) = hyperlink {
            mutations.push(NativeOfficeMutation::SetHyperlink {
                path: path.clone(),
                hyperlink,
            });
        }
        if let Some(merge) = merge_cells {
            mutations.push(if merge {
                NativeOfficeMutation::MergeCells { path: path.clone() }
            } else {
                NativeOfficeMutation::UnmergeCells { path: path.clone() }
            });
        }
        editor.apply_batch(&mutations)?;
        let has_merge = mutations.iter().find_map(|mutation| match mutation {
            NativeOfficeMutation::MergeCells { .. } => Some(true),
            NativeOfficeMutation::UnmergeCells { .. } => Some(false),
            _ => None,
        });
        let has_format = mutations
            .iter()
            .any(|mutation| matches!(mutation, NativeOfficeMutation::SetTextFormat { .. }));
        let has_hyperlink = mutations
            .iter()
            .any(|mutation| matches!(mutation, NativeOfficeMutation::SetHyperlink { .. }));
        let has_cell_format = mutations
            .iter()
            .any(|mutation| matches!(mutation, NativeOfficeMutation::SetCellFormat { .. }));
        let has_content = mutations.iter().any(|mutation| {
            matches!(
                mutation,
                NativeOfficeMutation::SetText { .. } | NativeOfficeMutation::SetCellValue { .. }
            )
        });
        let base_operation = if has_cell_format {
            match (has_content, has_format, has_hyperlink) {
                (false, false, false) => "set-cell-format",
                (false, true, false) => "set-text-and-cell-format",
                (true, false, false) => "set-content-and-cell-format",
                (true, true, false) => "set-content-and-text-and-cell-format",
                (false, false, true) => "set-cell-format-and-hyperlink",
                (false, true, true) => "set-text-and-cell-format-and-hyperlink",
                (true, false, true) => "set-content-and-cell-format-and-hyperlink",
                (true, true, true) => "set-content-and-text-and-cell-format-and-hyperlink",
            }
        } else {
            match (has_content, has_format, has_hyperlink, mutations.first()) {
                (false, false, true, _) => "set-hyperlink",
                (false, true, false, _) => "set-text-format",
                (false, true, true, _) => "set-text-format-and-hyperlink",
                (true, false, true, _) => "set-content-and-hyperlink",
                (true, true, true, _) => "set-content-format-and-hyperlink",
                (true, true, false, _) => "set-content-and-text-format",
                (true, false, false, Some(NativeOfficeMutation::SetCellValue { .. })) => {
                    "set-cell-value"
                }
                _ => "set-text",
            }
        };
        match (has_merge, mutations.len()) {
            (Some(true), 1) => "merge-cells",
            (Some(false), 1) => "unmerge-cells",
            (Some(true), _) => "set-and-merge-cells",
            (Some(false), _) => "set-and-unmerge-cells",
            (None, _) => base_operation,
        }
    };
    let node = editor.snapshot()?.get(&result_path, 0)?;
    let changed = editor.is_dirty();
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;

    Ok(CommandOutput::success(
        format!(
            "Updated {result_path} and saved '{}'.",
            output_path.display()
        ),
        serde_json::json!({
            "operation": operation,
            "changed": changed,
            "path": result_path,
            "node": node,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

async fn replace_text(parsed: ParsedArguments) -> UseResult<CommandOutput> {
    let find = parsed
        .find
        .as_deref()
        .ok_or_else(|| usage_error("office native set --replace requires --find <text>"))?;
    let replacement = parsed
        .replacement
        .as_deref()
        .ok_or_else(|| usage_error("office native set --find requires --replace <text>"))?;
    if [
        parsed.text.is_some(),
        parsed.number.is_some(),
        parsed.boolean.is_some(),
        parsed.formula.is_some(),
        parsed.width_emu.is_some(),
        parsed.bold.is_some(),
        parsed.italic.is_some(),
        parsed.underline.is_some(),
        parsed.script.is_some(),
        parsed.strikethrough.is_some(),
        parsed.double_strikethrough.is_some(),
        parsed.text_case.is_some(),
        parsed.highlight.is_some(),
        parsed.language.is_some(),
        parsed.font_family.is_some(),
        parsed.font_size.is_some(),
        parsed.text_color.is_some(),
        parsed.alignment.is_some(),
        parsed.number_format.is_some(),
        parsed.fill.is_some(),
        parsed.border.is_present(),
        parsed.vertical_alignment.is_some(),
        parsed.wrap_text.is_some(),
        parsed.text_rotation.is_some(),
        parsed.indent.is_some(),
        parsed.shrink_to_fit.is_some(),
        parsed.reading_order.is_some(),
        parsed.merge_cells.is_some(),
        parsed.url.is_some(),
        parsed.location.is_some(),
        parsed.display.is_some(),
        parsed.tooltip.is_some(),
        parsed.author.is_some(),
        parsed.initials.is_some(),
        parsed.x_emu.is_some(),
        parsed.y_emu.is_some(),
        parsed.has_data_validation_options(),
        parsed.has_named_range_options(),
        parsed.has_conditional_format_options(),
        parsed.has_spreadsheet_table_options(),
    ]
    .into_iter()
    .any(|present| present)
    {
        return Err(usage_error(
            "--find and --replace cannot be combined with other native Office set values",
        ));
    }

    let replacement = if parsed.regex {
        NativeOfficeTextReplacement::regex(find, replacement)?
    } else {
        NativeOfficeTextReplacement::literal(find, replacement)?
    };
    let source = &parsed.positionals[0];
    let path = &parsed.positionals[1];
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let result = editor.replace_text(path, replacement)?;
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;
    let human = if result.changed {
        format!(
            "Replaced {} match(es) in {path} and saved '{}'.",
            result.match_count,
            output_path.display()
        )
    } else {
        format!(
            "Found {} match(es) in {path}; no document text changed.",
            result.match_count
        )
    };
    Ok(CommandOutput::success(
        human,
        serde_json::json!({
            "operation": "replace-text",
            "changed": result.changed,
            "path": path,
            "matches": result.match_count,
            "result": result,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

fn parse_merge_cells(value: &str) -> UseResult<bool> {
    match value.to_ascii_lowercase().as_str() {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(usage_error(format!(
            "--merge-cells requires true or false, received '{value}'"
        ))),
    }
}

fn has_conditional_format_conflicts(parsed: &ParsedArguments) -> bool {
    [
        parsed.number.is_some(),
        parsed.boolean.is_some(),
        parsed.width_emu.is_some(),
        parsed.italic.is_some(),
        parsed.underline.is_some(),
        parsed.script.is_some(),
        parsed.strikethrough.is_some(),
        parsed.double_strikethrough.is_some(),
        parsed.text_case.is_some(),
        parsed.highlight.is_some(),
        parsed.language.is_some(),
        parsed.font_family.is_some(),
        parsed.font_size.is_some(),
        parsed.alignment.is_some(),
        parsed.number_format.is_some(),
        parsed.border.is_present(),
        parsed.vertical_alignment.is_some(),
        parsed.wrap_text.is_some(),
        parsed.text_rotation.is_some(),
        parsed.indent.is_some(),
        parsed.shrink_to_fit.is_some(),
        parsed.reading_order.is_some(),
        parsed.merge_cells.is_some(),
        parsed.url.is_some(),
        parsed.location.is_some(),
        parsed.display.is_some(),
        parsed.tooltip.is_some(),
        parsed.author.is_some(),
        parsed.initials.is_some(),
        parsed.x_emu.is_some(),
        parsed.y_emu.is_some(),
        parsed.has_data_validation_specific_options(),
        parsed.has_named_range_options(),
        parsed.has_spreadsheet_table_specific_options(),
        parsed.has_spreadsheet_filter_specific_options(),
    ]
    .into_iter()
    .any(|present| present)
}

fn has_spreadsheet_filter_conflicts(parsed: &ParsedArguments) -> bool {
    [
        parsed.text.is_some(),
        parsed.number.is_some(),
        parsed.boolean.is_some(),
        parsed.formula.is_some(),
        parsed.width_emu.is_some(),
        parsed.bold.is_some(),
        parsed.italic.is_some(),
        parsed.underline.is_some(),
        parsed.script.is_some(),
        parsed.strikethrough.is_some(),
        parsed.double_strikethrough.is_some(),
        parsed.text_case.is_some(),
        parsed.highlight.is_some(),
        parsed.language.is_some(),
        parsed.font_family.is_some(),
        parsed.font_size.is_some(),
        parsed.text_color.is_some(),
        parsed.alignment.is_some(),
        parsed.number_format.is_some(),
        parsed.fill.is_some(),
        parsed.border.is_present(),
        parsed.vertical_alignment.is_some(),
        parsed.wrap_text.is_some(),
        parsed.text_rotation.is_some(),
        parsed.indent.is_some(),
        parsed.shrink_to_fit.is_some(),
        parsed.reading_order.is_some(),
        parsed.merge_cells.is_some(),
        parsed.url.is_some(),
        parsed.location.is_some(),
        parsed.display.is_some(),
        parsed.tooltip.is_some(),
        parsed.author.is_some(),
        parsed.initials.is_some(),
        parsed.x_emu.is_some(),
        parsed.y_emu.is_some(),
        parsed.has_data_validation_specific_options(),
        parsed.validation_operator.is_some(),
        parsed.validation_formula1.is_some(),
        parsed.validation_formula2.is_some(),
        parsed.has_conditional_format_options(),
        parsed.has_named_range_options(),
        parsed.has_spreadsheet_table_specific_options(),
    ]
    .into_iter()
    .any(|present| present)
}

async fn remove(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::MUTATE)?;
    if parsed.positionals.len() != 2 {
        return Err(usage_error(
            "office native remove requires <file> and <path>",
        ));
    }
    let source = &parsed.positionals[0];
    let path = &parsed.positionals[1];
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    editor.remove(path)?;
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;

    Ok(CommandOutput::success(
        format!("Removed {path} and saved '{}'.", output_path.display()),
        serde_json::json!({
            "operation": "remove",
            "changed": true,
            "path": path,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

#[derive(Debug, Clone, Copy)]
enum StructureOperation {
    InsertRows,
    DeleteRows,
    InsertColumns,
    DeleteColumns,
}

impl StructureOperation {
    fn name(self) -> &'static str {
        match self {
            Self::InsertRows => "insert-rows",
            Self::DeleteRows => "delete-rows",
            Self::InsertColumns => "insert-columns",
            Self::DeleteColumns => "delete-columns",
        }
    }
}

async fn edit_structure(
    args: &[String],
    operation: StructureOperation,
) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::STRUCTURE)?;
    if parsed.positionals.len() != 3 {
        return Err(usage_error(format!(
            "office native {} requires <file>, <sheet>, and <start>",
            operation.name()
        )));
    }
    let source = &parsed.positionals[0];
    let sheet = &parsed.positionals[1];
    let start = &parsed.positionals[2];
    let count = parsed.count.unwrap_or(1);
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let path = match operation {
        StructureOperation::InsertRows => {
            editor.insert_rows(sheet, parse_row_start(start)?, count)?
        }
        StructureOperation::DeleteRows => {
            editor.delete_rows(sheet, parse_row_start(start)?, count)?
        }
        StructureOperation::InsertColumns => editor.insert_columns(sheet, start, count)?,
        StructureOperation::DeleteColumns => editor.delete_columns(sheet, start, count)?,
    };
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;
    Ok(CommandOutput::success(
        format!(
            "Applied {} at {path} and saved '{}'.",
            operation.name(),
            output_path.display()
        ),
        serde_json::json!({
            "operation": operation.name(),
            "changed": true,
            "sheet": sheet,
            "start": start,
            "count": count,
            "path": path,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

async fn rename_sheet(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::MUTATE)?;
    if parsed.positionals.len() != 3 {
        return Err(usage_error(
            "office native rename-sheet requires <file>, <sheet>, and <new-name>",
        ));
    }
    let source = &parsed.positionals[0];
    let sheet = &parsed.positionals[1];
    let name = &parsed.positionals[2];
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let path = editor.rename_worksheet(sheet, name)?;
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;
    Ok(CommandOutput::success(
        format!(
            "Renamed {sheet} to {path} and saved '{}'.",
            output_path.display()
        ),
        serde_json::json!({
            "operation": "rename-sheet",
            "changed": sheet != &path,
            "from": sheet,
            "path": path,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

async fn move_sheet(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::MUTATE)?;
    if parsed.positionals.len() != 3 {
        return Err(usage_error(
            "office native move-sheet requires <file>, <sheet>, and <one-based-position>",
        ));
    }
    let source = &parsed.positionals[0];
    let sheet = &parsed.positionals[1];
    let position = parsed.positionals[2].parse::<usize>().map_err(|_| {
        usage_error(format!(
            "move-sheet position must be a positive integer, received '{}'",
            parsed.positionals[2]
        ))
    })?;
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let path = editor.move_worksheet(sheet, position)?;
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;
    Ok(CommandOutput::success(
        format!(
            "Moved {sheet} to position {position} and saved '{}'.",
            output_path.display()
        ),
        serde_json::json!({
            "operation": "move-sheet",
            "changed": true,
            "path": path,
            "position": position,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

async fn copy_sheet(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::COPY)?;
    if parsed.positionals.len() != 3 {
        return Err(usage_error(
            "office native copy-sheet requires <file>, <sheet>, and <new-name>",
        ));
    }
    if parsed.position == Some(0) {
        return Err(usage_error("copy-sheet --position must be at least 1"));
    }
    let source = &parsed.positionals[0];
    let sheet = &parsed.positionals[1];
    let name = &parsed.positionals[2];
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let path = editor.copy_worksheet(sheet, name, parsed.position)?;
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;
    Ok(CommandOutput::success(
        format!(
            "Copied {sheet} to {path} and saved '{}'.",
            output_path.display()
        ),
        serde_json::json!({
            "operation": "copy-sheet",
            "changed": true,
            "from": sheet,
            "path": path,
            "position": parsed.position,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

fn parse_row_start(value: &str) -> UseResult<u32> {
    value.parse::<u32>().map_err(|_| {
        usage_error(format!(
            "Spreadsheet row start must be a positive integer, received '{value}'"
        ))
    })
}

async fn batch(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::BATCH)?;
    if parsed.positionals.len() != 1 {
        return Err(usage_error("office native batch requires <file>"));
    }
    let input_path = parsed
        .input
        .as_deref()
        .ok_or_else(|| usage_error("office native batch requires --input <batch.json>"))?;
    let input = replay::read_batch_input(input_path).await?;
    let mutation_count = input.mutation_count();
    let replayed = matches!(&input, replay::NativeBatchInput::Replay(_));
    let source = &parsed.positionals[0];
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let result = match &input {
        replay::NativeBatchInput::Mutations(mutations) => editor.apply_batch(mutations)?,
        replay::NativeBatchInput::Replay(artifact) => editor.apply_replay(artifact)?,
    };
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;

    Ok(CommandOutput::success(
        format!(
            "Applied {} native Office mutation(s) and saved '{}'.",
            result.applied,
            output_path.display()
        ),
        serde_json::json!({
            "operation": "batch",
            "changed": result.applied > 0,
            "atomic": true,
            "replay": replayed,
            "inputMutations": mutation_count,
            "result": result,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

async fn save_editor(editor: &mut NativeOfficeEditor, output: Option<&str>) -> UseResult<()> {
    if let Some(output) = output {
        editor.save_as(output).await
    } else {
        editor.save().await
    }
}

fn format_node(node: &DocumentNode, level: usize) -> String {
    let mut output = format!(
        "{}{} ({}) \"{}\" children={}",
        "  ".repeat(level),
        node.path,
        node.node_type.label(),
        single_line(&node.text),
        node.child_count
    );
    for (key, value) in &node.format {
        output.push_str(&format!(" {key}={}", single_line(value)));
    }
    for child in &node.children {
        output.push('\n');
        output.push_str(&format_node(child, level + 1));
    }
    output
}

fn single_line(value: &str) -> String {
    value
        .replace('\\', "\\\\")
        .replace('\n', "\\n")
        .replace('"', "\\\"")
}

fn usage_error(message: impl Into<String>) -> UseError {
    UseError::new("use.cli.invalid_usage", message).with_suggestion(HELP)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_arguments_are_bounded_and_reject_unknown_options() {
        assert_eq!(
            ParsedArguments::parse(
                &["get".into(), "file.docx".into(), "--depth".into()],
                AllowedOptions::GET,
            )
            .unwrap_err()
            .code,
            "use.cli.invalid_usage"
        );
        assert_eq!(
            ParsedArguments::parse(&["view".into(), "--output".into()], AllowedOptions::NONE)
                .unwrap_err()
                .code,
            "use.cli.invalid_usage"
        );
        assert_eq!(
            ParsedArguments::parse(
                &[
                    "set".into(),
                    "file.docx".into(),
                    "/body/p[1]".into(),
                    "--text".into(),
                    "one".into(),
                    "--text".into(),
                    "two".into(),
                ],
                AllowedOptions::SET,
            )
            .unwrap_err()
            .code,
            "use.cli.invalid_usage"
        );

        let parsed = ParsedArguments::parse(
            &[
                "add".into(),
                "file.docx".into(),
                "/body".into(),
                "--type".into(),
                "table".into(),
                "--rows".into(),
                "2".into(),
                "--cols".into(),
                "3".into(),
            ],
            AllowedOptions::ADD,
        )
        .unwrap();
        assert_eq!(parsed.rows, Some(2));
        assert_eq!(parsed.columns, Some(3));

        let parsed = ParsedArguments::parse(
            &[
                "view".into(),
                "book.xlsx".into(),
                "issues".into(),
                "--type".into(),
                "formula_not_evaluated".into(),
                "--limit".into(),
                "10".into(),
            ],
            AllowedOptions::VIEW,
        )
        .unwrap();
        assert_eq!(parsed.node_type.as_deref(), Some("formula_not_evaluated"));
        assert_eq!(parsed.limit, Some(10));

        assert_eq!(
            ParsedArguments::parse(
                &[
                    "add".into(),
                    "file.docx".into(),
                    "/body".into(),
                    "--columns".into(),
                    "2".into(),
                    "--cols".into(),
                    "3".into(),
                ],
                AllowedOptions::ADD,
            )
            .unwrap_err()
            .code,
            "use.cli.invalid_usage"
        );
    }

    #[tokio::test]
    async fn native_batch_inputs_are_versioned_and_size_bounded() {
        let oversized = tempfile::NamedTempFile::new().unwrap();
        oversized
            .as_file()
            .set_len(MAX_BATCH_INPUT_BYTES + 1)
            .unwrap();
        let error = replay::read_batch_input(oversized.path().to_str().unwrap())
            .await
            .unwrap_err();
        assert_eq!(error.code, "use.office.batch_input_too_large");

        let unsupported = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(unsupported.path(), br#"{"schemaVersion":2,"mutations":[]}"#).unwrap();
        let error = replay::read_batch_input(unsupported.path().to_str().unwrap())
            .await
            .unwrap_err();
        assert_eq!(error.code, "use.office.batch_schema_unsupported");

        let unsupported_replay = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(
            unsupported_replay.path(),
            serde_json::to_vec(&serde_json::json!({
                "format": "a3s.office.native-replay",
                "schemaVersion": 2,
                "documentKind": "word",
                "scope": "/",
                "base": "blank",
                "baseSha256": "0".repeat(64),
                "resultSha256": "0".repeat(64),
                "mutations": []
            }))
            .unwrap(),
        )
        .unwrap();
        let error = replay::read_batch_input(unsupported_replay.path().to_str().unwrap())
            .await
            .unwrap_err();
        assert_eq!(error.code, "use.office.replay_schema_unsupported");

        let oversized_xml = tempfile::NamedTempFile::new().unwrap();
        oversized_xml
            .as_file()
            .set_len(raw::MAX_RAW_XML_INPUT_BYTES + 1)
            .unwrap();
        let error = raw::read_xml_input(oversized_xml.path().to_str().unwrap())
            .await
            .unwrap_err();
        assert_eq!(error.code, "use.office.raw_input_too_large");

        let invalid_utf8 = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(invalid_utf8.path(), [0xff, 0xfe, 0xfd]).unwrap();
        let error = raw::read_xml_input(invalid_utf8.path().to_str().unwrap())
            .await
            .unwrap_err();
        assert_eq!(error.code, "use.office.raw_input_invalid");

        #[cfg(unix)]
        {
            let directory = tempfile::tempdir().unwrap();
            let link = directory.path().join("input.xml");
            std::os::unix::fs::symlink(invalid_utf8.path(), &link).unwrap();
            let error = raw::read_xml_input(link.to_str().unwrap())
                .await
                .unwrap_err();
            assert_eq!(error.code, "use.office.raw_input_invalid");
        }
    }
}
