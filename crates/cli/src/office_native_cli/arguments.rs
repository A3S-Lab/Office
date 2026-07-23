use a3s_use_core::UseResult;

use super::usage_error;

mod border;

pub(super) use border::BorderArguments;

#[derive(Debug, Default)]
pub(super) struct ParsedArguments {
    pub positionals: Vec<String>,
    pub depth: Option<usize>,
    pub limit: Option<usize>,
    pub timeout_ms: Option<u64>,
    pub text: Option<String>,
    pub find: Option<String>,
    pub replacement: Option<String>,
    pub output: Option<String>,
    pub input: Option<String>,
    pub node_type: Option<String>,
    pub name: Option<String>,
    pub alt: Option<String>,
    pub width: Option<u32>,
    pub width_emu: Option<u64>,
    pub height: Option<u32>,
    pub rows: Option<usize>,
    pub columns: Option<usize>,
    pub number: Option<String>,
    pub boolean: Option<String>,
    pub formula: Option<String>,
    pub bold: Option<String>,
    pub italic: Option<String>,
    pub underline: Option<String>,
    pub script: Option<String>,
    pub strikethrough: Option<String>,
    pub double_strikethrough: Option<String>,
    pub text_case: Option<String>,
    pub highlight: Option<String>,
    pub language: Option<String>,
    pub font_family: Option<String>,
    pub font_size: Option<String>,
    pub text_color: Option<String>,
    pub alignment: Option<String>,
    pub number_format: Option<String>,
    pub fill: Option<String>,
    pub border: BorderArguments,
    pub vertical_alignment: Option<String>,
    pub wrap_text: Option<String>,
    pub text_rotation: Option<String>,
    pub indent: Option<String>,
    pub shrink_to_fit: Option<String>,
    pub reading_order: Option<String>,
    pub merge_cells: Option<String>,
    pub validation_ranges: Vec<String>,
    pub validation_type: Option<String>,
    pub validation_operator: Option<String>,
    pub validation_formula1: Option<String>,
    pub validation_formula2: Option<String>,
    pub validation_allow_blank: Option<String>,
    pub validation_show_input: Option<String>,
    pub validation_show_error: Option<String>,
    pub validation_prompt_title: Option<String>,
    pub validation_prompt: Option<String>,
    pub validation_error_title: Option<String>,
    pub validation_error_message: Option<String>,
    pub validation_error_style: Option<String>,
    pub validation_in_cell_dropdown: Option<String>,
    pub named_range_ref: Option<String>,
    pub named_range_scope: Option<String>,
    pub named_range_comment: Option<String>,
    pub named_range_volatile: Option<String>,
    pub table_display_name: Option<String>,
    pub table_columns: Vec<String>,
    pub table_header_row: Option<String>,
    pub table_totals_row: Option<String>,
    pub table_style: Option<String>,
    pub table_show_first_column: Option<String>,
    pub table_show_last_column: Option<String>,
    pub table_show_row_stripes: Option<String>,
    pub table_show_column_stripes: Option<String>,
    pub spreadsheet_filters: Vec<String>,
    pub clear_filters: bool,
    pub conditional_format_type: Option<String>,
    pub conditional_stop_if_true: Option<String>,
    pub conditional_rank: Option<u32>,
    pub conditional_percent: Option<String>,
    pub conditional_bottom: Option<String>,
    pub conditional_above: Option<String>,
    pub conditional_equal: Option<String>,
    pub conditional_standard_deviations: Option<u32>,
    pub conditional_period: Option<String>,
    pub conditional_color: Option<String>,
    pub conditional_min: Option<String>,
    pub conditional_max: Option<String>,
    pub conditional_show_value: Option<String>,
    pub conditional_min_length: Option<u32>,
    pub conditional_max_length: Option<u32>,
    pub conditional_min_color: Option<String>,
    pub conditional_mid_color: Option<String>,
    pub conditional_max_color: Option<String>,
    pub conditional_midpoint: Option<String>,
    pub conditional_icon_set: Option<String>,
    pub conditional_reverse: Option<String>,
    pub conditional_thresholds: Vec<String>,
    pub url: Option<String>,
    pub location: Option<String>,
    pub display: Option<String>,
    pub tooltip: Option<String>,
    pub author: Option<String>,
    pub initials: Option<String>,
    pub x_emu: Option<i32>,
    pub y_emu: Option<i32>,
    pub count: Option<u32>,
    pub position: Option<usize>,
    pub index: Option<usize>,
    pub target_parent: Option<String>,
    pub before: Option<String>,
    pub after: Option<String>,
    pub data: Option<String>,
    pub force: bool,
    pub regex: bool,
}

impl ParsedArguments {
    pub fn parse(args: &[String], allowed: AllowedOptions) -> UseResult<Self> {
        let mut parsed = Self::default();
        let mut index = 1;
        while index < args.len() {
            match args[index].as_str() {
                "--json" => index += 1,
                "--" => {
                    parsed.positionals.extend_from_slice(&args[index + 1..]);
                    break;
                }
                "--depth" if allowed.depth => {
                    if parsed.depth.is_some() {
                        return Err(usage_error("--depth may be specified only once"));
                    }
                    let value = option_value(args, index, "--depth")?;
                    parsed.depth = Some(value.parse::<usize>().map_err(|_| {
                        usage_error(format!(
                            "--depth requires a non-negative integer, received '{value}'"
                        ))
                    })?);
                    index += 2;
                }
                "--limit" if allowed.limit => {
                    set_usize_option(&mut parsed.limit, args, index, "--limit")?;
                    index += 2;
                }
                "--timeout-ms" if allowed.timeout_ms => {
                    set_u64_option(&mut parsed.timeout_ms, args, index, "--timeout-ms")?;
                    index += 2;
                }
                "--text" if allowed.text => {
                    set_string_option(&mut parsed.text, args, index, "--text")?;
                    index += 2;
                }
                "--find" if allowed.find => {
                    set_string_option(&mut parsed.find, args, index, "--find")?;
                    index += 2;
                }
                "--replace" if allowed.replacement => {
                    set_string_option(&mut parsed.replacement, args, index, "--replace")?;
                    index += 2;
                }
                "--number" if allowed.number => {
                    set_string_option(&mut parsed.number, args, index, "--number")?;
                    index += 2;
                }
                "--boolean" | "--bool" if allowed.boolean => {
                    set_string_option(&mut parsed.boolean, args, index, "--boolean")?;
                    index += 2;
                }
                "--formula" if allowed.formula => {
                    set_string_option(&mut parsed.formula, args, index, "--formula")?;
                    index += 2;
                }
                "--bold" if allowed.bold => {
                    set_string_option(&mut parsed.bold, args, index, "--bold")?;
                    index += 2;
                }
                "--italic" if allowed.italic => {
                    set_string_option(&mut parsed.italic, args, index, "--italic")?;
                    index += 2;
                }
                "--underline" if allowed.underline => {
                    set_string_option(&mut parsed.underline, args, index, "--underline")?;
                    index += 2;
                }
                "--script" if allowed.script => {
                    set_string_option(&mut parsed.script, args, index, "--script")?;
                    index += 2;
                }
                "--strikethrough" if allowed.strikethrough => {
                    set_string_option(&mut parsed.strikethrough, args, index, "--strikethrough")?;
                    index += 2;
                }
                "--double-strikethrough" if allowed.double_strikethrough => {
                    set_string_option(
                        &mut parsed.double_strikethrough,
                        args,
                        index,
                        "--double-strikethrough",
                    )?;
                    index += 2;
                }
                "--text-case" if allowed.text_case => {
                    set_string_option(&mut parsed.text_case, args, index, "--text-case")?;
                    index += 2;
                }
                "--highlight" if allowed.highlight => {
                    set_string_option(&mut parsed.highlight, args, index, "--highlight")?;
                    index += 2;
                }
                "--language" | "--lang" if allowed.language => {
                    set_string_option(&mut parsed.language, args, index, "--language")?;
                    index += 2;
                }
                "--font-family" if allowed.font_family => {
                    set_string_option(&mut parsed.font_family, args, index, "--font-family")?;
                    index += 2;
                }
                "--font-size" if allowed.font_size => {
                    set_string_option(&mut parsed.font_size, args, index, "--font-size")?;
                    index += 2;
                }
                "--text-color" if allowed.text_color => {
                    set_string_option(&mut parsed.text_color, args, index, "--text-color")?;
                    index += 2;
                }
                "--align" | "--alignment" if allowed.alignment => {
                    set_string_option(&mut parsed.alignment, args, index, "--align")?;
                    index += 2;
                }
                "--number-format" | "--numfmt" if allowed.number_format => {
                    set_string_option(&mut parsed.number_format, args, index, "--number-format")?;
                    index += 2;
                }
                "--fill" | "--fill-color" if allowed.fill => {
                    set_string_option(&mut parsed.fill, args, index, "--fill")?;
                    index += 2;
                }
                "--border"
                | "--border-all"
                | "--border-color"
                | "--border-left"
                | "--border-left-color"
                | "--border-right"
                | "--border-right-color"
                | "--border-top"
                | "--border-top-color"
                | "--border-bottom"
                | "--border-bottom-color"
                | "--border-diagonal"
                | "--border-diagonal-color"
                | "--border-diagonal-up"
                | "--border-diagonal-down"
                    if allowed.border =>
                {
                    parsed.border.parse(&args[index], args, index)?;
                    index += 2;
                }
                "--vertical-align" | "--valign" if allowed.vertical_alignment => {
                    set_string_option(
                        &mut parsed.vertical_alignment,
                        args,
                        index,
                        "--vertical-align",
                    )?;
                    index += 2;
                }
                "--wrap-text" | "--wrap" if allowed.wrap_text => {
                    set_string_option(&mut parsed.wrap_text, args, index, "--wrap-text")?;
                    index += 2;
                }
                "--text-rotation" | "--rotation" if allowed.text_rotation => {
                    set_string_option(&mut parsed.text_rotation, args, index, "--text-rotation")?;
                    index += 2;
                }
                "--indent" if allowed.indent => {
                    set_string_option(&mut parsed.indent, args, index, "--indent")?;
                    index += 2;
                }
                "--shrink-to-fit" if allowed.shrink_to_fit => {
                    set_string_option(&mut parsed.shrink_to_fit, args, index, "--shrink-to-fit")?;
                    index += 2;
                }
                "--reading-order" | "--cell-direction" if allowed.reading_order => {
                    set_string_option(&mut parsed.reading_order, args, index, "--reading-order")?;
                    index += 2;
                }
                "--merge-cells" if allowed.merge_cells => {
                    set_string_option(&mut parsed.merge_cells, args, index, "--merge-cells")?;
                    index += 2;
                }
                "--range" | "--sqref"
                    if allowed.data_validation
                        || allowed.conditional_formatting
                        || allowed.spreadsheet_filter =>
                {
                    parsed
                        .validation_ranges
                        .push(option_value(args, index, "--range")?.to_string());
                    index += 2;
                }
                "--validation-type" if allowed.data_validation => {
                    set_string_option(
                        &mut parsed.validation_type,
                        args,
                        index,
                        "--validation-type",
                    )?;
                    index += 2;
                }
                "--operator" if allowed.data_validation || allowed.conditional_formatting => {
                    set_string_option(&mut parsed.validation_operator, args, index, "--operator")?;
                    index += 2;
                }
                "--formula1" if allowed.data_validation || allowed.conditional_formatting => {
                    set_string_option(&mut parsed.validation_formula1, args, index, "--formula1")?;
                    index += 2;
                }
                "--formula2" if allowed.data_validation || allowed.conditional_formatting => {
                    set_string_option(&mut parsed.validation_formula2, args, index, "--formula2")?;
                    index += 2;
                }
                "--allow-blank" if allowed.data_validation => {
                    set_string_option(
                        &mut parsed.validation_allow_blank,
                        args,
                        index,
                        "--allow-blank",
                    )?;
                    index += 2;
                }
                "--show-input" if allowed.data_validation => {
                    set_string_option(
                        &mut parsed.validation_show_input,
                        args,
                        index,
                        "--show-input",
                    )?;
                    index += 2;
                }
                "--show-error" if allowed.data_validation => {
                    set_string_option(
                        &mut parsed.validation_show_error,
                        args,
                        index,
                        "--show-error",
                    )?;
                    index += 2;
                }
                "--prompt-title" if allowed.data_validation => {
                    set_string_option(
                        &mut parsed.validation_prompt_title,
                        args,
                        index,
                        "--prompt-title",
                    )?;
                    index += 2;
                }
                "--prompt" if allowed.data_validation => {
                    set_string_option(&mut parsed.validation_prompt, args, index, "--prompt")?;
                    index += 2;
                }
                "--error-title" if allowed.data_validation => {
                    set_string_option(
                        &mut parsed.validation_error_title,
                        args,
                        index,
                        "--error-title",
                    )?;
                    index += 2;
                }
                "--error-message" if allowed.data_validation => {
                    set_string_option(
                        &mut parsed.validation_error_message,
                        args,
                        index,
                        "--error-message",
                    )?;
                    index += 2;
                }
                "--error-style" if allowed.data_validation => {
                    set_string_option(
                        &mut parsed.validation_error_style,
                        args,
                        index,
                        "--error-style",
                    )?;
                    index += 2;
                }
                "--in-cell-dropdown" if allowed.data_validation => {
                    set_string_option(
                        &mut parsed.validation_in_cell_dropdown,
                        args,
                        index,
                        "--in-cell-dropdown",
                    )?;
                    index += 2;
                }
                "--ref" | "--refers-to" if allowed.named_range => {
                    set_string_option(&mut parsed.named_range_ref, args, index, "--ref")?;
                    index += 2;
                }
                "--scope" if allowed.named_range => {
                    set_string_option(&mut parsed.named_range_scope, args, index, "--scope")?;
                    index += 2;
                }
                "--comment" if allowed.named_range => {
                    set_string_option(&mut parsed.named_range_comment, args, index, "--comment")?;
                    index += 2;
                }
                "--volatile" if allowed.named_range => {
                    set_string_option(&mut parsed.named_range_volatile, args, index, "--volatile")?;
                    index += 2;
                }
                "--display-name" if allowed.spreadsheet_table => {
                    set_string_option(
                        &mut parsed.table_display_name,
                        args,
                        index,
                        "--display-name",
                    )?;
                    index += 2;
                }
                "--table-column" if allowed.spreadsheet_table => {
                    parsed
                        .table_columns
                        .push(option_value(args, index, "--table-column")?.to_string());
                    index += 2;
                }
                "--header-row" if allowed.spreadsheet_table => {
                    set_string_option(&mut parsed.table_header_row, args, index, "--header-row")?;
                    index += 2;
                }
                "--totals-row" if allowed.spreadsheet_table => {
                    set_string_option(&mut parsed.table_totals_row, args, index, "--totals-row")?;
                    index += 2;
                }
                "--style" | "--table-style" if allowed.spreadsheet_table => {
                    set_string_option(&mut parsed.table_style, args, index, "--style")?;
                    index += 2;
                }
                "--show-first-column" if allowed.spreadsheet_table => {
                    set_string_option(
                        &mut parsed.table_show_first_column,
                        args,
                        index,
                        "--show-first-column",
                    )?;
                    index += 2;
                }
                "--show-last-column" if allowed.spreadsheet_table => {
                    set_string_option(
                        &mut parsed.table_show_last_column,
                        args,
                        index,
                        "--show-last-column",
                    )?;
                    index += 2;
                }
                "--show-row-stripes" if allowed.spreadsheet_table => {
                    set_string_option(
                        &mut parsed.table_show_row_stripes,
                        args,
                        index,
                        "--show-row-stripes",
                    )?;
                    index += 2;
                }
                "--show-column-stripes" if allowed.spreadsheet_table => {
                    set_string_option(
                        &mut parsed.table_show_column_stripes,
                        args,
                        index,
                        "--show-column-stripes",
                    )?;
                    index += 2;
                }
                "--filter" if allowed.spreadsheet_filter => {
                    parsed
                        .spreadsheet_filters
                        .push(option_value(args, index, "--filter")?.to_string());
                    index += 2;
                }
                "--clear-filters" if allowed.spreadsheet_filter => {
                    if parsed.clear_filters {
                        return Err(usage_error("--clear-filters may be specified only once"));
                    }
                    parsed.clear_filters = true;
                    index += 1;
                }
                "--rule-type" | "--cf-type" if allowed.conditional_formatting => {
                    set_string_option(
                        &mut parsed.conditional_format_type,
                        args,
                        index,
                        "--rule-type",
                    )?;
                    index += 2;
                }
                "--stop-if-true" if allowed.conditional_formatting => {
                    set_string_option(
                        &mut parsed.conditional_stop_if_true,
                        args,
                        index,
                        "--stop-if-true",
                    )?;
                    index += 2;
                }
                "--rank" if allowed.conditional_formatting => {
                    set_u32_option(&mut parsed.conditional_rank, args, index, "--rank")?;
                    index += 2;
                }
                "--percent" if allowed.conditional_formatting => {
                    set_string_option(&mut parsed.conditional_percent, args, index, "--percent")?;
                    index += 2;
                }
                "--bottom" if allowed.conditional_formatting => {
                    set_string_option(&mut parsed.conditional_bottom, args, index, "--bottom")?;
                    index += 2;
                }
                "--above" | "--above-average" if allowed.conditional_formatting => {
                    set_string_option(&mut parsed.conditional_above, args, index, "--above")?;
                    index += 2;
                }
                "--equal-average" if allowed.conditional_formatting => {
                    set_string_option(
                        &mut parsed.conditional_equal,
                        args,
                        index,
                        "--equal-average",
                    )?;
                    index += 2;
                }
                "--std-dev" | "--standard-deviations" if allowed.conditional_formatting => {
                    set_u32_option(
                        &mut parsed.conditional_standard_deviations,
                        args,
                        index,
                        "--std-dev",
                    )?;
                    index += 2;
                }
                "--period" if allowed.conditional_formatting => {
                    set_string_option(&mut parsed.conditional_period, args, index, "--period")?;
                    index += 2;
                }
                "--color" | "--bar-color" if allowed.conditional_formatting => {
                    set_string_option(&mut parsed.conditional_color, args, index, "--color")?;
                    index += 2;
                }
                "--min" if allowed.conditional_formatting => {
                    set_string_option(&mut parsed.conditional_min, args, index, "--min")?;
                    index += 2;
                }
                "--max" if allowed.conditional_formatting => {
                    set_string_option(&mut parsed.conditional_max, args, index, "--max")?;
                    index += 2;
                }
                "--show-value" if allowed.conditional_formatting => {
                    set_string_option(
                        &mut parsed.conditional_show_value,
                        args,
                        index,
                        "--show-value",
                    )?;
                    index += 2;
                }
                "--min-length" if allowed.conditional_formatting => {
                    set_u32_option(
                        &mut parsed.conditional_min_length,
                        args,
                        index,
                        "--min-length",
                    )?;
                    index += 2;
                }
                "--max-length" if allowed.conditional_formatting => {
                    set_u32_option(
                        &mut parsed.conditional_max_length,
                        args,
                        index,
                        "--max-length",
                    )?;
                    index += 2;
                }
                "--min-color" if allowed.conditional_formatting => {
                    set_string_option(
                        &mut parsed.conditional_min_color,
                        args,
                        index,
                        "--min-color",
                    )?;
                    index += 2;
                }
                "--mid-color" if allowed.conditional_formatting => {
                    set_string_option(
                        &mut parsed.conditional_mid_color,
                        args,
                        index,
                        "--mid-color",
                    )?;
                    index += 2;
                }
                "--max-color" if allowed.conditional_formatting => {
                    set_string_option(
                        &mut parsed.conditional_max_color,
                        args,
                        index,
                        "--max-color",
                    )?;
                    index += 2;
                }
                "--midpoint" | "--mid-point" if allowed.conditional_formatting => {
                    set_string_option(&mut parsed.conditional_midpoint, args, index, "--midpoint")?;
                    index += 2;
                }
                "--icon-set" | "--icons" if allowed.conditional_formatting => {
                    set_string_option(&mut parsed.conditional_icon_set, args, index, "--icon-set")?;
                    index += 2;
                }
                "--reverse" if allowed.conditional_formatting => {
                    set_string_option(&mut parsed.conditional_reverse, args, index, "--reverse")?;
                    index += 2;
                }
                "--threshold" if allowed.conditional_formatting => {
                    parsed
                        .conditional_thresholds
                        .push(option_value(args, index, "--threshold")?.to_string());
                    index += 2;
                }
                "--url" | "--link" | "--href" if allowed.url => {
                    set_string_option(&mut parsed.url, args, index, "--url")?;
                    index += 2;
                }
                "--location" if allowed.location => {
                    set_string_option(&mut parsed.location, args, index, "--location")?;
                    index += 2;
                }
                "--display" if allowed.display => {
                    set_string_option(&mut parsed.display, args, index, "--display")?;
                    index += 2;
                }
                "--tooltip" if allowed.tooltip => {
                    set_string_option(&mut parsed.tooltip, args, index, "--tooltip")?;
                    index += 2;
                }
                "--author" if allowed.author => {
                    set_string_option(&mut parsed.author, args, index, "--author")?;
                    index += 2;
                }
                "--initials" if allowed.initials => {
                    set_string_option(&mut parsed.initials, args, index, "--initials")?;
                    index += 2;
                }
                "--x-emu" if allowed.x_emu => {
                    set_i32_option(&mut parsed.x_emu, args, index, "--x-emu")?;
                    index += 2;
                }
                "--y-emu" if allowed.y_emu => {
                    set_i32_option(&mut parsed.y_emu, args, index, "--y-emu")?;
                    index += 2;
                }
                "--output" if allowed.output => {
                    set_string_option(&mut parsed.output, args, index, "--output")?;
                    index += 2;
                }
                "--input" if allowed.input => {
                    set_string_option(&mut parsed.input, args, index, "--input")?;
                    index += 2;
                }
                "--type" if allowed.node_type => {
                    set_string_option(&mut parsed.node_type, args, index, "--type")?;
                    index += 2;
                }
                "--name" if allowed.name => {
                    set_string_option(&mut parsed.name, args, index, "--name")?;
                    index += 2;
                }
                "--alt" if allowed.alt => {
                    set_string_option(&mut parsed.alt, args, index, "--alt")?;
                    index += 2;
                }
                "--width" if allowed.width => {
                    set_u32_option(&mut parsed.width, args, index, "--width")?;
                    index += 2;
                }
                "--width-emu" if allowed.width_emu => {
                    set_u64_option(&mut parsed.width_emu, args, index, "--width-emu")?;
                    index += 2;
                }
                "--height" if allowed.height => {
                    set_u32_option(&mut parsed.height, args, index, "--height")?;
                    index += 2;
                }
                "--rows" if allowed.rows => {
                    set_usize_option(&mut parsed.rows, args, index, "--rows")?;
                    index += 2;
                }
                "--columns" | "--cols" if allowed.columns => {
                    set_usize_option(&mut parsed.columns, args, index, "--columns")?;
                    index += 2;
                }
                "--count" if allowed.count => {
                    if parsed.count.is_some() {
                        return Err(usage_error("--count may be specified only once"));
                    }
                    let value = option_value(args, index, "--count")?;
                    parsed.count = Some(value.parse::<u32>().map_err(|_| {
                        usage_error(format!(
                            "--count requires a non-negative integer, received '{value}'"
                        ))
                    })?);
                    index += 2;
                }
                "--position" if allowed.position => {
                    set_usize_option(&mut parsed.position, args, index, "--position")?;
                    index += 2;
                }
                "--index" if allowed.index => {
                    set_usize_option(&mut parsed.index, args, index, "--index")?;
                    index += 2;
                }
                "--to" if allowed.target_parent => {
                    set_string_option(&mut parsed.target_parent, args, index, "--to")?;
                    index += 2;
                }
                "--before" if allowed.before => {
                    set_string_option(&mut parsed.before, args, index, "--before")?;
                    index += 2;
                }
                "--after" if allowed.after => {
                    set_string_option(&mut parsed.after, args, index, "--after")?;
                    index += 2;
                }
                "--data" if allowed.data => {
                    set_string_option(&mut parsed.data, args, index, "--data")?;
                    index += 2;
                }
                "--force" if allowed.force => {
                    if parsed.force {
                        return Err(usage_error("--force may be specified only once"));
                    }
                    parsed.force = true;
                    index += 1;
                }
                "--regex" if allowed.regex => {
                    if parsed.regex {
                        return Err(usage_error("--regex may be specified only once"));
                    }
                    parsed.regex = true;
                    index += 1;
                }
                option if option.starts_with('-') => {
                    return Err(usage_error(format!(
                        "unknown native Office option '{option}'"
                    )));
                }
                value => {
                    parsed.positionals.push(value.to_string());
                    index += 1;
                }
            }
        }
        Ok(parsed)
    }

    pub(super) fn has_data_validation_options(&self) -> bool {
        self.has_shared_rule_options() || self.has_data_validation_specific_options()
    }

    pub(super) fn has_shared_rule_options(&self) -> bool {
        !self.validation_ranges.is_empty()
            || self.validation_operator.is_some()
            || self.validation_formula1.is_some()
            || self.validation_formula2.is_some()
    }

    pub(super) fn has_data_validation_specific_options(&self) -> bool {
        self.validation_type.is_some()
            || self.validation_allow_blank.is_some()
            || self.validation_show_input.is_some()
            || self.validation_show_error.is_some()
            || self.validation_prompt_title.is_some()
            || self.validation_prompt.is_some()
            || self.validation_error_title.is_some()
            || self.validation_error_message.is_some()
            || self.validation_error_style.is_some()
            || self.validation_in_cell_dropdown.is_some()
    }

    pub(super) fn has_named_range_options(&self) -> bool {
        self.name.is_some()
            || self.named_range_ref.is_some()
            || self.named_range_scope.is_some()
            || self.named_range_comment.is_some()
            || self.named_range_volatile.is_some()
    }

    pub(super) fn has_spreadsheet_table_options(&self) -> bool {
        self.name.is_some()
            || !self.validation_ranges.is_empty()
            || self.table_display_name.is_some()
            || !self.table_columns.is_empty()
            || self.table_header_row.is_some()
            || self.table_totals_row.is_some()
            || self.table_style.is_some()
            || self.table_show_first_column.is_some()
            || self.table_show_last_column.is_some()
            || self.table_show_row_stripes.is_some()
            || self.table_show_column_stripes.is_some()
            || self.has_spreadsheet_filter_specific_options()
    }

    pub(super) fn has_spreadsheet_table_specific_options(&self) -> bool {
        self.table_display_name.is_some()
            || !self.table_columns.is_empty()
            || self.table_header_row.is_some()
            || self.table_totals_row.is_some()
            || self.table_style.is_some()
            || self.table_show_first_column.is_some()
            || self.table_show_last_column.is_some()
            || self.table_show_row_stripes.is_some()
            || self.table_show_column_stripes.is_some()
    }

    pub(super) fn has_spreadsheet_filter_options(&self) -> bool {
        !self.validation_ranges.is_empty() || self.has_spreadsheet_filter_specific_options()
    }

    pub(super) fn has_spreadsheet_filter_specific_options(&self) -> bool {
        !self.spreadsheet_filters.is_empty() || self.clear_filters
    }

    pub(super) fn has_conditional_format_options(&self) -> bool {
        self.conditional_format_type.is_some()
            || self.conditional_stop_if_true.is_some()
            || self.conditional_rank.is_some()
            || self.conditional_percent.is_some()
            || self.conditional_bottom.is_some()
            || self.conditional_above.is_some()
            || self.conditional_equal.is_some()
            || self.conditional_standard_deviations.is_some()
            || self.conditional_period.is_some()
            || self.conditional_color.is_some()
            || self.conditional_min.is_some()
            || self.conditional_max.is_some()
            || self.conditional_show_value.is_some()
            || self.conditional_min_length.is_some()
            || self.conditional_max_length.is_some()
            || self.conditional_min_color.is_some()
            || self.conditional_mid_color.is_some()
            || self.conditional_max_color.is_some()
            || self.conditional_midpoint.is_some()
            || self.conditional_icon_set.is_some()
            || self.conditional_reverse.is_some()
            || !self.conditional_thresholds.is_empty()
    }

    pub(super) fn has_conditional_format_update_options(&self) -> bool {
        self.has_conditional_format_options()
            || self.has_shared_rule_options()
            || self.text.is_some()
            || self.formula.is_some()
            || self.fill.is_some()
            || self.text_color.is_some()
            || self.bold.is_some()
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct AllowedOptions {
    depth: bool,
    limit: bool,
    timeout_ms: bool,
    text: bool,
    find: bool,
    replacement: bool,
    output: bool,
    input: bool,
    node_type: bool,
    name: bool,
    alt: bool,
    width: bool,
    width_emu: bool,
    height: bool,
    rows: bool,
    columns: bool,
    number: bool,
    boolean: bool,
    formula: bool,
    bold: bool,
    italic: bool,
    underline: bool,
    script: bool,
    strikethrough: bool,
    double_strikethrough: bool,
    text_case: bool,
    highlight: bool,
    language: bool,
    font_family: bool,
    font_size: bool,
    text_color: bool,
    alignment: bool,
    number_format: bool,
    fill: bool,
    border: bool,
    vertical_alignment: bool,
    wrap_text: bool,
    text_rotation: bool,
    indent: bool,
    shrink_to_fit: bool,
    reading_order: bool,
    merge_cells: bool,
    url: bool,
    location: bool,
    display: bool,
    tooltip: bool,
    author: bool,
    initials: bool,
    x_emu: bool,
    y_emu: bool,
    count: bool,
    position: bool,
    index: bool,
    target_parent: bool,
    before: bool,
    after: bool,
    data: bool,
    force: bool,
    regex: bool,
    data_validation: bool,
    named_range: bool,
    conditional_formatting: bool,
    spreadsheet_table: bool,
    spreadsheet_filter: bool,
}

impl AllowedOptions {
    pub const NONE: Self = Self {
        depth: false,
        limit: false,
        timeout_ms: false,
        text: false,
        find: false,
        replacement: false,
        output: false,
        input: false,
        node_type: false,
        name: false,
        alt: false,
        width: false,
        width_emu: false,
        height: false,
        rows: false,
        columns: false,
        number: false,
        boolean: false,
        formula: false,
        bold: false,
        italic: false,
        underline: false,
        script: false,
        strikethrough: false,
        double_strikethrough: false,
        text_case: false,
        highlight: false,
        language: false,
        font_family: false,
        font_size: false,
        text_color: false,
        alignment: false,
        number_format: false,
        fill: false,
        border: false,
        vertical_alignment: false,
        wrap_text: false,
        text_rotation: false,
        indent: false,
        shrink_to_fit: false,
        reading_order: false,
        merge_cells: false,
        url: false,
        location: false,
        display: false,
        tooltip: false,
        author: false,
        initials: false,
        x_emu: false,
        y_emu: false,
        count: false,
        position: false,
        index: false,
        target_parent: false,
        before: false,
        after: false,
        data: false,
        force: false,
        regex: false,
        data_validation: false,
        named_range: false,
        conditional_formatting: false,
        spreadsheet_table: false,
        spreadsheet_filter: false,
    };
    pub const GET: Self = Self {
        depth: true,
        ..Self::NONE
    };
    pub const VIEW: Self = Self {
        output: true,
        timeout_ms: true,
        node_type: true,
        limit: true,
        ..Self::NONE
    };
    pub const SET: Self = Self {
        text: true,
        find: true,
        replacement: true,
        output: true,
        number: true,
        boolean: true,
        formula: true,
        bold: true,
        italic: true,
        underline: true,
        script: true,
        strikethrough: true,
        double_strikethrough: true,
        text_case: true,
        highlight: true,
        language: true,
        font_family: true,
        font_size: true,
        text_color: true,
        alignment: true,
        number_format: true,
        fill: true,
        border: true,
        vertical_alignment: true,
        wrap_text: true,
        text_rotation: true,
        indent: true,
        shrink_to_fit: true,
        reading_order: true,
        merge_cells: true,
        url: true,
        location: true,
        display: true,
        tooltip: true,
        author: true,
        initials: true,
        x_emu: true,
        y_emu: true,
        width_emu: true,
        regex: true,
        data_validation: true,
        named_range: true,
        conditional_formatting: true,
        name: true,
        spreadsheet_table: true,
        spreadsheet_filter: true,
        ..Self::NONE
    };
    pub const BATCH: Self = Self {
        output: true,
        input: true,
        ..Self::NONE
    };
    pub const RAW: Self = Self {
        output: true,
        ..Self::NONE
    };
    pub const DUMP: Self = Self {
        output: true,
        ..Self::NONE
    };
    pub const RAW_SET: Self = Self {
        output: true,
        input: true,
        ..Self::NONE
    };
    pub const ADD: Self = Self {
        text: true,
        output: true,
        input: true,
        node_type: true,
        name: true,
        alt: true,
        width: true,
        height: true,
        rows: true,
        columns: true,
        index: true,
        url: true,
        location: true,
        display: true,
        tooltip: true,
        author: true,
        initials: true,
        x_emu: true,
        y_emu: true,
        data_validation: true,
        named_range: true,
        conditional_formatting: true,
        spreadsheet_table: true,
        spreadsheet_filter: true,
        formula: true,
        bold: true,
        text_color: true,
        fill: true,
        ..Self::NONE
    };
    pub const ADD_PART: Self = Self {
        output: true,
        node_type: true,
        ..Self::NONE
    };
    pub const MUTATE: Self = Self {
        output: true,
        ..Self::NONE
    };
    pub const STRUCTURE: Self = Self {
        output: true,
        count: true,
        ..Self::NONE
    };
    pub const COPY: Self = Self {
        output: true,
        position: true,
        ..Self::NONE
    };
    pub const MOVE_NODE: Self = Self {
        output: true,
        index: true,
        target_parent: true,
        before: true,
        after: true,
        ..Self::NONE
    };
    pub const COPY_NODE: Self = Self {
        output: true,
        name: true,
        index: true,
        target_parent: true,
        before: true,
        after: true,
        ..Self::NONE
    };
    pub const MERGE: Self = Self {
        data: true,
        force: true,
        ..Self::NONE
    };
}

pub(super) fn parse_boolean_option(value: &str) -> UseResult<bool> {
    match value.to_ascii_lowercase().as_str() {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(usage_error(format!(
            "--boolean requires true or false, received '{value}'"
        ))),
    }
}

fn option_value<'a>(args: &'a [String], index: usize, option: &str) -> UseResult<&'a str> {
    args.get(index + 1)
        .map(String::as_str)
        .ok_or_else(|| usage_error(format!("{option} requires a value")))
}

fn set_string_option(
    target: &mut Option<String>,
    args: &[String],
    index: usize,
    option: &str,
) -> UseResult<()> {
    if target.is_some() {
        return Err(usage_error(format!("{option} may be specified only once")));
    }
    *target = Some(option_value(args, index, option)?.to_string());
    Ok(())
}

fn set_usize_option(
    target: &mut Option<usize>,
    args: &[String],
    index: usize,
    option: &str,
) -> UseResult<()> {
    if target.is_some() {
        return Err(usage_error(format!("{option} may be specified only once")));
    }
    let value = option_value(args, index, option)?;
    *target = Some(value.parse::<usize>().map_err(|_| {
        usage_error(format!(
            "{option} requires a non-negative integer, received '{value}'"
        ))
    })?);
    Ok(())
}

fn set_u32_option(
    target: &mut Option<u32>,
    args: &[String],
    index: usize,
    option: &str,
) -> UseResult<()> {
    if target.is_some() {
        return Err(usage_error(format!("{option} may be specified only once")));
    }
    let value = option_value(args, index, option)?;
    *target = Some(value.parse::<u32>().map_err(|_| {
        usage_error(format!(
            "{option} requires a non-negative integer, received '{value}'"
        ))
    })?);
    Ok(())
}

fn set_u64_option(
    target: &mut Option<u64>,
    args: &[String],
    index: usize,
    option: &str,
) -> UseResult<()> {
    if target.is_some() {
        return Err(usage_error(format!("{option} may be specified only once")));
    }
    let value = option_value(args, index, option)?;
    *target = Some(value.parse::<u64>().map_err(|_| {
        usage_error(format!(
            "{option} requires a non-negative integer, received '{value}'"
        ))
    })?);
    Ok(())
}

fn set_i32_option(
    target: &mut Option<i32>,
    args: &[String],
    index: usize,
    option: &str,
) -> UseResult<()> {
    if target.is_some() {
        return Err(usage_error(format!("{option} may be specified only once")));
    }
    let value = option_value(args, index, option)?;
    *target = Some(value.parse::<i32>().map_err(|_| {
        usage_error(format!(
            "{option} requires a signed 32-bit integer, received '{value}'"
        ))
    })?);
    Ok(())
}
