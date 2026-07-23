use a3s_office::{
    DocumentKind, NativeCreatedImage, NativeOfficeComment, NativeOfficeCommentPosition,
    NativeOfficeEditor, NativeOfficeImage,
};
use a3s_use_core::UseResult;

use super::arguments::{AllowedOptions, ParsedArguments};
use super::bounded_input::{read_bounded_input, NativeInputKind};
use super::conditional_formatting;
use super::data_validation;
use super::format::parse_hyperlink;
use super::named_range;
use super::spreadsheet_filter;
use super::spreadsheet_table;
use super::{save_editor, usage_error, MAX_IMAGE_INPUT_BYTES};
use crate::CommandOutput;

pub(super) async fn run(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::ADD)?;
    if parsed.positionals.len() != 2 {
        return Err(usage_error(
            "office native add requires <file> and <parent>",
        ));
    }
    let node_type = parsed
        .node_type
        .as_deref()
        .ok_or_else(|| usage_error("office native add requires --type <node-type>"))?;
    let source = &parsed.positionals[0];
    let parent = &parsed.positionals[1];
    validate_options(node_type, &parsed)?;
    let mut editor = NativeOfficeEditor::open(source).await?;
    let source_path = editor.package().path().to_path_buf();
    let text = parsed.text.as_deref().unwrap_or_default();
    let (operation, path, created_image): (&str, String, Option<NativeCreatedImage>) =
        match node_type {
            "paragraph" | "p" => ("add-paragraph", editor.add_paragraph(parent, text)?, None),
            "table" | "tbl" if editor.package().kind() == DocumentKind::Spreadsheet => {
                if parsed.rows.is_some() || parsed.columns.is_some() {
                    return Err(usage_error(
                        "native Spreadsheet table add uses --range and repeated --table-column, not --rows or --columns",
                    ));
                }
                (
                    "add-spreadsheet-table",
                    editor.add_spreadsheet_table(parent, spreadsheet_table::build_new(&parsed)?)?,
                    None,
                )
            }
            "table" | "tbl" => {
                if parsed.has_spreadsheet_table_specific_options()
                    || !parsed.validation_ranges.is_empty()
                    || parsed.name.is_some()
                {
                    return Err(usage_error(
                        "Word and Presentation table add does not accept Spreadsheet table options",
                    ));
                }
                (
                    "add-table",
                    editor.add_table(
                        parent,
                        parsed.rows.unwrap_or(1),
                        parsed.columns.unwrap_or(1),
                    )?,
                    None,
                )
            }
            "row" | "tr" => (
                "add-table-row",
                editor.add_table_row(parent, parsed.columns)?,
                None,
            ),
            "column" | "col" => (
                "add-table-column",
                editor.add_table_column(parent, parsed.index, text)?,
                None,
            ),
            "cell" | "tc" => ("add-table-cell", editor.add_table_cell(parent, text)?, None),
            "sheet" | "worksheet" => {
                if parent != "/" {
                    return Err(usage_error("native worksheets can be added only to /"));
                }
                let name = parsed
                    .name
                    .as_deref()
                    .ok_or_else(|| usage_error("native worksheet add requires --name <name>"))?;
                ("add-worksheet", editor.add_worksheet(name)?, None)
            }
            "slide" => ("add-slide", editor.add_slide(parent, text)?, None),
            "shape" => ("add-shape", editor.add_shape(parent, text)?, None),
            "hyperlink" | "link" => {
                let display = parsed.display.as_deref().or(parsed.text.as_deref());
                let hyperlink = parse_hyperlink(&parsed, display)?.ok_or_else(|| {
                    usage_error("native hyperlink add requires --url or --location")
                })?;
                (
                    "set-hyperlink",
                    editor.set_hyperlink(parent, hyperlink)?,
                    None,
                )
            }
            "comment" | "note-comment" => {
                let author = parsed
                    .author
                    .as_deref()
                    .ok_or_else(|| usage_error("native comment add requires --author <name>"))?;
                let body = parsed
                    .text
                    .as_deref()
                    .ok_or_else(|| usage_error("native comment add requires --text <value>"))?;
                let mut comment = NativeOfficeComment::new(author, body)?;
                if let Some(initials) = &parsed.initials {
                    comment = comment.with_initials(initials);
                }
                match (parsed.x_emu, parsed.y_emu) {
                    (Some(x_emu), Some(y_emu)) => {
                        comment =
                            comment.with_position(NativeOfficeCommentPosition::new(x_emu, y_emu));
                    }
                    (None, None) => {}
                    _ => return Err(usage_error(
                        "native Presentation comment coordinates require both --x-emu and --y-emu",
                    )),
                }
                ("add-comment", editor.add_comment(parent, comment)?, None)
            }
            "data-validation" | "datavalidation" | "validation" => (
                "add-data-validation",
                editor.add_data_validation(parent, data_validation::build_new(&parsed)?)?,
                None,
            ),
            "auto-filter" | "autofilter" | "filter" => (
                "add-spreadsheet-auto-filter",
                editor
                    .add_spreadsheet_auto_filter(parent, spreadsheet_filter::build_new(&parsed)?)?,
                None,
            ),
            "conditional-format" | "conditional-formatting" | "conditionalformatting" | "cf" => (
                "add-conditional-format",
                editor
                    .add_conditional_format(parent, conditional_formatting::build_new(&parsed)?)?,
                None,
            ),
            "named-range" | "namedrange" | "defined-name" | "definedname" => (
                "add-named-range",
                editor.add_named_range(named_range::build_new(parent, &parsed)?)?,
                None,
            ),
            "picture" | "image" | "img" => {
                let input = parsed.input.as_deref().ok_or_else(|| {
                    usage_error("native picture add requires --input <png|jpeg|gif>")
                })?;
                let bytes =
                    read_bounded_input(input, MAX_IMAGE_INPUT_BYTES, NativeInputKind::Image)
                        .await?;
                let mut image = NativeOfficeImage::from_bytes(bytes)?;
                if let Some(name) = &parsed.name {
                    image = image.with_name(name);
                }
                if let Some(alt) = &parsed.alt {
                    image = image.with_alt_text(alt);
                }
                if let Some(width) = parsed.width {
                    image = image.with_width_px(width);
                }
                if let Some(height) = parsed.height {
                    image = image.with_height_px(height);
                }
                let created = editor.add_image(parent, image)?;
                let path = created.path.clone();
                ("add-image", path, Some(created))
            }
            _ => {
                return Err(usage_error(format!(
                    "native Office add type '{node_type}' is not supported yet"
                )))
            }
        };
    let node = editor.snapshot()?.get(&path, 2)?;
    save_editor(&mut editor, parsed.output.as_deref()).await?;
    let output_path = editor.package().path().to_path_buf();
    let in_place = output_path == source_path;

    Ok(CommandOutput::success(
        format!("Added {path} and saved '{}'.", output_path.display()),
        serde_json::json!({
            "operation": operation,
            "changed": true,
            "parent": parent,
            "path": path,
            "node": node,
            "createdImage": created_image,
            "kind": editor.package().kind(),
            "outputPath": output_path,
            "inPlace": in_place,
            "revision": editor.package().source_revision()
        }),
    ))
}

fn validate_options(node_type: &str, parsed: &ParsedArguments) -> UseResult<()> {
    let is_picture = matches!(node_type, "picture" | "image" | "img");
    let is_hyperlink = matches!(node_type, "hyperlink" | "link");
    let is_comment = matches!(node_type, "comment" | "note-comment");
    let is_data_validation = matches!(
        node_type,
        "data-validation" | "datavalidation" | "validation"
    );
    let is_conditional_format = matches!(
        node_type,
        "conditional-format" | "conditional-formatting" | "conditionalformatting" | "cf"
    );
    let is_auto_filter = matches!(node_type, "auto-filter" | "autofilter" | "filter");
    let is_named_range = matches!(
        node_type,
        "named-range" | "namedrange" | "defined-name" | "definedname"
    );
    let is_table = matches!(node_type, "table" | "tbl");
    let accepts_rows = matches!(node_type, "table" | "tbl");
    let accepts_columns = matches!(node_type, "table" | "tbl" | "row" | "tr");
    let accepts_index = matches!(node_type, "column" | "col");
    let accepts_name =
        is_picture || is_named_range || is_table || matches!(node_type, "sheet" | "worksheet");
    let accepts_text = matches!(
        node_type,
        "paragraph"
            | "p"
            | "column"
            | "col"
            | "cell"
            | "tc"
            | "slide"
            | "shape"
            | "hyperlink"
            | "link"
            | "comment"
            | "note-comment"
    ) || is_conditional_format;
    if parsed.rows.is_some() && !accepts_rows {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept --rows"
        )));
    }
    if parsed.columns.is_some() && !accepts_columns {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept --columns"
        )));
    }
    if parsed.index.is_some() && !accepts_index {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept --index"
        )));
    }
    if parsed.name.is_some() && !accepts_name {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept --name"
        )));
    }
    if parsed.text.is_some() && !accepts_text {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept --text"
        )));
    }
    if parsed.has_data_validation_specific_options() && !is_data_validation {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept data-validation options"
        )));
    }
    if parsed.has_shared_rule_options()
        && !is_data_validation
        && !is_conditional_format
        && !is_table
        && !is_auto_filter
    {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept --range, --operator, --formula1, or --formula2"
        )));
    }
    if (is_table || is_auto_filter)
        && (parsed.validation_operator.is_some()
            || parsed.validation_formula1.is_some()
            || parsed.validation_formula2.is_some())
    {
        return Err(usage_error(
            "native Spreadsheet table and AutoFilter add accept --range but not --operator, --formula1, or --formula2",
        ));
    }
    if parsed.has_conditional_format_options() && !is_conditional_format {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept conditional-format options"
        )));
    }
    if parsed.has_spreadsheet_table_specific_options() && !is_table {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept Spreadsheet table options"
        )));
    }
    if parsed.has_spreadsheet_filter_specific_options() && !is_table && !is_auto_filter {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept Spreadsheet filter options"
        )));
    }
    for (present, option) in [
        (parsed.formula.is_some(), "--formula"),
        (parsed.fill.is_some(), "--fill"),
        (parsed.text_color.is_some(), "--text-color"),
        (parsed.bold.is_some(), "--bold"),
    ] {
        if present && !is_conditional_format {
            return Err(usage_error(format!(
                "native Office add type '{node_type}' does not accept {option}"
            )));
        }
    }
    if parsed.has_named_range_options() && !is_named_range && parsed.name.is_none() {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept named-range options"
        )));
    }
    for (present, option) in [
        (parsed.named_range_ref.is_some(), "--ref"),
        (parsed.named_range_scope.is_some(), "--scope"),
        (parsed.named_range_comment.is_some(), "--comment"),
        (parsed.named_range_volatile.is_some(), "--volatile"),
    ] {
        if present && !is_named_range {
            return Err(usage_error(format!(
                "native Office add type '{node_type}' does not accept {option}"
            )));
        }
    }
    if is_hyperlink && parsed.text.is_some() && parsed.display.is_some() {
        return Err(usage_error(
            "native hyperlink add accepts at most one of --text or --display",
        ));
    }
    if parsed.input.is_some() && !is_picture {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept --input"
        )));
    }
    if parsed.alt.is_some() && !is_picture {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept --alt"
        )));
    }
    if parsed.width.is_some() && !is_picture {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept --width"
        )));
    }
    if parsed.height.is_some() && !is_picture {
        return Err(usage_error(format!(
            "native Office add type '{node_type}' does not accept --height"
        )));
    }
    for (present, option) in [
        (parsed.url.is_some(), "--url"),
        (parsed.location.is_some(), "--location"),
        (parsed.display.is_some(), "--display"),
        (parsed.tooltip.is_some(), "--tooltip"),
    ] {
        if present && !is_hyperlink {
            return Err(usage_error(format!(
                "native Office add type '{node_type}' does not accept {option}"
            )));
        }
    }
    for (present, option) in [
        (parsed.author.is_some(), "--author"),
        (parsed.initials.is_some(), "--initials"),
        (parsed.x_emu.is_some(), "--x-emu"),
        (parsed.y_emu.is_some(), "--y-emu"),
    ] {
        if present && !is_comment {
            return Err(usage_error(format!(
                "native Office add type '{node_type}' does not accept {option}"
            )));
        }
    }
    Ok(())
}
