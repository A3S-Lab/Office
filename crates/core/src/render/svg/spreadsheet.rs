use a3s_use_core::UseResult;

use super::text::wrap;
use super::write_path;
use crate::render::image::{resolve, write_data_url};
use crate::render::output::BoundedOutput;
use crate::{DocumentNode, NativeOfficeDocument, OfficeNodeType};

const CANVAS_WIDTH: f64 = 1_100.0;
const PAGE_MARGIN: f64 = 28.0;
const SHEET_PADDING: f64 = 20.0;
const SHEET_LABEL_HEIGHT: f64 = 38.0;
const SHEET_GAP: f64 = 30.0;
const ROW_LABEL_HEIGHT: f64 = 25.0;
const ROW_GAP: f64 = 14.0;
const CELL_GAP: f64 = 8.0;
const TEXT_LINE_HEIGHT: f64 = 18.0;
const FORMULA_LINE_HEIGHT: f64 = 16.0;
const PICTURE_HEIGHT: f64 = 160.0;
const MAX_CELL_CHARACTERS: usize = 102;
const MAX_FORMULA_CHARACTERS: usize = 112;

pub(super) fn render(document: &NativeOfficeDocument, limit: usize) -> UseResult<String> {
    let sheets = document
        .root()
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Worksheet)
        .collect::<Vec<_>>();
    let total_height = if sheets.is_empty() {
        180.0
    } else {
        PAGE_MARGIN * 2.0
            + sheets.iter().map(|sheet| sheet_height(sheet)).sum::<f64>()
            + sheets.len().saturating_sub(1) as f64 * SHEET_GAP
    };

    let mut output = BoundedOutput::new(limit);
    output.push("<?xml version=\"1.0\" encoding=\"UTF-8\"?><svg xmlns=\"http://www.w3.org/2000/svg\" role=\"img\" aria-labelledby=\"a3s-title a3s-desc\" data-renderer=\"a3s-office-semantic-v1\" data-document-kind=\"spreadsheet\" viewBox=\"0 0 ")?;
    output.push_fmt(format_args!("{CANVAS_WIDTH:.4} {total_height:.4}"))?;
    output.push("\"><title id=\"a3s-title\">A3S Native Office Spreadsheet semantic preview</title><desc id=\"a3s-desc\">Deterministic sparse semantic preview; gaps are not expanded and it does not claim Microsoft Excel layout or print fidelity.</desc><style>text{font-family:ui-sans-serif,system-ui,sans-serif;fill:#172033}.canvas{fill:#eef1f5}.sheet{fill:#fff;stroke:#aab4c3;stroke-width:1.5}.sheet-label{font-size:20px;font-weight:700}.row-label{font-size:14px;font-weight:700;fill:#526079}.cell{fill:#fff;stroke:#aab4c3}.cell-reference{font-size:14px;font-weight:700;fill:#43506a}.cell-value{font-size:15px}.cell-formula{font-size:13px;fill:#6b4ca5}.empty-text{font-size:14px;fill:#667085;font-style:italic}.picture{fill:#f8fafc;stroke:#8491a5}.missing-image{fill:#f8fafc;stroke:#8491a5;stroke-dasharray:7 5}.picture-label{font-size:13px;fill:#526079}</style><rect class=\"canvas\" x=\"0\" y=\"0\" width=\"1100\" height=\"")?;
    output.push_fmt(format_args!("{total_height:.4}"))?;
    output.push("\"/>")?;

    if sheets.is_empty() {
        output.push("<text class=\"empty-text\" x=\"28\" y=\"60\">No worksheets</text>")?;
    }

    let mut y = PAGE_MARGIN;
    for sheet in sheets {
        let height = sheet_height(sheet);
        render_sheet(document, &mut output, sheet, y, height)?;
        y += height + SHEET_GAP;
    }
    output.push("</svg>")?;
    Ok(output.into_string())
}

fn sheet_height(sheet: &DocumentNode) -> f64 {
    let rows = rows(sheet);
    let pictures = pictures(sheet);
    let row_height = rows.iter().map(|row| row_height(row)).sum::<f64>()
        + rows.len().saturating_sub(1) as f64 * ROW_GAP;
    let picture_height =
        pictures.len() as f64 * PICTURE_HEIGHT + pictures.len().saturating_sub(1) as f64 * CELL_GAP;
    let section_gap = if rows.is_empty() || pictures.is_empty() {
        0.0
    } else {
        CELL_GAP
    };
    let content_height = if rows.is_empty() && pictures.is_empty() {
        44.0
    } else {
        row_height + section_gap + picture_height
    };
    SHEET_LABEL_HEIGHT + SHEET_PADDING * 2.0 + content_height
}

fn row_height(row: &DocumentNode) -> f64 {
    let cells = cells(row);
    ROW_LABEL_HEIGHT
        + cells.iter().map(|cell| cell_height(cell)).sum::<f64>()
        + cells.len().saturating_sub(1) as f64 * CELL_GAP
}

fn cell_height(cell: &DocumentNode) -> f64 {
    let value_lines = wrap(&cell.text, MAX_CELL_CHARACTERS).len().max(1);
    let formula_lines = cell
        .format
        .get("formula")
        .map(|formula| wrap(formula, MAX_FORMULA_CHARACTERS).len())
        .unwrap_or(0);
    (30.0 + value_lines as f64 * TEXT_LINE_HEIGHT + formula_lines as f64 * FORMULA_LINE_HEIGHT)
        .max(54.0)
}

fn render_sheet(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    sheet: &DocumentNode,
    y: f64,
    height: f64,
) -> UseResult<()> {
    output.push("<g")?;
    write_path(output, sheet)?;
    if let Some(state) = sheet.format.get("state") {
        write_data_attribute(output, "data-state", state)?;
    }
    output.push("><rect class=\"sheet\"")?;
    write_rect(
        output,
        PAGE_MARGIN,
        y,
        CANVAS_WIDTH - PAGE_MARGIN * 2.0,
        height,
    )?;
    output.push(" rx=\"8\"/><text class=\"sheet-label\"")?;
    output.push_fmt(format_args!(
        " x=\"{:.4}\" y=\"{:.4}\">Worksheet: ",
        PAGE_MARGIN + SHEET_PADDING,
        y + 27.0
    ))?;
    output.text(sheet.path.trim_start_matches('/'))?;
    output.push("</text>")?;

    let x = PAGE_MARGIN + SHEET_PADDING;
    let width = CANVAS_WIDTH - 2.0 * (PAGE_MARGIN + SHEET_PADDING);
    let mut content_y = y + SHEET_LABEL_HEIGHT + SHEET_PADDING;
    let sheet_rows = rows(sheet);
    let sheet_pictures = pictures(sheet);
    if sheet_rows.is_empty() && sheet_pictures.is_empty() {
        output.push("<text class=\"empty-text\"")?;
        output.push_fmt(format_args!(
            " x=\"{x:.4}\" y=\"{:.4}\">No observed cells or pictures.</text>",
            content_y + 24.0
        ))?;
    }

    for (index, row) in sheet_rows.iter().enumerate() {
        let height = row_height(row);
        render_row(output, row, x, content_y, width)?;
        content_y += height;
        if index + 1 < sheet_rows.len() {
            content_y += ROW_GAP;
        }
    }
    if !sheet_rows.is_empty() && !sheet_pictures.is_empty() {
        content_y += CELL_GAP;
    }
    let owner = sheet.format.get("part").map(String::as_str);
    for (index, picture) in sheet_pictures.iter().enumerate() {
        render_picture(document, output, picture, owner, x, content_y, width)?;
        content_y += PICTURE_HEIGHT;
        if index + 1 < sheet_pictures.len() {
            content_y += CELL_GAP;
        }
    }
    output.push("</g>")
}

fn render_row(
    output: &mut BoundedOutput,
    row: &DocumentNode,
    x: f64,
    y: f64,
    width: f64,
) -> UseResult<()> {
    output.push("<g")?;
    write_path(output, row)?;
    output.push("><text class=\"row-label\"")?;
    output.push_fmt(format_args!(
        " x=\"{x:.4}\" y=\"{:.4}\">Observed ",
        y + 17.0
    ))?;
    output.text(row.path.rsplit('/').next().unwrap_or(&row.path))?;
    output.push("</text>")?;
    let mut cell_y = y + ROW_LABEL_HEIGHT;
    for cell in cells(row) {
        let height = cell_height(cell);
        render_cell(output, cell, x, cell_y, width, height)?;
        cell_y += height + CELL_GAP;
    }
    output.push("</g>")
}

fn render_cell(
    output: &mut BoundedOutput,
    cell: &DocumentNode,
    x: f64,
    y: f64,
    width: f64,
    height: f64,
) -> UseResult<()> {
    output.push("<g")?;
    write_path(output, cell)?;
    for (name, key) in [
        ("data-value-type", "valueType"),
        ("data-number-format", "numberFormat"),
        ("data-formula", "formula"),
        ("data-fill", "fill"),
        ("data-border-left", "borderLeft"),
        ("data-border-left-color", "borderLeftColor"),
        ("data-border-right", "borderRight"),
        ("data-border-right-color", "borderRightColor"),
        ("data-border-top", "borderTop"),
        ("data-border-top-color", "borderTopColor"),
        ("data-border-bottom", "borderBottom"),
        ("data-border-bottom-color", "borderBottomColor"),
        ("data-border-diagonal", "borderDiagonal"),
        ("data-border-diagonal-color", "borderDiagonalColor"),
        ("data-border-diagonal-up", "borderDiagonalUp"),
        ("data-border-diagonal-down", "borderDiagonalDown"),
        ("data-horizontal-alignment", "alignment"),
        ("data-vertical-alignment", "verticalAlignment"),
        ("data-wrap-text", "wrapText"),
        ("data-text-rotation", "textRotation"),
        ("data-indent", "indent"),
        ("data-shrink-to-fit", "shrinkToFit"),
        ("data-reading-order", "readingOrder"),
        ("data-merge", "merge"),
        ("data-merge-anchor", "mergeAnchor"),
        ("data-validation", "dataValidation"),
        ("data-validation-type", "validationType"),
    ] {
        if let Some(value) = cell.format.get(key) {
            write_data_attribute(output, name, value)?;
        }
    }
    output.push("><rect class=\"cell\"")?;
    write_rect(output, x, y, width, height)?;
    output.push(" rx=\"4\"/><text class=\"cell-reference\"")?;
    output.push_fmt(format_args!(
        " x=\"{:.4}\" y=\"{:.4}\">",
        x + 10.0,
        y + 19.0
    ))?;
    output.text(cell.path.rsplit('/').next().unwrap_or(&cell.path))?;
    output.push("</text>")?;

    let value_lines = wrap(&cell.text, MAX_CELL_CHARACTERS);
    write_lines(
        output,
        &value_lines,
        x + 10.0,
        y + 40.0,
        TEXT_LINE_HEIGHT,
        "cell-value",
        "",
    )?;
    if let Some(formula) = cell.format.get("formula") {
        let lines = wrap(formula, MAX_FORMULA_CHARACTERS);
        let formula_y = y + 40.0 + value_lines.len().max(1) as f64 * TEXT_LINE_HEIGHT;
        write_lines(
            output,
            &lines,
            x + 10.0,
            formula_y,
            FORMULA_LINE_HEIGHT,
            "cell-formula",
            "=",
        )?;
    }
    output.push("</g>")
}

fn render_picture(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    node: &DocumentNode,
    owner: Option<&str>,
    x: f64,
    y: f64,
    width: f64,
) -> UseResult<()> {
    output.push("<g")?;
    write_path(output, node)?;
    output.push(">")?;
    if let Some(image) = resolve(document, node, owner)? {
        output.push("<rect class=\"picture\"")?;
        write_rect(output, x, y, width, PICTURE_HEIGHT)?;
        output.push("/><image preserveAspectRatio=\"xMidYMid meet\"")?;
        write_rect(
            output,
            x + 6.0,
            y + 6.0,
            width - 12.0,
            PICTURE_HEIGHT - 12.0,
        )?;
        output.push(" href=\"")?;
        write_data_url(output, &image)?;
        output.push("\"><title>")?;
        output.text(picture_label(node))?;
        output.push("</title></image>")?;
    } else {
        output.push("<rect class=\"missing-image\"")?;
        write_rect(output, x, y, width, PICTURE_HEIGHT)?;
        output.push("/><text class=\"picture-label\"")?;
        output.push_fmt(format_args!(
            " x=\"{:.4}\" y=\"{:.4}\">Embedded image unavailable</text>",
            x + 10.0,
            y + 24.0
        ))?;
    }
    output.push("</g>")
}

fn rows(sheet: &DocumentNode) -> Vec<&DocumentNode> {
    sheet
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Row)
        .collect()
}

fn cells(row: &DocumentNode) -> Vec<&DocumentNode> {
    row.children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Cell)
        .collect()
}

fn pictures(sheet: &DocumentNode) -> Vec<&DocumentNode> {
    sheet
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Picture)
        .collect()
}

fn picture_label(node: &DocumentNode) -> &str {
    node.format
        .get("alt")
        .or_else(|| node.format.get("name"))
        .map_or("Embedded image", String::as_str)
}

#[allow(clippy::too_many_arguments)]
fn write_lines(
    output: &mut BoundedOutput,
    lines: &[String],
    x: f64,
    y: f64,
    line_height: f64,
    class: &str,
    prefix: &str,
) -> UseResult<()> {
    if lines.is_empty() {
        return Ok(());
    }
    output.push("<text class=\"")?;
    output.push(class)?;
    output.push_fmt(format_args!("\" x=\"{x:.4}\" y=\"{y:.4}\">"))?;
    for (index, line) in lines.iter().enumerate() {
        output.push_fmt(format_args!(
            "<tspan x=\"{x:.4}\" dy=\"{:.4}\">",
            if index == 0 { 0.0 } else { line_height }
        ))?;
        if index == 0 {
            output.text(prefix)?;
        }
        output.text(line)?;
        output.push("</tspan>")?;
    }
    output.push("</text>")
}

fn write_data_attribute(output: &mut BoundedOutput, name: &str, value: &str) -> UseResult<()> {
    output.push(" ")?;
    output.push(name)?;
    output.push("=\"")?;
    output.attribute(value)?;
    output.push("\"")
}

fn write_rect(
    output: &mut BoundedOutput,
    x: f64,
    y: f64,
    width: f64,
    height: f64,
) -> UseResult<()> {
    output.push_fmt(format_args!(
        " x=\"{x:.4}\" y=\"{y:.4}\" width=\"{width:.4}\" height=\"{height:.4}\""
    ))
}
