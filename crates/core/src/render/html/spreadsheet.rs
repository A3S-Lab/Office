use a3s_use_core::UseResult;

use super::{write_node_attributes, write_optional_attribute};
use crate::render::image::{resolve, write_data_url};
use crate::render::output::BoundedOutput;
use crate::{DocumentNode, NativeOfficeDocument, OfficeNodeType};

pub(super) fn render(document: &NativeOfficeDocument, output: &mut BoundedOutput) -> UseResult<()> {
    for sheet in &document.root().children {
        if sheet.node_type != OfficeNodeType::Worksheet {
            continue;
        }
        output.push("<section class=\"sheet\"")?;
        write_node_attributes(output, sheet)?;
        write_optional_attribute(
            output,
            "data-state",
            sheet.format.get("state").map(String::as_str),
        )?;
        output.push("><h2>")?;
        output.text(sheet.path.trim_start_matches('/'))?;
        output.push("</h2>")?;
        render_cells(output, sheet)?;
        let owner = sheet.format.get("part").map(String::as_str);
        for picture in sheet
            .children
            .iter()
            .filter(|node| node.node_type == OfficeNodeType::Picture)
        {
            render_picture(document, output, picture, owner)?;
        }
        output.push("</section>")?;
    }
    Ok(())
}

fn render_cells(output: &mut BoundedOutput, sheet: &DocumentNode) -> UseResult<()> {
    let rows = sheet
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Row)
        .collect::<Vec<_>>();
    if rows.is_empty() {
        return output.push("<p>No observed cells.</p>");
    }
    output.push("<table class=\"sparse-grid\"><thead><tr><th>Observed row</th><th>Cells (sparse; gaps are not expanded)</th></tr></thead><tbody>")?;
    for row in rows {
        output.push("<tr")?;
        write_node_attributes(output, row)?;
        output.push("><th class=\"sheet-row-label\">")?;
        output.text(row.path.rsplit('/').next().unwrap_or(&row.path))?;
        output.push("</th><td>")?;
        for cell in row
            .children
            .iter()
            .filter(|node| node.node_type == OfficeNodeType::Cell)
        {
            render_cell(output, cell)?;
        }
        output.push("</td></tr>")?;
    }
    output.push("</tbody></table>")
}

fn render_cell(output: &mut BoundedOutput, cell: &DocumentNode) -> UseResult<()> {
    output.push("<div class=\"cell\"")?;
    write_node_attributes(output, cell)?;
    for (attribute, key) in [
        ("data-value-type", "valueType"),
        ("data-number-format", "numberFormat"),
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
        write_optional_attribute(output, attribute, cell.format.get(key).map(String::as_str))?;
    }
    output.push("><strong class=\"cell-reference\">")?;
    output.text(cell.path.rsplit('/').next().unwrap_or(&cell.path))?;
    output.push("</strong>: <span class=\"cell-value\">")?;
    output.text(&cell.text)?;
    output.push("</span>")?;
    if let Some(formula) = cell.format.get("formula") {
        output.push("<code class=\"cell-formula\">=")?;
        output.text(formula)?;
        output.push("</code>")?;
    }
    output.push("</div>")
}

fn render_picture(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    node: &DocumentNode,
    owner: Option<&str>,
) -> UseResult<()> {
    output.push("<figure class=\"sheet-picture\"")?;
    write_node_attributes(output, node)?;
    output.push(">")?;
    if let Some(image) = resolve(document, node, owner)? {
        output.push("<img src=\"")?;
        write_data_url(output, &image)?;
        output.push("\" alt=\"")?;
        output.attribute(
            node.format
                .get("alt")
                .or_else(|| node.format.get("name"))
                .map_or("", String::as_str),
        )?;
        output.push("\"")?;
        let width = dimension(node, "widthPx", image.width_px);
        let height = dimension(node, "heightPx", image.height_px);
        output.push_fmt(format_args!(" width=\"{width}\" height=\"{height}\">"))?;
    } else {
        output.push("<span class=\"image-placeholder\" role=\"img\" aria-label=\"Unavailable embedded image\">Embedded image unavailable</span>")?;
    }
    output.push("<figcaption>Anchor: ")?;
    output.text(
        node.format
            .get("anchorCell")
            .map_or("unspecified", String::as_str),
    )?;
    output.push("</figcaption></figure>")
}

fn dimension(node: &DocumentNode, key: &str, fallback: u32) -> u32 {
    node.format
        .get(key)
        .and_then(|value| value.parse::<u32>().ok())
        .filter(|value| *value > 0 && *value <= 100_000)
        .unwrap_or(fallback)
}
