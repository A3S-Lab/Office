use a3s_use_core::UseResult;

use super::text::wrap;
use super::write_path;
use crate::render::image::{resolve, write_data_url};
use crate::render::output::BoundedOutput;
use crate::{DocumentNode, NativeOfficeDocument, OfficeNodeType};

const CANVAS_WIDTH: f64 = 1_000.0;
const PAGE_MARGIN: f64 = 28.0;
const REGION_PADDING: f64 = 20.0;
const REGION_LABEL_HEIGHT: f64 = 36.0;
const REGION_GAP: f64 = 28.0;
const BLOCK_GAP: f64 = 12.0;
const LINE_HEIGHT: f64 = 21.0;
const TABLE_LINE_HEIGHT: f64 = 17.0;
const PICTURE_HEIGHT: f64 = 150.0;
const MAX_TEXT_CHARACTERS: usize = 88;
const MAX_CELL_CHARACTERS: usize = 30;

pub(super) fn render(document: &NativeOfficeDocument, limit: usize) -> UseResult<String> {
    let regions = document
        .root()
        .children
        .iter()
        .filter(|node| {
            matches!(
                node.node_type,
                OfficeNodeType::Body | OfficeNodeType::Header | OfficeNodeType::Footer
            )
        })
        .collect::<Vec<_>>();
    let total_height = if regions.is_empty() {
        180.0
    } else {
        PAGE_MARGIN * 2.0
            + regions
                .iter()
                .map(|region| region_height(region))
                .sum::<f64>()
            + regions.len().saturating_sub(1) as f64 * REGION_GAP
    };

    let mut output = BoundedOutput::new(limit);
    output.push("<?xml version=\"1.0\" encoding=\"UTF-8\"?><svg xmlns=\"http://www.w3.org/2000/svg\" role=\"img\" aria-labelledby=\"a3s-title a3s-desc\" data-renderer=\"a3s-office-semantic-v1\" data-document-kind=\"word\" viewBox=\"0 0 ")?;
    output.push_fmt(format_args!("{CANVAS_WIDTH:.4} {total_height:.4}"))?;
    output.push("\"><title id=\"a3s-title\">A3S Native Office Word semantic preview</title><desc id=\"a3s-desc\">Deterministic semantic preview; it does not claim Microsoft Office pagination or layout fidelity.</desc><style>text{font-family:ui-sans-serif,system-ui,sans-serif;fill:#172033}.canvas{fill:#eef1f5}.region{fill:#fff;stroke:#aab4c3;stroke-width:1.5}.region-label{font-size:20px;font-weight:700}.paragraph{fill:#fff;stroke:#d2d8e2}.paragraph-text{font-size:16px}.empty-text{font-size:14px;fill:#667085;font-style:italic}.table{fill:#fff;stroke:#8491a5}.cell{fill:#fff;stroke:#aab4c3}.cell-text{font-size:13px}.picture{fill:#f8fafc;stroke:#8491a5}.missing-image{fill:#f8fafc;stroke:#8491a5;stroke-dasharray:7 5}.picture-label{font-size:13px;fill:#526079}</style><rect class=\"canvas\" x=\"0\" y=\"0\" width=\"1000\" height=\"")?;
    output.push_fmt(format_args!("{total_height:.4}"))?;
    output.push("\"/>")?;

    if regions.is_empty() {
        output.push("<text class=\"empty-text\" x=\"28\" y=\"60\">No Word regions</text>")?;
    }

    let mut y = PAGE_MARGIN;
    for region in regions {
        let height = region_height(region);
        render_region(document, &mut output, region, y, height)?;
        y += height + REGION_GAP;
    }
    output.push("</svg>")?;
    Ok(output.into_string())
}

fn region_height(region: &DocumentNode) -> f64 {
    let blocks = region
        .children
        .iter()
        .map(block_height)
        .filter(|height| *height > 0.0)
        .collect::<Vec<_>>();
    REGION_LABEL_HEIGHT
        + REGION_PADDING * 2.0
        + blocks.iter().sum::<f64>()
        + blocks.len().saturating_sub(1) as f64 * BLOCK_GAP
}

fn block_height(node: &DocumentNode) -> f64 {
    let picture_height = picture_nodes(node).len() as f64 * (PICTURE_HEIGHT + BLOCK_GAP);
    match node.node_type {
        OfficeNodeType::Paragraph => {
            let line_count = wrap(&node.text, MAX_TEXT_CHARACTERS).len().max(1);
            20.0 + line_count as f64 * LINE_HEIGHT + picture_height
        }
        OfficeNodeType::Table => {
            let rows = node.children.iter().map(row_height).sum::<f64>();
            rows.max(44.0) + picture_height
        }
        OfficeNodeType::Picture => PICTURE_HEIGHT,
        _ => 0.0,
    }
}

fn row_height(row: &DocumentNode) -> f64 {
    row.children
        .iter()
        .map(|cell| {
            16.0 + wrap(&cell.text, MAX_CELL_CHARACTERS).len().max(1) as f64 * TABLE_LINE_HEIGHT
        })
        .fold(36.0, f64::max)
}

fn render_region(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    region: &DocumentNode,
    y: f64,
    height: f64,
) -> UseResult<()> {
    output.push("<g")?;
    write_path(output, region)?;
    output.push("><rect class=\"region\" x=\"")?;
    output.push_fmt(format_args!(
        "{PAGE_MARGIN:.4}\" y=\"{y:.4}\" width=\"{:.4}\" height=\"{height:.4}\" rx=\"8\"/>",
        CANVAS_WIDTH - PAGE_MARGIN * 2.0
    ))?;
    output.push("<text class=\"region-label\" x=\"")?;
    output.push_fmt(format_args!(
        "{:.4}\" y=\"{:.4}\">",
        PAGE_MARGIN + REGION_PADDING,
        y + 26.0
    ))?;
    output.push(region_label(region.node_type))?;
    output.push("</text>")?;

    let owner = region.format.get("part").map(String::as_str);
    let mut block_y = y + REGION_LABEL_HEIGHT + REGION_PADDING;
    for block in &region.children {
        let block_height = block_height(block);
        if block_height == 0.0 {
            continue;
        }
        render_block(document, output, block, owner, block_y, block_height)?;
        block_y += block_height + BLOCK_GAP;
    }
    output.push("</g>")
}

fn render_block(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    node: &DocumentNode,
    owner: Option<&str>,
    y: f64,
    height: f64,
) -> UseResult<()> {
    let x = PAGE_MARGIN + REGION_PADDING;
    let width = CANVAS_WIDTH - 2.0 * (PAGE_MARGIN + REGION_PADDING);
    match node.node_type {
        OfficeNodeType::Paragraph => {
            output.push("<g")?;
            write_path(output, node)?;
            output.push("><rect class=\"paragraph\"")?;
            write_rect(output, x, y, width, height)?;
            output.push(" rx=\"4\"/>")?;
            let lines = wrap(&node.text, MAX_TEXT_CHARACTERS);
            if lines.is_empty() {
                output.push("<text class=\"empty-text\"")?;
                output.push_fmt(format_args!(
                    " x=\"{:.4}\" y=\"{:.4}\">Empty paragraph</text>",
                    x + 10.0,
                    y + 24.0
                ))?;
            } else {
                write_lines(
                    output,
                    &lines,
                    x + 10.0,
                    y + 22.0,
                    LINE_HEIGHT,
                    "paragraph-text",
                )?;
            }
            let text_height = 20.0 + lines.len().max(1) as f64 * LINE_HEIGHT;
            render_pictures(
                document,
                output,
                node,
                owner,
                x + 10.0,
                y + text_height,
                width - 20.0,
            )?;
            output.push("</g>")
        }
        OfficeNodeType::Table => render_table(document, output, node, owner, x, y, width),
        OfficeNodeType::Picture => render_picture(document, output, node, owner, x, y, width),
        _ => Ok(()),
    }
}

fn render_table(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    table: &DocumentNode,
    owner: Option<&str>,
    x: f64,
    y: f64,
    width: f64,
) -> UseResult<()> {
    output.push("<g")?;
    write_path(output, table)?;
    output.push("><rect class=\"table\"")?;
    let grid_height = table.children.iter().map(row_height).sum::<f64>().max(44.0);
    write_rect(output, x, y, width, grid_height)?;
    output.push("/>")?;
    let mut row_y = y;
    for row in &table.children {
        let height = row_height(row);
        let columns = row.children.len().max(1);
        let cell_width = width / columns as f64;
        output.push("<g")?;
        write_path(output, row)?;
        output.push(">")?;
        for (index, cell) in row.children.iter().enumerate() {
            let cell_x = x + index as f64 * cell_width;
            output.push("<g")?;
            write_path(output, cell)?;
            output.push("><rect class=\"cell\"")?;
            write_rect(output, cell_x, row_y, cell_width, height)?;
            output.push("/>")?;
            let lines = wrap(&cell.text, MAX_CELL_CHARACTERS);
            write_lines(
                output,
                &lines,
                cell_x + 6.0,
                row_y + 17.0,
                TABLE_LINE_HEIGHT,
                "cell-text",
            )?;
            output.push("</g>")?;
        }
        output.push("</g>")?;
        row_y += height;
    }
    render_pictures(
        document,
        output,
        table,
        owner,
        x + 10.0,
        y + grid_height + BLOCK_GAP,
        width - 20.0,
    )?;
    output.push("</g>")
}

fn render_pictures(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    container: &DocumentNode,
    owner: Option<&str>,
    x: f64,
    mut y: f64,
    width: f64,
) -> UseResult<()> {
    for picture in picture_nodes(container) {
        render_picture(document, output, picture, owner, x, y, width)?;
        y += PICTURE_HEIGHT + BLOCK_GAP;
    }
    Ok(())
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

fn picture_nodes(node: &DocumentNode) -> Vec<&DocumentNode> {
    let mut pictures = Vec::new();
    collect_pictures(node, &mut pictures);
    pictures
}

fn collect_pictures<'a>(node: &'a DocumentNode, output: &mut Vec<&'a DocumentNode>) {
    if node.node_type == OfficeNodeType::Picture {
        output.push(node);
        return;
    }
    for child in &node.children {
        collect_pictures(child, output);
    }
}

fn picture_label(node: &DocumentNode) -> &str {
    node.format
        .get("alt")
        .or_else(|| node.format.get("name"))
        .map_or("Embedded image", String::as_str)
}

fn region_label(node_type: OfficeNodeType) -> &'static str {
    match node_type {
        OfficeNodeType::Body => "Document body",
        OfficeNodeType::Header => "Header",
        OfficeNodeType::Footer => "Footer",
        _ => "Word region",
    }
}

fn write_lines(
    output: &mut BoundedOutput,
    lines: &[String],
    x: f64,
    y: f64,
    line_height: f64,
    class: &str,
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
        output.text(line)?;
        output.push("</tspan>")?;
    }
    output.push("</text>")
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
