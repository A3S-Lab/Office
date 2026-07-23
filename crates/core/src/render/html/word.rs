use a3s_use_core::UseResult;

use super::{flag, point_size, safe_color, write_node_attributes, write_optional_attribute};
use crate::render::image::{resolve, write_data_url};
use crate::render::output::BoundedOutput;
use crate::{DocumentNode, NativeOfficeDocument, OfficeNodeType};

pub(super) fn render(document: &NativeOfficeDocument, output: &mut BoundedOutput) -> UseResult<()> {
    for region in &document.root().children {
        if !matches!(
            region.node_type,
            OfficeNodeType::Body | OfficeNodeType::Header | OfficeNodeType::Footer
        ) {
            continue;
        }
        output.push("<section class=\"word-region\"")?;
        write_node_attributes(output, region)?;
        output.push("><h2>")?;
        output.push(match region.node_type {
            OfficeNodeType::Body => "Document body",
            OfficeNodeType::Header => "Header",
            OfficeNodeType::Footer => "Footer",
            _ => unreachable!("filtered Word region"),
        })?;
        output.push("</h2>")?;
        let owner = region.format.get("part").map(String::as_str);
        for child in &region.children {
            render_block(document, output, child, owner)?;
        }
        output.push("</section>")?;
    }
    Ok(())
}

fn render_block(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    node: &DocumentNode,
    owner: Option<&str>,
) -> UseResult<()> {
    match node.node_type {
        OfficeNodeType::Paragraph => render_paragraph(document, output, node, owner),
        OfficeNodeType::Table => render_table(document, output, node, owner),
        OfficeNodeType::Picture => render_picture(document, output, node, owner),
        _ => Ok(()),
    }
}

fn render_paragraph(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    node: &DocumentNode,
    owner: Option<&str>,
) -> UseResult<()> {
    output.push("<p")?;
    write_node_attributes(output, node)?;
    let style = paragraph_style(node);
    write_optional_attribute(
        output,
        "style",
        (!style.is_empty()).then_some(style.as_str()),
    )?;
    if let Some(direction) = node
        .format
        .get("direction")
        .filter(|value| matches!(value.as_str(), "ltr" | "rtl"))
    {
        write_optional_attribute(output, "dir", Some(direction))?;
    }
    output.push(">")?;
    if node.children.is_empty() {
        if node.text.is_empty() {
            output.push("<br>")?;
        } else {
            output.text(&node.text)?;
        }
    } else {
        for child in &node.children {
            render_inline(document, output, child, owner)?;
        }
    }
    output.push("</p>")
}

fn render_inline(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    node: &DocumentNode,
    owner: Option<&str>,
) -> UseResult<()> {
    match node.node_type {
        OfficeNodeType::Picture => render_picture(document, output, node, owner),
        OfficeNodeType::Hyperlink => {
            output.push("<span class=\"hyperlink\"")?;
            write_node_attributes(output, node)?;
            output.push(">")?;
            for child in &node.children {
                render_inline(document, output, child, owner)?;
            }
            if node.children.is_empty() {
                output.text(&node.text)?;
            }
            output.push("</span>")
        }
        _ => render_run(output, node),
    }
}

fn render_run(output: &mut BoundedOutput, node: &DocumentNode) -> UseResult<()> {
    output.push("<span class=\"run")?;
    for (key, class) in [
        ("bold", " is-bold"),
        ("italic", " is-italic"),
        ("strike", " is-strike"),
        ("hidden", " is-hidden"),
    ] {
        if flag(node, key) {
            output.push(class)?;
        }
    }
    output.push("\"")?;
    write_node_attributes(output, node)?;
    write_optional_attribute(
        output,
        "lang",
        node.format.get("language").map(String::as_str),
    )?;
    write_optional_attribute(
        output,
        "data-font",
        node.format.get("font").map(String::as_str),
    )?;
    let style = run_style(node);
    write_optional_attribute(
        output,
        "style",
        (!style.is_empty()).then_some(style.as_str()),
    )?;
    output.push(">")?;
    output.text(&node.text)?;
    output.push("</span>")
}

fn render_table(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    table: &DocumentNode,
    owner: Option<&str>,
) -> UseResult<()> {
    output.push("<table")?;
    write_node_attributes(output, table)?;
    output.push("><tbody>")?;
    for row in &table.children {
        output.push("<tr")?;
        write_node_attributes(output, row)?;
        output.push(">")?;
        for cell in &row.children {
            output.push("<td")?;
            write_node_attributes(output, cell)?;
            output.push(">")?;
            for block in &cell.children {
                render_block(document, output, block, owner)?;
            }
            if cell.children.is_empty() {
                output.text(&cell.text)?;
            }
            output.push("</td>")?;
        }
        output.push("</tr>")?;
    }
    output.push("</tbody></table>")
}

fn render_picture(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    node: &DocumentNode,
    owner: Option<&str>,
) -> UseResult<()> {
    output.push("<span class=\"picture-inline\"")?;
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
        let width = requested_dimension(node, "widthPx", image.width_px);
        let height = requested_dimension(node, "heightPx", image.height_px);
        output.push_fmt(format_args!(" width=\"{width}\" height=\"{height}\">"))?;
    } else {
        output.push("<span class=\"image-placeholder\" role=\"img\" aria-label=\"Unavailable embedded image\">Embedded image unavailable</span>")?;
    }
    output.push("</span>")
}

fn paragraph_style(node: &DocumentNode) -> String {
    match node.format.get("alignment").map(String::as_str) {
        Some("center" | "ctr") => "text-align:center".into(),
        Some("right" | "r" | "end") => "text-align:right".into(),
        Some("both" | "just" | "justify") => "text-align:justify".into(),
        _ => String::new(),
    }
}

fn run_style(node: &DocumentNode) -> String {
    let mut style = String::new();
    if let Some(size) = point_size(node.format.get("size").map(String::as_str)) {
        style.push_str(&format!("font-size:{size:.2}pt;"));
    }
    if let Some(color) = safe_color(node.format.get("color").map(String::as_str)) {
        style.push_str("color:");
        style.push_str(&color);
    }
    style
}

fn requested_dimension(node: &DocumentNode, key: &str, fallback: u32) -> u32 {
    node.format
        .get(key)
        .and_then(|value| value.parse::<u32>().ok())
        .filter(|value| *value > 0 && *value <= 100_000)
        .unwrap_or(fallback)
}
