use a3s_use_core::UseResult;

use super::{flag, point_size, safe_color, write_node_attributes, write_optional_attribute};
use crate::render::image::{resolve, write_data_url};
use crate::render::output::BoundedOutput;
use crate::{DocumentNode, NativeOfficeDocument, OfficeNodeType};

const DEFAULT_SLIDE_WIDTH_CM: f64 = 33.8667;
const DEFAULT_SLIDE_HEIGHT_CM: f64 = 19.05;

pub(super) fn render(document: &NativeOfficeDocument, output: &mut BoundedOutput) -> UseResult<()> {
    let (slide_width, slide_height) = slide_size(document.root());
    if document.root().children.is_empty() {
        return output.push(
            "<section class=\"slide-card\"><h2>Presentation</h2><p>No slides.</p></section>",
        );
    }
    for (index, slide) in document.root().children.iter().enumerate() {
        if slide.node_type != OfficeNodeType::Slide {
            continue;
        }
        output.push("<section class=\"slide-card\"")?;
        write_node_attributes(output, slide)?;
        output.push_fmt(format_args!("><h2>Slide {}</h2><div class=\"slide-canvas\" style=\"aspect-ratio:{slide_width:.4}/{slide_height:.4}\">", index + 1))?;
        let owner = slide.format.get("part").map(String::as_str);
        let mut object_count = 0_usize;
        for child in &slide.children {
            if child.node_type == OfficeNodeType::Notes {
                continue;
            }
            object_count += 1;
            render_object(document, output, child, owner, slide_width, slide_height)?;
        }
        if object_count == 0 {
            output.push("<div class=\"slide-empty\">Empty slide</div>")?;
        }
        output.push("</div>")?;
        if let Some(notes) = slide
            .children
            .iter()
            .find(|node| node.node_type == OfficeNodeType::Notes && !node.text.is_empty())
        {
            output.push("<aside class=\"slide-notes\"")?;
            write_node_attributes(output, notes)?;
            output.push("><strong>Notes</strong><br>")?;
            output.text(&notes.text)?;
            output.push("</aside>")?;
        }
        output.push("</section>")?;
    }
    Ok(())
}

fn render_object(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    node: &DocumentNode,
    owner: Option<&str>,
    slide_width: f64,
    slide_height: f64,
) -> UseResult<()> {
    match node.node_type {
        OfficeNodeType::Picture => {
            render_picture(document, output, node, owner, slide_width, slide_height)
        }
        OfficeNodeType::Table => render_table(output, node, slide_width, slide_height),
        OfficeNodeType::Shape | OfficeNodeType::Placeholder => {
            render_shape(output, node, slide_width, slide_height)
        }
        OfficeNodeType::Group => {
            start_object(output, node, "group", slide_width, slide_height)?;
            for child in &node.children {
                render_object(document, output, child, owner, slide_width, slide_height)?;
            }
            output.push("</div>")
        }
        OfficeNodeType::Chart | OfficeNodeType::Connector => {
            let class = if node.node_type == OfficeNodeType::Chart {
                "chart"
            } else {
                "connector"
            };
            start_object(output, node, class, slide_width, slide_height)?;
            output.push(if node.node_type == OfficeNodeType::Chart {
                "Chart"
            } else {
                "Connector"
            })?;
            output.push("</div>")
        }
        _ => Ok(()),
    }
}

fn render_shape(
    output: &mut BoundedOutput,
    node: &DocumentNode,
    slide_width: f64,
    slide_height: f64,
) -> UseResult<()> {
    let class = if node.node_type == OfficeNodeType::Placeholder {
        "placeholder"
    } else {
        "shape"
    };
    start_object(output, node, class, slide_width, slide_height)?;
    let paragraphs = node
        .children
        .iter()
        .filter(|child| child.node_type == OfficeNodeType::Paragraph)
        .collect::<Vec<_>>();
    if paragraphs.is_empty() {
        output.text(&node.text)?;
    } else {
        for paragraph in paragraphs {
            output.push("<p")?;
            write_node_attributes(output, paragraph)?;
            output.push(">")?;
            if paragraph.children.is_empty() {
                output.text(&paragraph.text)?;
            } else {
                for run in &paragraph.children {
                    output.push("<span class=\"run")?;
                    if flag(run, "bold") {
                        output.push(" is-bold")?;
                    }
                    if flag(run, "italic") {
                        output.push(" is-italic")?;
                    }
                    output.push("\"")?;
                    write_node_attributes(output, run)?;
                    let style = text_style(run);
                    write_optional_attribute(
                        output,
                        "style",
                        (!style.is_empty()).then_some(style.as_str()),
                    )?;
                    output.push(">")?;
                    output.text(&run.text)?;
                    output.push("</span>")?;
                }
            }
            output.push("</p>")?;
        }
    }
    output.push("</div>")
}

fn render_table(
    output: &mut BoundedOutput,
    node: &DocumentNode,
    slide_width: f64,
    slide_height: f64,
) -> UseResult<()> {
    start_object(output, node, "table", slide_width, slide_height)?;
    output.push("<table><tbody>")?;
    for row in &node.children {
        output.push("<tr")?;
        write_node_attributes(output, row)?;
        output.push(">")?;
        for cell in &row.children {
            output.push("<td")?;
            write_node_attributes(output, cell)?;
            output.push(">")?;
            output.text(&cell.text)?;
            output.push("</td>")?;
        }
        output.push("</tr>")?;
    }
    output.push("</tbody></table></div>")
}

fn render_picture(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    node: &DocumentNode,
    owner: Option<&str>,
    slide_width: f64,
    slide_height: f64,
) -> UseResult<()> {
    start_object(output, node, "picture", slide_width, slide_height)?;
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
        output.push("\">")?;
    } else {
        output.push("<span class=\"image-placeholder\" role=\"img\" aria-label=\"Unavailable embedded image\">Embedded image unavailable</span>")?;
    }
    output.push("</div>")
}

fn start_object(
    output: &mut BoundedOutput,
    node: &DocumentNode,
    class: &str,
    slide_width: f64,
    slide_height: f64,
) -> UseResult<()> {
    let style = position_style(node, slide_width, slide_height);
    output.push("<div class=\"slide-object ")?;
    output.push(class)?;
    if style.is_none() {
        output.push(" unpositioned")?;
    }
    output.push("\"")?;
    write_node_attributes(output, node)?;
    write_optional_attribute(output, "style", style.as_deref())?;
    output.push(">")
}

fn position_style(node: &DocumentNode, slide_width: f64, slide_height: f64) -> Option<String> {
    let x = centimeters(node.format.get("x").map(String::as_str))?;
    let y = centimeters(node.format.get("y").map(String::as_str))?;
    let width = centimeters(node.format.get("width").map(String::as_str))?;
    let height = centimeters(node.format.get("height").map(String::as_str))?;
    let mut style = format!(
        "left:{:.4}%;top:{:.4}%;width:{:.4}%;height:{:.4}%;",
        x / slide_width * 100.0,
        y / slide_height * 100.0,
        width / slide_width * 100.0,
        height / slide_height * 100.0
    );
    if let Some(fill) = safe_color(node.format.get("fill").map(String::as_str)) {
        style.push_str("background-color:");
        style.push_str(&fill);
        style.push(';');
    }
    if let Some(rotation) = degrees(node.format.get("rotation").map(String::as_str)) {
        style.push_str(&format!("transform:rotate({rotation:.4}deg);"));
    }
    Some(style)
}

fn text_style(node: &DocumentNode) -> String {
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

pub(super) fn slide_size(root: &DocumentNode) -> (f64, f64) {
    let width = centimeters(root.format.get("slideWidth").map(String::as_str))
        .unwrap_or(DEFAULT_SLIDE_WIDTH_CM);
    let height = centimeters(root.format.get("slideHeight").map(String::as_str))
        .unwrap_or(DEFAULT_SLIDE_HEIGHT_CM);
    (width, height)
}

pub(super) fn centimeters(value: Option<&str>) -> Option<f64> {
    let value = value?.trim().strip_suffix("cm")?.parse::<f64>().ok()?;
    value
        .is_finite()
        .then_some(value)
        .filter(|value| *value >= 0.0 && *value <= 10_000.0)
}

pub(super) fn degrees(value: Option<&str>) -> Option<f64> {
    let value = value?.trim().strip_suffix("deg")?.parse::<f64>().ok()?;
    value
        .is_finite()
        .then_some(value)
        .filter(|value| (-3600.0..=3600.0).contains(value))
}
