mod spreadsheet;
mod text;
mod word;

use a3s_use_core::UseResult;

use super::image::{resolve, write_data_url};
use super::output::BoundedOutput;
use crate::{DocumentKind, DocumentNode, NativeOfficeDocument, OfficeNodeType};

const CANVAS_WIDTH: f64 = 1_200.0;
const SLIDE_GAP: f64 = 56.0;
const LABEL_HEIGHT: f64 = 36.0;
const DEFAULT_SLIDE_WIDTH_CM: f64 = 33.8667;
const DEFAULT_SLIDE_HEIGHT_CM: f64 = 19.05;

pub(super) fn render(document: &NativeOfficeDocument, limit: usize) -> UseResult<String> {
    match document.kind() {
        DocumentKind::Word => word::render(document, limit),
        DocumentKind::Spreadsheet => spreadsheet::render(document, limit),
        DocumentKind::Presentation => render_presentation(document, limit),
    }
}

pub(super) fn render_presentation(
    document: &NativeOfficeDocument,
    limit: usize,
) -> UseResult<String> {
    let mut output = BoundedOutput::new(limit);
    let (width_cm, height_cm) = slide_size(document.root());
    let canvas_height = CANVAS_WIDTH * height_cm / width_cm;
    let slides = document
        .root()
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Slide)
        .collect::<Vec<_>>();
    let total_height = if slides.is_empty() {
        canvas_height
    } else {
        slides.len() as f64 * (LABEL_HEIGHT + canvas_height)
            + slides.len().saturating_sub(1) as f64 * SLIDE_GAP
    };
    output.push("<?xml version=\"1.0\" encoding=\"UTF-8\"?><svg xmlns=\"http://www.w3.org/2000/svg\" role=\"img\" aria-labelledby=\"a3s-title a3s-desc\" data-renderer=\"a3s-office-semantic-v1\" data-document-kind=\"presentation\" viewBox=\"0 0 ")?;
    output.push_fmt(format_args!("{CANVAS_WIDTH:.4} {total_height:.4}"))?;
    output.push("\"><title id=\"a3s-title\">A3S Native Office presentation semantic preview</title><desc id=\"a3s-desc\">Deterministic semantic preview; it does not claim Microsoft Office layout fidelity.</desc><style>text{font-family:ui-sans-serif,system-ui,sans-serif;fill:#172033}.slide-label{font-size:22px;font-weight:700}.shape-text{font-size:18px}.table-text{font-size:13px}.placeholder-text{font-size:16px;fill:#526079}.slide-background{fill:#fff;stroke:#8491a5;stroke-width:2}.shape{stroke:#667085;stroke-width:1.5}.placeholder{stroke-dasharray:8 5}.chart{fill:#f3f6fa;stroke:#8491a5;stroke-dasharray:8 5}.connector{stroke:#667085;stroke-width:2}.missing-image{fill:#f8fafc;stroke:#8491a5;stroke-dasharray:8 5}</style>")?;

    if slides.is_empty() {
        output.push_fmt(format_args!("<rect class=\"slide-background\" x=\"0\" y=\"0\" width=\"{CANVAS_WIDTH:.4}\" height=\"{canvas_height:.4}\"/><text class=\"placeholder-text\" x=\"40\" y=\"60\">No slides</text>"))?;
    }

    for (index, slide) in slides.into_iter().enumerate() {
        let top = index as f64 * (LABEL_HEIGHT + canvas_height + SLIDE_GAP);
        let canvas_top = top + LABEL_HEIGHT;
        let clip_id = format!("a3s-slide-clip-{}", index + 1);
        output.push_fmt(format_args!("<text class=\"slide-label\" x=\"0\" y=\"{:.4}\">Slide {}</text><defs><clipPath id=\"{clip_id}\"><rect x=\"0\" y=\"{canvas_top:.4}\" width=\"{CANVAS_WIDTH:.4}\" height=\"{canvas_height:.4}\"/></clipPath></defs><g", top + 25.0, index + 1))?;
        write_path(&mut output, slide)?;
        output.push_fmt(format_args!(" clip-path=\"url(#{clip_id})\"><rect class=\"slide-background\" x=\"0\" y=\"{canvas_top:.4}\" width=\"{CANVAS_WIDTH:.4}\" height=\"{canvas_height:.4}\"/>"))?;
        if let Some(notes) = slide
            .children
            .iter()
            .find(|node| node.node_type == OfficeNodeType::Notes && !node.text.is_empty())
        {
            output.push("<desc>")?;
            output.text(&notes.text)?;
            output.push("</desc>")?;
        }
        let owner = slide.format.get("part").map(String::as_str);
        let mut ordinal = 0_usize;
        for child in &slide.children {
            if child.node_type == OfficeNodeType::Notes {
                continue;
            }
            render_object(
                document,
                &mut output,
                child,
                owner,
                width_cm,
                height_cm,
                canvas_top,
                canvas_height,
                ordinal,
            )?;
            ordinal += 1;
        }
        output.push("</g>")?;
    }
    output.push("</svg>")?;
    Ok(output.into_string())
}

#[allow(clippy::too_many_arguments)]
fn render_object(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    node: &DocumentNode,
    owner: Option<&str>,
    slide_width_cm: f64,
    slide_height_cm: f64,
    canvas_top: f64,
    canvas_height: f64,
    ordinal: usize,
) -> UseResult<()> {
    let geometry = geometry(
        node,
        slide_width_cm,
        slide_height_cm,
        canvas_top,
        canvas_height,
        ordinal,
    );
    match node.node_type {
        OfficeNodeType::Shape | OfficeNodeType::Placeholder => render_shape(output, node, geometry),
        OfficeNodeType::Picture => render_picture(document, output, node, owner, geometry),
        OfficeNodeType::Table => render_table(output, node, geometry),
        OfficeNodeType::Chart => render_placeholder(output, node, geometry, "Chart"),
        OfficeNodeType::Connector => render_connector(output, node, geometry),
        OfficeNodeType::Group => {
            output.push("<g")?;
            write_path(output, node)?;
            output.push(">")?;
            for (index, child) in node.children.iter().enumerate() {
                render_object(
                    document,
                    output,
                    child,
                    owner,
                    slide_width_cm,
                    slide_height_cm,
                    canvas_top,
                    canvas_height,
                    ordinal.saturating_add(index),
                )?;
            }
            output.push("</g>")
        }
        _ => Ok(()),
    }
}

fn render_shape(
    output: &mut BoundedOutput,
    node: &DocumentNode,
    geometry: Geometry,
) -> UseResult<()> {
    let class = if node.node_type == OfficeNodeType::Placeholder {
        "shape placeholder"
    } else {
        "shape"
    };
    output.push("<g")?;
    write_path(output, node)?;
    write_rotation(output, node, geometry)?;
    output.push("><rect class=\"")?;
    output.push(class)?;
    output.push("\" fill=\"")?;
    output.push(
        &safe_color(node.format.get("fill").map(String::as_str)).unwrap_or_else(|| {
            if node.node_type == OfficeNodeType::Placeholder {
                "#F8FAFC".to_string()
            } else {
                "#FFFFFF".to_string()
            }
        }),
    )?;
    output.push("\"")?;
    write_rect_geometry(output, geometry)?;
    output.push("/>")?;
    let font_size = if node
        .format
        .get("title")
        .is_some_and(|value| value == "true")
    {
        28.0
    } else {
        18.0
    };
    write_text_lines(
        output,
        &node.text,
        geometry.x + 10.0,
        geometry.y + font_size + 4.0,
        font_size,
        "shape-text",
    )?;
    output.push("</g>")
}

fn render_picture(
    document: &NativeOfficeDocument,
    output: &mut BoundedOutput,
    node: &DocumentNode,
    owner: Option<&str>,
    geometry: Geometry,
) -> UseResult<()> {
    output.push("<g")?;
    write_path(output, node)?;
    write_rotation(output, node, geometry)?;
    output.push(">")?;
    if let Some(image) = resolve(document, node, owner)? {
        output.push("<image preserveAspectRatio=\"xMidYMid meet\" x=\"")?;
        output.push_fmt(format_args!(
            "{:.4}\" y=\"{:.4}\" width=\"{:.4}\" height=\"{:.4}\" href=\"",
            geometry.x, geometry.y, geometry.width, geometry.height
        ))?;
        write_data_url(output, &image)?;
        output.push("\"><title>")?;
        output.text(
            node.format
                .get("alt")
                .or_else(|| node.format.get("name"))
                .map_or("Embedded image", String::as_str),
        )?;
        output.push("</title></image>")?;
    } else {
        output.push("<rect class=\"missing-image\"")?;
        write_rect_geometry(output, geometry)?;
        output.push("/><text class=\"placeholder-text\"")?;
        output.push_fmt(format_args!(
            " x=\"{:.4}\" y=\"{:.4}\">Embedded image unavailable</text>",
            geometry.x + 10.0,
            geometry.y + 24.0
        ))?;
    }
    output.push("</g>")
}

fn render_table(
    output: &mut BoundedOutput,
    node: &DocumentNode,
    geometry: Geometry,
) -> UseResult<()> {
    output.push("<g")?;
    write_path(output, node)?;
    output.push(">")?;
    let row_count = node.children.len().max(1);
    let row_height = geometry.height / row_count as f64;
    for (row_index, row) in node.children.iter().enumerate() {
        let column_count = row.children.len().max(1);
        let column_width = geometry.width / column_count as f64;
        for (column_index, cell) in row.children.iter().enumerate() {
            let cell_geometry = Geometry {
                x: geometry.x + column_index as f64 * column_width,
                y: geometry.y + row_index as f64 * row_height,
                width: column_width,
                height: row_height,
            };
            output.push("<g")?;
            write_path(output, cell)?;
            output.push("><rect class=\"shape\" fill=\"#FFFFFF\"")?;
            write_rect_geometry(output, cell_geometry)?;
            output.push("/>")?;
            write_text_lines(
                output,
                &cell.text,
                cell_geometry.x + 6.0,
                cell_geometry.y + 16.0,
                13.0,
                "table-text",
            )?;
            output.push("</g>")?;
        }
    }
    output.push("</g>")
}

fn render_placeholder(
    output: &mut BoundedOutput,
    node: &DocumentNode,
    geometry: Geometry,
    label: &str,
) -> UseResult<()> {
    output.push("<g")?;
    write_path(output, node)?;
    output.push("><rect class=\"chart\"")?;
    write_rect_geometry(output, geometry)?;
    output.push("/><text class=\"placeholder-text\"")?;
    output.push_fmt(format_args!(
        " x=\"{:.4}\" y=\"{:.4}\">",
        geometry.x + 10.0,
        geometry.y + 24.0
    ))?;
    output.text(label)?;
    output.push("</text></g>")
}

fn render_connector(
    output: &mut BoundedOutput,
    node: &DocumentNode,
    geometry: Geometry,
) -> UseResult<()> {
    output.push("<line class=\"connector\"")?;
    write_path(output, node)?;
    output.push_fmt(format_args!(
        " x1=\"{:.4}\" y1=\"{:.4}\" x2=\"{:.4}\" y2=\"{:.4}\"/>",
        geometry.x,
        geometry.y,
        geometry.x + geometry.width,
        geometry.y + geometry.height
    ))
}

fn write_text_lines(
    output: &mut BoundedOutput,
    text: &str,
    x: f64,
    y: f64,
    font_size: f64,
    class: &str,
) -> UseResult<()> {
    if text.is_empty() {
        return Ok(());
    }
    output.push("<text class=\"")?;
    output.push(class)?;
    output.push_fmt(format_args!(
        "\" font-size=\"{font_size:.2}\" x=\"{x:.4}\" y=\"{y:.4}\">"
    ))?;
    for (index, line) in text.split('\n').enumerate() {
        output.push_fmt(format_args!(
            "<tspan x=\"{x:.4}\" dy=\"{}\">",
            if index == 0 { 0.0 } else { font_size * 1.2 }
        ))?;
        output.text(line)?;
        output.push("</tspan>")?;
    }
    output.push("</text>")
}

fn write_path(output: &mut BoundedOutput, node: &DocumentNode) -> UseResult<()> {
    output.push(" data-path=\"")?;
    output.attribute(&node.path)?;
    output.push("\" data-node-type=\"")?;
    output.push(node.node_type.label())?;
    output.push("\"")
}

fn write_rect_geometry(output: &mut BoundedOutput, geometry: Geometry) -> UseResult<()> {
    output.push_fmt(format_args!(
        " x=\"{:.4}\" y=\"{:.4}\" width=\"{:.4}\" height=\"{:.4}\"",
        geometry.x, geometry.y, geometry.width, geometry.height
    ))
}

fn write_rotation(
    output: &mut BoundedOutput,
    node: &DocumentNode,
    geometry: Geometry,
) -> UseResult<()> {
    let Some(rotation) = degrees(node.format.get("rotation").map(String::as_str)) else {
        return Ok(());
    };
    output.push_fmt(format_args!(
        " transform=\"rotate({rotation:.4} {:.4} {:.4})\"",
        geometry.x + geometry.width / 2.0,
        geometry.y + geometry.height / 2.0
    ))
}

fn geometry(
    node: &DocumentNode,
    slide_width_cm: f64,
    slide_height_cm: f64,
    canvas_top: f64,
    canvas_height: f64,
    ordinal: usize,
) -> Geometry {
    let x = centimeters(node.format.get("x").map(String::as_str));
    let y = centimeters(node.format.get("y").map(String::as_str));
    let width = positive_centimeters(node.format.get("width").map(String::as_str));
    let height = positive_centimeters(node.format.get("height").map(String::as_str));
    match (x, y, width, height) {
        (Some(x), Some(y), Some(width), Some(height)) => Geometry {
            x: x / slide_width_cm * CANVAS_WIDTH,
            y: canvas_top + y / slide_height_cm * canvas_height,
            width: width / slide_width_cm * CANVAS_WIDTH,
            height: height / slide_height_cm * canvas_height,
        },
        _ => Geometry {
            x: 40.0,
            y: canvas_top + 40.0 + ordinal as f64 * 54.0,
            width: CANVAS_WIDTH - 80.0,
            height: 44.0,
        },
    }
}

fn slide_size(root: &DocumentNode) -> (f64, f64) {
    let width = positive_centimeters(root.format.get("slideWidth").map(String::as_str))
        .unwrap_or(DEFAULT_SLIDE_WIDTH_CM);
    let height = positive_centimeters(root.format.get("slideHeight").map(String::as_str))
        .unwrap_or(DEFAULT_SLIDE_HEIGHT_CM);
    (width, height)
}

fn centimeters(value: Option<&str>) -> Option<f64> {
    let value = value?.trim().strip_suffix("cm")?.parse::<f64>().ok()?;
    value
        .is_finite()
        .then_some(value)
        .filter(|value| (-10_000.0..=10_000.0).contains(value))
}

fn positive_centimeters(value: Option<&str>) -> Option<f64> {
    centimeters(value).filter(|value| *value > 0.0)
}

fn degrees(value: Option<&str>) -> Option<f64> {
    let value = value?.trim().strip_suffix("deg")?.parse::<f64>().ok()?;
    value
        .is_finite()
        .then_some(value)
        .filter(|value| (-3600.0..=3600.0).contains(value))
}

fn safe_color(value: Option<&str>) -> Option<String> {
    let value = value?.trim().trim_start_matches('#');
    if matches!(value.len(), 6 | 8) && value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        Some(format!("#{}", value.to_ascii_uppercase()))
    } else {
        None
    }
}

#[derive(Debug, Clone, Copy)]
struct Geometry {
    x: f64,
    y: f64,
    width: f64,
    height: f64,
}
