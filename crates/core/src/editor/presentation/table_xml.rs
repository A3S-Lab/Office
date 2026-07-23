use a3s_use_core::UseResult;

use super::{preserve_space_attribute, qualified};
use crate::xml_edit::{apply_patches, escape_text, IndexedXmlElement, XmlPatch};
use crate::LosslessXmlPart;

const TRANSITIONAL_DRAWING_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWING_NAMESPACE: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const STRICT_PRESENTATION_NAMESPACE: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const TABLE_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/table";
const DEFAULT_X_EMU: u64 = 457_200;
const DEFAULT_Y_EMU: u64 = 1_600_200;
const DEFAULT_WIDTH_EMU: u64 = 8_229_600;
const DEFAULT_ROW_HEIGHT_EMU: u64 = 370_840;

pub(super) fn graphic_frame_xml(
    id: u32,
    position: usize,
    rows: usize,
    columns: usize,
    height: u64,
    presentation_prefix: Option<&str>,
    drawing_namespace: &str,
) -> String {
    let graphic_frame = qualified(presentation_prefix, "graphicFrame");
    let non_visual = qualified(presentation_prefix, "nvGraphicFramePr");
    let non_visual_properties = qualified(presentation_prefix, "cNvPr");
    let non_visual_drawing = qualified(presentation_prefix, "cNvGraphicFramePr");
    let application_properties = qualified(presentation_prefix, "nvPr");
    let transform = qualified(presentation_prefix, "xfrm");
    format!(
        "<{graphic_frame} xmlns:a=\"{drawing_namespace}\"><{non_visual}><{non_visual_properties} id=\"{id}\" name=\"Table {position}\"/><{non_visual_drawing}/><{application_properties}/></{non_visual}><{transform}><a:off x=\"{DEFAULT_X_EMU}\" y=\"{DEFAULT_Y_EMU}\"/><a:ext cx=\"{DEFAULT_WIDTH_EMU}\" cy=\"{height}\"/></{transform}><a:graphic><a:graphicData uri=\"{TABLE_URI}\">{}</a:graphicData></a:graphic></{graphic_frame}>",
        table_xml(rows, columns, Some("a"))
    )
}

fn table_xml(rows: usize, columns: usize, drawing_prefix: Option<&str>) -> String {
    let base_width = DEFAULT_WIDTH_EMU / columns as u64;
    let remainder = DEFAULT_WIDTH_EMU % columns as u64;
    let grid_column = qualified(drawing_prefix, "gridCol");
    let grid = (0..columns)
        .map(|column| {
            let width = base_width + u64::from(column == columns - 1) * remainder;
            format!("<{grid_column} w=\"{width}\"/>")
        })
        .collect::<String>();
    let rows = (0..rows)
        .map(|_| row_xml(columns, drawing_prefix))
        .collect::<String>();
    let table = qualified(drawing_prefix, "tbl");
    let table_properties = qualified(drawing_prefix, "tblPr");
    let table_grid = qualified(drawing_prefix, "tblGrid");
    format!(
        "<{table}><{table_properties} firstRow=\"1\" bandRow=\"1\"/><{table_grid}>{grid}</{table_grid}>{rows}</{table}>"
    )
}

pub(super) fn row_xml(columns: usize, drawing_prefix: Option<&str>) -> String {
    let cells = (0..columns)
        .map(|_| cell_xml("", drawing_prefix))
        .collect::<String>();
    let row = qualified(drawing_prefix, "tr");
    format!("<{row} h=\"{DEFAULT_ROW_HEIGHT_EMU}\">{cells}</{row}>")
}

pub(super) fn grid_column_xml(width: u64, drawing_prefix: Option<&str>) -> String {
    let column = qualified(drawing_prefix, "gridCol");
    format!("<{column} w=\"{width}\"/>")
}

pub(super) fn cell_xml(text: &str, drawing_prefix: Option<&str>) -> String {
    let paragraph_tag = qualified(drawing_prefix, "p");
    let run_tag = qualified(drawing_prefix, "r");
    let run_properties_tag = qualified(drawing_prefix, "rPr");
    let text_tag = qualified(drawing_prefix, "t");
    let end_properties_tag = qualified(drawing_prefix, "endParaRPr");
    let paragraph = if text.is_empty() {
        format!("<{paragraph_tag}><{end_properties_tag} lang=\"en-US\"/></{paragraph_tag}>")
    } else {
        let escaped = escape_text(text);
        let space = preserve_space_attribute(text);
        format!(
            "<{paragraph_tag}><{run_tag}><{run_properties_tag} lang=\"en-US\"/><{text_tag}{space}>{escaped}</{text_tag}></{run_tag}><{end_properties_tag} lang=\"en-US\"/></{paragraph_tag}>"
        )
    };
    let cell = qualified(drawing_prefix, "tc");
    let text_body = qualified(drawing_prefix, "txBody");
    let body_properties = qualified(drawing_prefix, "bodyPr");
    let list_style = qualified(drawing_prefix, "lstStyle");
    let cell_properties = qualified(drawing_prefix, "tcPr");
    format!(
        "<{cell}><{text_body}><{body_properties}/><{list_style}/>{paragraph}</{text_body}><{cell_properties}/></{cell}>"
    )
}

pub(super) fn drawing_namespace(part: &LosslessXmlPart) -> &'static str {
    if part.root().namespace.as_deref() == Some(STRICT_PRESENTATION_NAMESPACE) {
        STRICT_DRAWING_NAMESPACE
    } else {
        TRANSITIONAL_DRAWING_NAMESPACE
    }
}

pub(super) fn replace_attribute_patch(
    element: &IndexedXmlElement,
    name: &str,
    value: String,
) -> XmlPatch {
    let mut attributes = element.qualified_attributes.clone();
    attributes.insert(name.to_string(), value);
    let attributes = attributes
        .into_iter()
        .map(|(name, value)| {
            format!(
                " {name}=\"{}\"",
                quick_xml::escape::escape(&value).into_owned()
            )
        })
        .collect::<String>();
    let terminator = if element.empty { "/>" } else { ">" };
    XmlPatch::new(
        element.start_tag_range.clone(),
        format!("<{}{attributes}{terminator}", element.qualified_name),
    )
}

pub(super) fn insert_ordered_child(
    part: &LosslessXmlPart,
    parent: &IndexedXmlElement,
    child: impl AsRef<[u8]>,
) -> UseResult<Vec<u8>> {
    if parent.empty {
        return crate::xml_edit::insert_child(part, parent, child);
    }
    let insertion = ordered_child_insertion(parent);
    apply_patches(
        part,
        vec![XmlPatch::new(insertion..insertion, child.as_ref().to_vec())],
    )
}

pub(super) fn ordered_child_insertion(parent: &IndexedXmlElement) -> usize {
    parent
        .children
        .iter()
        .find(|child| child.local_name == "extLst")
        .map_or(parent.content_range.end, |extension| {
            extension.full_range.start
        })
}
