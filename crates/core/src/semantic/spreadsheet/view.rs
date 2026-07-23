use super::{DocumentNode, OfficeNodeType};
use crate::xml_tree::{XmlElement, XmlNode};
use crate::NativeSpreadsheetFrozenPane;

pub(super) fn read(worksheet: &XmlElement, sheet_path: &str) -> Option<DocumentNode> {
    let views = worksheet
        .children_named("sheetViews")
        .filter(|views| views.namespace == worksheet.namespace)
        .collect::<Vec<_>>();
    let collection = *views.first()?;
    let sheet_views = collection
        .children_named("sheetView")
        .filter(|view| view.namespace == worksheet.namespace)
        .collect::<Vec<_>>();
    let view = sheet_views
        .iter()
        .copied()
        .find(|view| view.attribute("workbookViewId") == Some("0"))
        .or_else(|| sheet_views.first().copied())?;
    let panes = view
        .children_named("pane")
        .filter(|pane| pane.namespace == worksheet.namespace)
        .collect::<Vec<_>>();
    let pane = *panes.first()?;
    let frozen_rows = integer_split(pane.attribute("ySplit"));
    let frozen_columns = integer_split(pane.attribute("xSplit"));
    let top_left = pane.attribute("topLeftCell").unwrap_or_default();
    let candidate = NativeSpreadsheetFrozenPane::new(
        frozen_rows.unwrap_or_default(),
        frozen_columns.unwrap_or_default(),
        top_left,
    );
    let normalized = candidate.normalized().ok();
    let expected_active = normalized
        .as_ref()
        .map(NativeSpreadsheetFrozenPane::active_pane);
    let known_attributes = pane.attributes.iter().all(|attribute| {
        attribute.namespace.is_none()
            && matches!(
                attribute.local_name.as_str(),
                "xSplit" | "ySplit" | "topLeftCell" | "activePane" | "state"
            )
    });
    let empty = pane.children.iter().all(|child| match child {
        XmlNode::Text(text) => text.trim().is_empty(),
        XmlNode::Element(_) => false,
    });
    let mutable = views.len() == 1
        && view.attribute("workbookViewId") == Some("0")
        && panes.len() == 1
        && frozen_rows.is_some()
        && frozen_columns.is_some()
        && normalized.is_some()
        && pane.attribute("state") == Some("frozen")
        && pane.attribute("activePane") == expected_active
        && known_attributes
        && empty;

    let mut node = DocumentNode::new(
        format!("{sheet_path}/freeze"),
        "freeze",
        OfficeNodeType::FrozenPane,
    );
    node.text = format!(
        "{} frozen row(s), {} frozen column(s)",
        candidate.frozen_rows, candidate.frozen_columns
    );
    node.format
        .insert("frozenRows".into(), candidate.frozen_rows.to_string());
    node.format
        .insert("frozenColumns".into(), candidate.frozen_columns.to_string());
    node.format
        .insert("topLeftCell".into(), candidate.top_left_cell);
    if let Some(active) = pane.attribute("activePane") {
        node.format.insert("activePane".into(), active.into());
    }
    if let Some(state) = pane.attribute("state") {
        node.format.insert("state".into(), state.into());
    }
    node.format
        .insert("nativeMutable".into(), mutable.to_string());
    Some(node)
}

fn integer_split(value: Option<&str>) -> Option<u32> {
    match value {
        None => Some(0),
        Some(value) => value.parse::<u32>().ok(),
    }
}
