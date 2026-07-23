use a3s_use_core::UseResult;

use super::{editor_error, prefix, qualified, validate_mutation_path};
use crate::xml_edit::{apply_patches, index_xml, insert_child, insert_ordered_child, XmlPatch};
use crate::{
    DocumentKind, NativeOfficeDocument, NativeOfficePackage, NativeSpreadsheetFrozenPane,
    OfficeNodeType,
};

const WORKSHEET_CHILDREN_AFTER_VIEWS: &[&str] = &[
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
];

struct ResolvedSheet {
    path: String,
    part: String,
}

pub(super) fn is_path(path: &str) -> bool {
    let normalized = path.trim_matches('/');
    normalized
        .rsplit_once('/')
        .is_some_and(|(parent, segment)| {
            !parent.contains('/') && segment.eq_ignore_ascii_case("freeze")
        })
}

pub(super) fn set(
    package: &mut NativeOfficePackage,
    requested: &str,
    pane: &NativeSpreadsheetFrozenPane,
) -> UseResult<String> {
    let sheet = resolve_sheet(package, requested)?;
    let pane = pane.normalized()?;
    let worksheet = package.xml_part(&sheet.part)?;
    let root = index_xml(&worksheet)?;
    let views = root
        .children
        .iter()
        .filter(|child| child.local_name == "sheetViews" && child.namespace == root.namespace)
        .collect::<Vec<_>>();
    if views.len() > 1 {
        return Err(view_part_error(
            &sheet.part,
            "contains multiple sheetViews collections",
        ));
    }
    let fragment = pane_fragment(prefix(&root.qualified_name), &pane);
    let edited = if let Some(views) = views.first().copied() {
        let matching = views
            .children
            .iter()
            .filter(|child| {
                child.local_name == "sheetView"
                    && child.namespace == root.namespace
                    && child.attributes.get("workbookViewId").map(String::as_str) == Some("0")
            })
            .collect::<Vec<_>>();
        if matching.len() > 1 {
            return Err(view_part_error(
                &sheet.part,
                "contains multiple sheetView elements for workbookViewId 0",
            ));
        }
        if let Some(view) = matching.first().copied() {
            let panes = direct_panes(view, root.namespace.as_deref());
            if panes.len() > 1 {
                return Err(view_part_error(
                    &sheet.part,
                    "contains multiple pane elements in workbookViewId 0",
                ));
            }
            if let Some(existing) = panes.first().copied() {
                require_mutable_pane(&worksheet, existing)?;
                apply_patches(
                    &worksheet,
                    vec![XmlPatch::new(existing.full_range.clone(), fragment)],
                )?
            } else {
                insert_ordered_child(
                    &worksheet,
                    view,
                    fragment,
                    &["selection", "pivotSelection", "extLst"],
                )?
            }
        } else {
            let tag = qualified(prefix(&views.qualified_name), "sheetView");
            insert_child(
                &worksheet,
                views,
                format!("<{tag} workbookViewId=\"0\">{fragment}</{tag}>"),
            )?
        }
    } else {
        let view_prefix = prefix(&root.qualified_name);
        let views_tag = qualified(view_prefix, "sheetViews");
        let view_tag = qualified(view_prefix, "sheetView");
        insert_ordered_child(
            &worksheet,
            &root,
            format!(
                "<{views_tag}><{view_tag} workbookViewId=\"0\">{fragment}</{view_tag}></{views_tag}>"
            ),
            WORKSHEET_CHILDREN_AFTER_VIEWS,
        )?
    };
    package.set_part(&sheet.part, edited)?;
    Ok(format!("{}/freeze", sheet.path))
}

pub(super) fn remove(package: &mut NativeOfficePackage, requested: &str) -> UseResult<()> {
    validate_mutation_path(requested)?;
    if !is_path(requested) {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Removing a frozen Spreadsheet pane requires a path such as /Sheet1/freeze.",
        ));
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let node = snapshot.get(requested, 0)?;
    if node.node_type != OfficeNodeType::FrozenPane {
        return Err(node_not_found(requested));
    }
    if node.format.get("nativeMutable").map(String::as_str) != Some("true") {
        return Err(editor_error(
            "use.office.spreadsheet_freeze_unknown_content",
            format!(
                "Frozen Spreadsheet pane '{}' contains unsupported view state or unknown content.",
                node.path
            ),
        ));
    }
    let (sheet_path, _) = node
        .path
        .rsplit_once('/')
        .ok_or_else(|| node_not_found(requested))?;
    let sheet = resolve_sheet(package, sheet_path)?;
    let worksheet = package.xml_part(&sheet.part)?;
    let root = index_xml(&worksheet)?;
    let panes = root
        .children
        .iter()
        .filter(|child| child.local_name == "sheetViews" && child.namespace == root.namespace)
        .flat_map(|views| views.children.iter())
        .filter(|view| {
            view.local_name == "sheetView"
                && view.namespace == root.namespace
                && view.attributes.get("workbookViewId").map(String::as_str) == Some("0")
        })
        .flat_map(|view| direct_panes(view, root.namespace.as_deref()))
        .collect::<Vec<_>>();
    if panes.len() != 1 {
        return Err(view_part_error(
            &sheet.part,
            "does not contain exactly one frozen pane for workbookViewId 0",
        ));
    }
    require_mutable_pane(&worksheet, panes[0])?;
    let edited = apply_patches(
        &worksheet,
        vec![XmlPatch::new(panes[0].full_range.clone(), Vec::new())],
    )?;
    package.set_part(&sheet.part, edited)
}

fn resolve_sheet(package: &NativeOfficePackage, requested: &str) -> UseResult<ResolvedSheet> {
    if package.kind() != DocumentKind::Spreadsheet {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Frozen pane operations are available only for Spreadsheet documents.",
        ));
    }
    validate_mutation_path(requested)?;
    if requested.trim_start_matches('/').contains('/') {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Setting a frozen Spreadsheet pane requires a worksheet path such as /Sheet1.",
        ));
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let sheet = snapshot
        .root()
        .children
        .iter()
        .find(|node| {
            node.node_type == OfficeNodeType::Worksheet && node.path.eq_ignore_ascii_case(requested)
        })
        .ok_or_else(|| node_not_found(requested))?;
    let part = sheet.format.get("part").cloned().ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_invalid",
            format!("Worksheet '{}' has no source part.", sheet.path),
        )
    })?;
    Ok(ResolvedSheet {
        path: sheet.path.clone(),
        part,
    })
}

fn pane_fragment(prefix: Option<&str>, pane: &NativeSpreadsheetFrozenPane) -> String {
    let tag = qualified(prefix, "pane");
    let x_split = if pane.frozen_columns > 0 {
        format!(" xSplit=\"{}\"", pane.frozen_columns)
    } else {
        String::new()
    };
    let y_split = if pane.frozen_rows > 0 {
        format!(" ySplit=\"{}\"", pane.frozen_rows)
    } else {
        String::new()
    };
    format!(
        "<{tag}{x_split}{y_split} topLeftCell=\"{}\" activePane=\"{}\" state=\"frozen\"/>",
        pane.top_left_cell,
        pane.active_pane()
    )
}

fn direct_panes<'a>(
    view: &'a crate::xml_edit::IndexedXmlElement,
    namespace: Option<&str>,
) -> Vec<&'a crate::xml_edit::IndexedXmlElement> {
    view.children
        .iter()
        .filter(|child| child.local_name == "pane" && child.namespace.as_deref() == namespace)
        .collect()
}

fn require_mutable_pane(
    part: &crate::LosslessXmlPart,
    pane: &crate::xml_edit::IndexedXmlElement,
) -> UseResult<()> {
    let known = ["xSplit", "ySplit", "topLeftCell", "activePane", "state"];
    let unknown_attribute = pane
        .qualified_attributes
        .keys()
        .any(|name| !known.contains(&name.as_str()));
    let content = std::str::from_utf8(&part.parse_bytes()[pane.content_range.clone()])
        .unwrap_or("<invalid>")
        .trim();
    if unknown_attribute || !pane.children.is_empty() || !content.is_empty() {
        return Err(editor_error(
            "use.office.spreadsheet_freeze_unknown_content",
            "Frozen Spreadsheet pane contains unknown attributes or child content.",
        ));
    }
    Ok(())
}

fn node_not_found(path: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.node_not_found",
        format!("Office semantic path '{path}' does not exist."),
    )
    .with_detail("path", path)
}

fn view_part_error(part: &str, reason: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_freeze_invalid",
        format!("Spreadsheet worksheet part '{part}' {reason}."),
    )
    .with_detail("part", part)
}
