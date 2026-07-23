use a3s_use_core::UseResult;

use super::{node_not_found, require_spreadsheet, sort_error, validate_mutation_path};
use crate::semantic::{DocumentNode, NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::{CellRange, CellReference};
use crate::xml_edit::{apply_patches, index_xml, insert_ordered_child, XmlPatch};
use crate::{NativeOfficePackage, NativeSpreadsheetSort, NativeSpreadsheetSortDirection};

pub(super) fn is_path(path: &str) -> bool {
    let normalized = path.trim_matches('/');
    normalized.rsplit_once('/').is_some_and(|(sheet, leaf)| {
        !sheet.is_empty() && !sheet.contains('/') && leaf.eq_ignore_ascii_case("sort")
    })
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    require_spreadsheet(package)?;
    validate_mutation_path(path)?;
    if !is_path(path) {
        return Err(sort_error(
            "use.office.mutation_path_unsupported",
            "Spreadsheet sort-state removal requires a path such as /Sheet1/sort.",
        ));
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let node = snapshot.get(path, 0)?;
    if node.node_type != OfficeNodeType::SortState {
        return Err(node_not_found(path));
    }
    require_mutable(&node)?;
    let sheet_path = node
        .path
        .rsplit_once('/')
        .map(|(sheet, _)| sheet)
        .ok_or_else(|| node_not_found(path))?;
    let sheet = snapshot.get(sheet_path, 0)?;
    let part_name = sheet.format.get("part").ok_or_else(|| {
        sort_error(
            "use.office.spreadsheet_sheet_invalid",
            format!("Worksheet '{sheet_path}' has no source part."),
        )
    })?;
    let part = package.xml_part(part_name)?;
    let root = index_xml(&part)?;
    let states = direct_sort_states(&root);
    if states.len() != 1 {
        return Err(sort_error(
            "use.office.spreadsheet_sort_invalid",
            format!("Worksheet part '{part_name}' does not contain exactly one sortState."),
        ));
    }
    let edited = apply_patches(
        &part,
        vec![XmlPatch::new(states[0].full_range.clone(), Vec::new())],
    )?;
    package.set_part(part_name, edited)
}

pub(super) fn replace(
    package: &mut NativeOfficePackage,
    part_name: &str,
    range: CellRange,
    data_range: CellRange,
    keys: &[(u32, NativeSpreadsheetSortDirection)],
    request: &NativeSpreadsheetSort,
) -> UseResult<()> {
    let part = package.xml_part(part_name)?;
    let root = index_xml(&part)?;
    let states = direct_sort_states(&root);
    if states.len() > 1 {
        return Err(sort_error(
            "use.office.spreadsheet_sort_invalid",
            format!("Worksheet part '{part_name}' contains multiple sortState elements."),
        ));
    }
    if let Some(state) = states.first() {
        let edited = apply_patches(
            &part,
            vec![XmlPatch::new(state.full_range.clone(), Vec::new())],
        )?;
        package.set_part(part_name, edited)?;
    }
    let part = package.xml_part(part_name)?;
    let root = index_xml(&part)?;
    let prefix = root
        .qualified_name
        .rsplit_once(':')
        .map(|(prefix, _)| prefix);
    let tag = prefix.map_or_else(|| "sortState".to_string(), |p| format!("{p}:sortState"));
    let condition_tag = prefix.map_or_else(
        || "sortCondition".to_string(),
        |p| format!("{p}:sortCondition"),
    );
    let case_sensitive = if request.case_sensitive {
        " caseSensitive=\"1\""
    } else {
        ""
    };
    let mut fragment = format!(
        "<{tag} ref=\"{}\"{case_sensitive}>",
        quick_xml::escape::escape(range.a1())
    );
    for (column, direction) in keys {
        let key_ref = CellRange {
            start: CellReference {
                column: *column,
                row: data_range.start.row,
            },
            end: CellReference {
                column: *column,
                row: data_range.end.row,
            },
        }
        .a1();
        let descending = if direction.ooxml_descending() {
            " descending=\"1\""
        } else {
            ""
        };
        fragment.push_str(&format!(
            "<{condition_tag} ref=\"{}\"{descending}/>",
            quick_xml::escape::escape(&key_ref)
        ));
    }
    fragment.push_str(&format!("</{tag}>"));
    let edited = insert_ordered_child(
        &part,
        &root,
        fragment,
        &[
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
        ],
    )?;
    package.set_part(part_name, edited)
}

fn direct_sort_states(
    root: &crate::xml_edit::IndexedXmlElement,
) -> Vec<&crate::xml_edit::IndexedXmlElement> {
    root.children
        .iter()
        .filter(|child| child.local_name == "sortState" && child.namespace == root.namespace)
        .collect()
}

pub(super) fn require_mutable(node: &DocumentNode) -> UseResult<()> {
    if node.format.get("nativeMutable").map(String::as_str) == Some("true") {
        Ok(())
    } else {
        Err(sort_error(
            "use.office.spreadsheet_sort_unknown_content",
            format!(
                "Spreadsheet sort state '{}' contains unsupported sort methods, attributes, or unknown content.",
                node.path
            ),
        ))
    }
}
