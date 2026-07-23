use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::table_error;
use crate::editor::part::OfficeDialect;
use crate::editor::NativeSpreadsheetTable;
use crate::spreadsheet_reference::CellRange;
use crate::xml_edit::{
    apply_patches, escape_attribute, index_xml, insert_child, insert_ordered_child,
    patch_start_tag_attributes, IndexedXmlElement, XmlPatch,
};
use crate::LosslessXmlPart;

const TABLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

pub(super) fn content_type() -> &'static str {
    TABLE_CONTENT_TYPE
}

pub(super) fn new_table_xml(
    dialect: OfficeDialect,
    id: u32,
    table: &NativeSpreadsheetTable,
    range: CellRange,
) -> UseResult<Vec<u8>> {
    let namespace = dialect.spreadsheet_namespace();
    let display_name = table.display_name.as_deref().unwrap_or(&table.name);
    let mut children = String::new();
    if table.header_row {
        children.push_str(&super::super::filter_xml::fragment(
            None,
            &table_filter(table, range),
        )?);
    }
    children.push_str(&table_columns_fragment(None, table));
    if let Some(style_name) = table.style.ooxml_name() {
        children.push_str(&format!(
            "<tableStyleInfo name=\"{}\" showColumnStripes=\"{}\" showFirstColumn=\"{}\" showLastColumn=\"{}\" showRowStripes=\"{}\"/>",
            escape_attribute(&style_name),
            bool_token(table.show_column_stripes),
            bool_token(table.show_first_column),
            bool_token(table.show_last_column),
            bool_token(table.show_row_stripes),
        ));
    }
    Ok(format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><table displayName=\"{}\" headerRowCount=\"{}\" id=\"{id}\" name=\"{}\" ref=\"{}\" totalsRowCount=\"{}\" totalsRowShown=\"{}\" xmlns=\"{namespace}\">{children}</table>",
        escape_attribute(display_name),
        bool_token(table.header_row),
        escape_attribute(&table.name),
        escape_attribute(&range.a1()),
        bool_token(table.totals_row),
        bool_token(table.totals_row),
    )
    .into_bytes())
}

pub(super) fn replace_table(
    part: &LosslessXmlPart,
    table: &NativeSpreadsheetTable,
    range: CellRange,
) -> UseResult<Vec<u8>> {
    let root = index_xml(part)?;
    if root.local_name != "table" {
        return Err(table_error(part.name(), "has an unexpected root element"));
    }
    validate_table_columns(part, &root)?;
    let display_name = table.display_name.as_deref().unwrap_or(&table.name);
    let updates = BTreeMap::from([
        ("name".to_string(), Some(table.name.clone())),
        ("displayName".to_string(), Some(display_name.to_string())),
        ("ref".to_string(), Some(range.a1())),
        (
            "headerRowCount".to_string(),
            Some(bool_token(table.header_row).to_string()),
        ),
        (
            "totalsRowCount".to_string(),
            Some(bool_token(table.totals_row).to_string()),
        ),
        (
            "totalsRowShown".to_string(),
            Some(bool_token(table.totals_row).to_string()),
        ),
    ]);
    let bytes = patch_start_tag_attributes(part, &root, &updates)?;
    let mut edited = LosslessXmlPart::parse(part.name().to_string(), bytes)?;

    let root = index_xml(&edited)?;
    let columns = single_child(&root, "tableColumns", part.name())?
        .ok_or_else(|| table_error(part.name(), "has no tableColumns collection"))?;
    let replacement = table_columns_fragment(
        columns
            .qualified_name
            .rsplit_once(':')
            .map(|(prefix, _)| prefix),
        table,
    );
    let bytes = apply_patches(
        &edited,
        vec![XmlPatch::new(columns.full_range.clone(), replacement)],
    )?;
    edited = LosslessXmlPart::parse(part.name().to_string(), bytes)?;
    edited = patch_auto_filter(edited, table, range)?;
    edited = patch_style(edited, table)?;
    Ok(edited.raw().to_vec())
}

fn patch_auto_filter(
    part: LosslessXmlPart,
    table: &NativeSpreadsheetTable,
    range: CellRange,
) -> UseResult<LosslessXmlPart> {
    let root = index_xml(&part)?;
    let existing = single_child(&root, "autoFilter", part.name())?;
    let bytes = match (table.header_row, existing) {
        (true, Some(existing)) => {
            super::super::filter_xml::validate_mutable(&part, existing)?;
            apply_patches(
                &part,
                vec![XmlPatch::new(
                    existing.full_range.clone(),
                    super::super::filter_xml::fragment(
                        root.qualified_name
                            .rsplit_once(':')
                            .map(|(prefix, _)| prefix),
                        &table_filter(table, range),
                    )?,
                )],
            )?
        }
        (true, None) => {
            let prefix = root
                .qualified_name
                .rsplit_once(':')
                .map(|(prefix, _)| prefix);
            let columns = root
                .children
                .iter()
                .find(|child| child.local_name == "tableColumns")
                .ok_or_else(|| table_error(part.name(), "has no tableColumns collection"))?;
            apply_patches(
                &part,
                vec![XmlPatch::new(
                    columns.full_range.start..columns.full_range.start,
                    super::super::filter_xml::fragment(prefix, &table_filter(table, range))?,
                )],
            )?
        }
        (false, Some(existing)) => {
            super::super::filter_xml::validate_mutable(&part, existing)?;
            apply_patches(
                &part,
                vec![XmlPatch::new(existing.full_range.clone(), Vec::new())],
            )?
        }
        (false, None) => part.raw().to_vec(),
    };
    LosslessXmlPart::parse(part.name().to_string(), bytes)
}

fn patch_style(
    part: LosslessXmlPart,
    table: &NativeSpreadsheetTable,
) -> UseResult<LosslessXmlPart> {
    let root = index_xml(&part)?;
    let existing = single_child(&root, "tableStyleInfo", part.name())?;
    let style_name = table.style.ooxml_name();
    let bytes = match (style_name, existing) {
        (Some(style_name), Some(existing)) => patch_start_tag_attributes(
            &part,
            existing,
            &BTreeMap::from([
                ("name".to_string(), Some(style_name)),
                (
                    "showFirstColumn".to_string(),
                    Some(bool_token(table.show_first_column).to_string()),
                ),
                (
                    "showLastColumn".to_string(),
                    Some(bool_token(table.show_last_column).to_string()),
                ),
                (
                    "showRowStripes".to_string(),
                    Some(bool_token(table.show_row_stripes).to_string()),
                ),
                (
                    "showColumnStripes".to_string(),
                    Some(bool_token(table.show_column_stripes).to_string()),
                ),
            ]),
        )?,
        (Some(style_name), None) => {
            let prefix = root
                .qualified_name
                .rsplit_once(':')
                .map(|(prefix, _)| prefix);
            let name = qualified(prefix, "tableStyleInfo");
            let fragment = format!(
                "<{name} name=\"{}\" showColumnStripes=\"{}\" showFirstColumn=\"{}\" showLastColumn=\"{}\" showRowStripes=\"{}\"/>",
                escape_attribute(&style_name),
                bool_token(table.show_column_stripes),
                bool_token(table.show_first_column),
                bool_token(table.show_last_column),
                bool_token(table.show_row_stripes),
            );
            insert_ordered_child(&part, &root, fragment, &["extLst"])?
        }
        (None, Some(existing)) => {
            if !removable_element(
                &part,
                existing,
                &[
                    "name",
                    "showFirstColumn",
                    "showLastColumn",
                    "showRowStripes",
                    "showColumnStripes",
                ],
            ) {
                return Err(unknown_content(
                    part.name(),
                    "The table style cannot be removed without discarding unknown tableStyleInfo data.",
                ));
            }
            apply_patches(
                &part,
                vec![XmlPatch::new(existing.full_range.clone(), Vec::new())],
            )?
        }
        (None, None) => part.raw().to_vec(),
    };
    LosslessXmlPart::parse(part.name().to_string(), bytes)
}

pub(super) fn add_table_part_reference(
    worksheet: &LosslessXmlPart,
    relationship_namespace: &str,
    relationship_id: &str,
) -> UseResult<Vec<u8>> {
    let root = index_xml(worksheet)?;
    let collection = single_child(&root, "tableParts", worksheet.name())?;
    let prefix = root
        .qualified_name
        .rsplit_once(':')
        .map(|(prefix, _)| prefix);
    let item_name = qualified(prefix, "tablePart");
    let item = format!(
        "<{item_name} xmlns:r=\"{}\" r:id=\"{}\"/>",
        escape_attribute(relationship_namespace),
        escape_attribute(relationship_id)
    );
    if let Some(collection) = collection {
        validate_table_parts_children(worksheet, collection)?;
        let existing = collection
            .children
            .iter()
            .filter(|child| child.local_name == "tablePart")
            .count();
        let inserted = insert_child(worksheet, collection, item)?;
        let inserted = LosslessXmlPart::parse(worksheet.name().to_string(), inserted)?;
        let index = index_xml(&inserted)?;
        let collection = single_child(&index, "tableParts", worksheet.name())?
            .ok_or_else(|| table_error(worksheet.name(), "lost tableParts during insertion"))?;
        patch_start_tag_attributes(
            &inserted,
            collection,
            &BTreeMap::from([("count".to_string(), Some((existing + 1).to_string()))]),
        )
    } else {
        let collection_name = qualified(prefix, "tableParts");
        insert_ordered_child(
            worksheet,
            &root,
            format!("<{collection_name} count=\"1\">{item}</{collection_name}>"),
            &["extLst"],
        )
    }
}

pub(super) fn remove_table_part_reference(
    worksheet: &LosslessXmlPart,
    relationship_id: &str,
) -> UseResult<Vec<u8>> {
    let root = index_xml(worksheet)?;
    let collection = single_child(&root, "tableParts", worksheet.name())?
        .ok_or_else(|| table_error(worksheet.name(), "has no tableParts collection"))?;
    let matches = collection
        .children
        .iter()
        .filter(|child| {
            child.local_name == "tablePart"
                && relationship_attribute(child).map(String::as_str) == Some(relationship_id)
        })
        .collect::<Vec<_>>();
    if matches.len() != 1 {
        return Err(table_error(
            worksheet.name(),
            format!(
                "contains {} tablePart entries for relationship '{relationship_id}'",
                matches.len()
            ),
        ));
    }
    let target = matches[0];
    let entries = collection
        .children
        .iter()
        .filter(|child| child.local_name == "tablePart")
        .count();
    if entries == 1 {
        if collection
            .qualified_attributes
            .keys()
            .any(|name| name != "count")
            || collection.children.len() != 1
            || !whitespace_outside_children(worksheet, collection)
        {
            return Err(unknown_content(
                worksheet.name(),
                "The final table cannot be removed without discarding unknown tableParts collection data.",
            ));
        }
        apply_patches(
            worksheet,
            vec![XmlPatch::new(collection.full_range.clone(), Vec::new())],
        )
    } else {
        let bytes = apply_patches(
            worksheet,
            vec![XmlPatch::new(target.full_range.clone(), Vec::new())],
        )?;
        let edited = LosslessXmlPart::parse(worksheet.name().to_string(), bytes)?;
        let index = index_xml(&edited)?;
        let collection = single_child(&index, "tableParts", worksheet.name())?
            .ok_or_else(|| table_error(worksheet.name(), "lost tableParts during removal"))?;
        patch_start_tag_attributes(
            &edited,
            collection,
            &BTreeMap::from([("count".to_string(), Some((entries - 1).to_string()))]),
        )
    }
}

fn validate_table_columns(part: &LosslessXmlPart, root: &IndexedXmlElement) -> UseResult<()> {
    let collection = single_child(root, "tableColumns", part.name())?
        .ok_or_else(|| table_error(part.name(), "has no tableColumns collection"))?;
    if collection
        .qualified_attributes
        .keys()
        .any(|name| name != "count")
        || !whitespace_outside_children(part, collection)
    {
        return Err(unknown_content(
            part.name(),
            "The table columns cannot be replaced without discarding unknown tableColumns data.",
        ));
    }
    for child in &collection.children {
        if child.local_name != "tableColumn"
            || child.namespace != root.namespace
            || child
                .qualified_attributes
                .keys()
                .any(|name| !matches!(name.as_str(), "id" | "name"))
            || !child.children.is_empty()
            || !whitespace_outside_children(part, child)
        {
            return Err(unknown_content(
                part.name(),
                "The table columns contain formulas, totals metadata, extensions, or unknown content that the typed table contract does not own.",
            ));
        }
    }
    Ok(())
}

fn validate_table_parts_children(
    part: &LosslessXmlPart,
    collection: &IndexedXmlElement,
) -> UseResult<()> {
    if collection
        .children
        .iter()
        .any(|child| child.local_name != "tablePart" || child.namespace != collection.namespace)
        || !whitespace_outside_children(part, collection)
    {
        return Err(unknown_content(
            part.name(),
            "A table cannot be added without risking invalid tableParts child order.",
        ));
    }
    Ok(())
}

fn table_columns_fragment(prefix: Option<&str>, table: &NativeSpreadsheetTable) -> String {
    let collection = qualified(prefix, "tableColumns");
    let column = qualified(prefix, "tableColumn");
    let children = table
        .columns
        .iter()
        .enumerate()
        .map(|(index, value)| {
            format!(
                "<{column} id=\"{}\" name=\"{}\"/>",
                index + 1,
                escape_attribute(&value.name)
            )
        })
        .collect::<String>();
    format!(
        "<{collection} count=\"{}\">{children}</{collection}>",
        table.columns.len()
    )
}

fn table_filter(
    table: &NativeSpreadsheetTable,
    range: CellRange,
) -> crate::NativeSpreadsheetAutoFilter {
    let mut filter = range;
    if table.totals_row {
        filter.end.row -= 1;
    }
    crate::NativeSpreadsheetAutoFilter {
        range: filter.a1(),
        columns: table.filters.clone(),
    }
}

fn single_child<'a>(
    parent: &'a IndexedXmlElement,
    local_name: &str,
    part_name: &str,
) -> UseResult<Option<&'a IndexedXmlElement>> {
    let children = parent
        .children
        .iter()
        .filter(|child| child.local_name == local_name && child.namespace == parent.namespace)
        .collect::<Vec<_>>();
    if children.len() > 1 {
        Err(table_error(
            part_name,
            format!("contains multiple {local_name} elements"),
        ))
    } else {
        Ok(children.first().copied())
    }
}

fn removable_element(
    part: &LosslessXmlPart,
    element: &IndexedXmlElement,
    known_attributes: &[&str],
) -> bool {
    element
        .qualified_attributes
        .keys()
        .all(|name| known_attributes.contains(&name.as_str()))
        && element.children.is_empty()
        && whitespace_outside_children(part, element)
}

fn whitespace_outside_children(part: &LosslessXmlPart, element: &IndexedXmlElement) -> bool {
    let bytes = part.parse_bytes();
    let mut cursor = element.content_range.start;
    for child in &element.children {
        if bytes
            .get(cursor..child.full_range.start)
            .is_none_or(|slice| !slice.iter().all(u8::is_ascii_whitespace))
        {
            return false;
        }
        cursor = child.full_range.end;
    }
    bytes
        .get(cursor..element.content_range.end)
        .is_some_and(|slice| slice.iter().all(u8::is_ascii_whitespace))
}

fn relationship_attribute(element: &IndexedXmlElement) -> Option<&String> {
    element
        .qualified_attributes
        .iter()
        .find(|(name, _)| name.ends_with(":id"))
        .map(|(_, value)| value)
}

fn qualified(prefix: Option<&str>, local_name: &str) -> String {
    prefix.map_or_else(
        || local_name.to_string(),
        |prefix| format!("{prefix}:{local_name}"),
    )
}

const fn bool_token(value: bool) -> &'static str {
    if value {
        "1"
    } else {
        "0"
    }
}

fn unknown_content(part_name: &str, message: &str) -> a3s_use_core::UseError {
    super::super::editor_error(
        "use.office.spreadsheet_table_unknown_content",
        message,
    )
        .with_suggestion(
            "Inspect the package with native raw XML and preserve or relocate the unsupported content first.",
        )
        .with_detail("part", part_name)
}
