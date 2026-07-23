use a3s_use_core::UseResult;

use super::{semantic_error, DocumentNode, OfficeNodeType};
use crate::editor::NativeSpreadsheetTableStyle;
use crate::spreadsheet_reference::CellRange;
use crate::xml_tree::XmlElement;
use crate::{NativeOfficePackage, OpcPackageModel, RelationshipSource, RelationshipTarget};

pub(super) fn read(
    package: &NativeOfficePackage,
    opc: &OpcPackageModel,
    worksheet: &XmlElement,
    worksheet_part: &str,
    sheet_path: &str,
) -> UseResult<Vec<DocumentNode>> {
    let collections = worksheet
        .child_elements()
        .filter(|child| child.local_name == "tableParts" && child.namespace == worksheet.namespace)
        .collect::<Vec<_>>();
    if collections.len() > 1 {
        return Err(table_error(
            worksheet_part,
            "contains multiple tableParts collections",
        ));
    }
    let Some(collection) = collections.first() else {
        return Ok(Vec::new());
    };
    let source = RelationshipSource::Part {
        part_name: worksheet_part.to_string(),
    };
    let mut output = Vec::new();
    for table_part in collection.children_named("tablePart") {
        if table_part.namespace != worksheet.namespace {
            continue;
        }
        let relationship_id = super::relationship_attribute(table_part, "id").ok_or_else(|| {
            table_error(
                worksheet_part,
                "contains a tablePart without a relationship ID",
            )
        })?;
        let relationship = opc
            .relationships()
            .relationship(&source, relationship_id)
            .ok_or_else(|| {
                table_error(
                    worksheet_part,
                    format!("references missing table relationship '{relationship_id}'"),
                )
            })?;
        if !relationship.relationship_type.ends_with("/table") {
            return Err(table_error(
                worksheet_part,
                format!("relationship '{relationship_id}' is not a Spreadsheet table relationship"),
            ));
        }
        let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
            return Err(table_error(
                worksheet_part,
                "uses an external Spreadsheet table relationship",
            ));
        };
        output.push(read_table(
            package,
            part_name,
            sheet_path,
            output.len() + 1,
        )?);
    }
    if let Some(count) = collection.attribute("count") {
        let count = count.parse::<usize>().map_err(|error| {
            table_error(
                worksheet_part,
                format!("contains invalid tableParts count '{count}': {error}"),
            )
        })?;
        if count != output.len() {
            return Err(table_error(
                worksheet_part,
                format!(
                    "declares {count} table parts but contains {} entries",
                    output.len()
                ),
            ));
        }
    }
    Ok(output)
}

fn read_table(
    package: &NativeOfficePackage,
    part_name: &str,
    sheet_path: &str,
    position: usize,
) -> UseResult<DocumentNode> {
    let part = package.xml_part(part_name)?;
    let table = crate::xml_tree::parse_xml_tree(&part)?;
    if table.local_name != "table"
        || !matches!(
            table.namespace.as_deref(),
            Some(super::SPREADSHEET_NAMESPACE | super::STRICT_SPREADSHEET_NAMESPACE)
        )
    {
        return Err(table_error(part_name, "has an unexpected root QName"));
    }
    let name = required_attribute(&table, "name", part_name)?;
    let display_name = table.attribute("displayName").unwrap_or(name);
    let reference = required_attribute(&table, "ref", part_name)?;
    let range = CellRange::parse(reference).map_err(|error| {
        table_error(
            part_name,
            format!("contains invalid range '{reference}': {error}"),
        )
    })?;
    let header_count = unsigned_attribute(&table, "headerRowCount", 1, part_name)?;
    let totals_count = unsigned_attribute(&table, "totalsRowCount", 0, part_name)?;
    let totals_shown = boolean_attribute(&table, "totalsRowShown", totals_count > 0, part_name)?;
    let header_row = header_count != 0;
    let totals_row = totals_count != 0 || totals_shown;

    let column_collections = table.children_named("tableColumns").collect::<Vec<_>>();
    if column_collections.len() != 1 {
        return Err(table_error(
            part_name,
            "must contain exactly one tableColumns collection",
        ));
    }
    let collection = column_collections[0];
    let columns = collection
        .children_named("tableColumn")
        .filter(|column| column.namespace == table.namespace)
        .collect::<Vec<_>>();
    let expected_width =
        usize::try_from(range.end.column - range.start.column + 1).unwrap_or(usize::MAX);
    let declared_count = unsigned_attribute(
        collection,
        "count",
        u64::try_from(columns.len()).unwrap_or(u64::MAX),
        part_name,
    )?;
    if columns.len() != expected_width
        || usize::try_from(declared_count).ok() != Some(columns.len())
    {
        return Err(table_error(
            part_name,
            format!(
                "has {} table columns for range '{}'",
                columns.len(),
                range.a1()
            ),
        ));
    }

    let totals_state_supported = matches!((totals_count, totals_shown), (0, false) | (1, true));
    let mut mutable = header_count <= 1
        && totals_count <= 1
        && totals_state_supported
        && root_supports_typed_mutation(&table);
    let mut node = DocumentNode::new(
        format!("{sheet_path}/table[{position}]"),
        "table",
        OfficeNodeType::Table,
    );
    node.text = name.to_string();
    node.format.insert("name".into(), name.into());
    node.format
        .insert("displayName".into(), display_name.into());
    node.format.insert("ref".into(), range.a1());
    node.format
        .insert("headerRow".into(), header_row.to_string());
    node.format
        .insert("totalsRow".into(), totals_row.to_string());
    let id = required_attribute(&table, "id", part_name)?;
    let id = id
        .parse::<u32>()
        .ok()
        .filter(|id| *id > 0)
        .ok_or_else(|| table_error(part_name, format!("contains invalid table ID '{id}'")))?;
    node.format.insert("id".into(), id.to_string());

    for (index, column) in columns.into_iter().enumerate() {
        let column_name = required_attribute(column, "name", part_name)?;
        let allowed_attributes = ["id", "name"];
        if column
            .attributes
            .iter()
            .any(|attribute| !allowed_attributes.contains(&attribute.local_name.as_str()))
            || column.child_elements().next().is_some()
        {
            mutable = false;
        }
        let mut child = DocumentNode::new(
            format!("{}/column[{}]", node.path, index + 1),
            "column",
            OfficeNodeType::TableColumn,
        );
        child.text = column_name.to_string();
        child.format.insert("name".into(), column_name.into());
        if let Some(id) = column.attribute("id") {
            child.format.insert("id".into(), id.into());
        }
        node.children.push(child);
    }

    let styles = table.children_named("tableStyleInfo").collect::<Vec<_>>();
    if styles.len() > 1 {
        return Err(table_error(
            part_name,
            "contains multiple tableStyleInfo elements",
        ));
    }
    let style = if let Some(style) = styles.first() {
        let parsed = NativeSpreadsheetTableStyle::from_ooxml_name(style.attribute("name"));
        if parsed.is_none() {
            mutable = false;
            if let Some(name) = style.attribute("name") {
                node.format.insert("styleName".into(), name.into());
            }
        }
        node.format.insert(
            "showFirstColumn".into(),
            boolean_attribute(style, "showFirstColumn", false, part_name)?.to_string(),
        );
        node.format.insert(
            "showLastColumn".into(),
            boolean_attribute(style, "showLastColumn", false, part_name)?.to_string(),
        );
        node.format.insert(
            "showRowStripes".into(),
            boolean_attribute(style, "showRowStripes", false, part_name)?.to_string(),
        );
        node.format.insert(
            "showColumnStripes".into(),
            boolean_attribute(style, "showColumnStripes", false, part_name)?.to_string(),
        );
        parsed
    } else {
        node.format.insert("showFirstColumn".into(), "false".into());
        node.format.insert("showLastColumn".into(), "false".into());
        node.format.insert("showRowStripes".into(), "false".into());
        node.format
            .insert("showColumnStripes".into(), "false".into());
        Some(NativeSpreadsheetTableStyle::None)
    };
    if let Some(style) = style {
        match style {
            NativeSpreadsheetTableStyle::None => {
                node.format.insert("styleFamily".into(), "none".into());
            }
            NativeSpreadsheetTableStyle::Light { number } => {
                node.format.insert("styleFamily".into(), "light".into());
                node.format.insert("styleNumber".into(), number.to_string());
            }
            NativeSpreadsheetTableStyle::Medium { number } => {
                node.format.insert("styleFamily".into(), "medium".into());
                node.format.insert("styleNumber".into(), number.to_string());
            }
            NativeSpreadsheetTableStyle::Dark { number } => {
                node.format.insert("styleFamily".into(), "dark".into());
                node.format.insert("styleNumber".into(), number.to_string());
            }
        }
    }
    let (filter, filter_mutable) =
        read_auto_filter(&table, part_name, &node.path, range, header_row, totals_row)?;
    mutable &= filter_mutable;
    if let Some(filter) = filter {
        node.children.push(filter);
    }
    node.format
        .insert("nativeMutable".into(), mutable.to_string());
    Ok(node)
}

fn root_supports_typed_mutation(table: &XmlElement) -> bool {
    const OWNED_ATTRIBUTES: &[&str] = &[
        "id",
        "name",
        "displayName",
        "ref",
        "headerRowCount",
        "totalsRowCount",
        "totalsRowShown",
        "tableType",
    ];
    let attributes_supported = table.attributes.iter().all(|attribute| {
        attribute.namespace.is_some() || OWNED_ATTRIBUTES.contains(&attribute.local_name.as_str())
    });
    let table_type_supported = table
        .attribute("tableType")
        .is_none_or(|value| value == "worksheet");
    let children_supported = table.child_elements().all(|child| {
        child.namespace != table.namespace
            || matches!(
                child.local_name.as_str(),
                "autoFilter" | "tableColumns" | "tableStyleInfo" | "extLst"
            )
    });
    attributes_supported && table_type_supported && children_supported
}

fn read_auto_filter(
    table: &XmlElement,
    part_name: &str,
    table_path: &str,
    range: CellRange,
    header_row: bool,
    totals_row: bool,
) -> UseResult<(Option<DocumentNode>, bool)> {
    let filters = table
        .children_named("autoFilter")
        .filter(|filter| filter.namespace == table.namespace)
        .collect::<Vec<_>>();
    if filters.len() > 1 {
        return Err(table_error(
            part_name,
            "contains multiple table autoFilter elements",
        ));
    }
    let Some(filter) = filters.first() else {
        return Ok((None, !header_row));
    };
    let node =
        super::auto_filter::read_element(filter, part_name, &format!("{table_path}/autofilter"))?;
    let observed_range = node
        .format
        .get("ref")
        .and_then(|value| CellRange::parse(value).ok());
    let Some(expected_range) = table_filter_range(range, totals_row) else {
        return Ok((Some(node), false));
    };
    let mutable = header_row
        && observed_range == Some(expected_range)
        && node.format.get("nativeMutable").map(String::as_str) == Some("true");
    Ok((Some(node), mutable))
}

fn table_filter_range(mut range: CellRange, totals_row: bool) -> Option<CellRange> {
    if totals_row {
        range.end.row = range.end.row.checked_sub(1)?;
        if range.end.row < range.start.row {
            return None;
        }
    }
    Some(range)
}

fn required_attribute<'a>(
    element: &'a XmlElement,
    name: &str,
    part_name: &str,
) -> UseResult<&'a str> {
    element.attribute(name).ok_or_else(|| {
        table_error(
            part_name,
            format!("contains an element without required '{name}'"),
        )
    })
}

fn unsigned_attribute(
    element: &XmlElement,
    name: &str,
    default: u64,
    part_name: &str,
) -> UseResult<u64> {
    element
        .attribute(name)
        .map(|value| {
            value.parse::<u64>().map_err(|error| {
                table_error(
                    part_name,
                    format!("contains invalid {name} '{value}': {error}"),
                )
            })
        })
        .transpose()
        .map(|value| value.unwrap_or(default))
}

fn boolean_attribute(
    element: &XmlElement,
    name: &str,
    default: bool,
    part_name: &str,
) -> UseResult<bool> {
    match element.attribute(name) {
        None => Ok(default),
        Some("1" | "true") => Ok(true),
        Some("0" | "false") => Ok(false),
        Some(value) => Err(table_error(
            part_name,
            format!("contains invalid boolean {name} '{value}'"),
        )),
    }
}

fn table_error(part_name: &str, reason: impl Into<String>) -> a3s_use_core::UseError {
    semantic_error(
        "use.office.spreadsheet_table_invalid",
        format!("Spreadsheet table part '{part_name}' {}.", reason.into()),
    )
    .with_detail("part", part_name)
}
