use super::*;
use crate::editor::{escape_attribute, prefix, qualified};
use crate::xml_edit::{apply_patches, index_xml, insert_child, XmlPatch};
use crate::{Relationship, RelationshipSource, RelationshipTarget};

pub(super) fn add_chart(
    package: &mut NativeOfficePackage,
    parent: &str,
) -> UseResult<NativeCreatedPart> {
    let sheet = semantic_parent(package, parent, OfficeNodeType::Worksheet, "worksheet")?;
    let sheet_part = source_part(&sheet, "Spreadsheet worksheet")?;
    let dialect = dialect(package)?;
    let drawing_part = worksheet_drawing(package, &sheet_part, dialect)?;
    let position = relationship_count(package, &drawing_part, "chart")? + 1;
    create_owned_part(CreatePart {
        package,
        parent: &sheet.path,
        owner: &drawing_part,
        directory: "xl/charts",
        stem: "chart",
        content_type: CHART_CONTENT_TYPE,
        relationship_type: &dialect.relationship_type("chart"),
        xml: &chart_xml(dialect),
        path: &format!("{}/chart[{position}]", sheet.path),
        part_type: NativeOfficePartType::Chart,
    })
}

pub(in crate::editor) fn worksheet_drawing(
    package: &mut NativeOfficePackage,
    worksheet_part: &str,
    dialect: OfficeDialect,
) -> UseResult<String> {
    let worksheet = package.xml_part(worksheet_part)?;
    let index = index_xml(&worksheet)?;
    let drawing_nodes = index
        .children
        .iter()
        .filter(|child| child.local_name == "drawing")
        .collect::<Vec<_>>();
    if drawing_nodes.len() > 1 {
        return Err(part_error(
            "use.office.spreadsheet_drawing_invalid",
            "Spreadsheet worksheet contains more than one drawing element.",
        ));
    }
    let source = RelationshipSource::Part {
        part_name: worksheet_part.to_string(),
    };
    let relationships = package
        .opc_model()?
        .relationships()
        .relationships_from(&source)
        .iter()
        .filter(|relationship| relationship.relationship_type.ends_with("/drawing"))
        .cloned()
        .collect::<Vec<_>>();

    if let Some(drawing) = drawing_nodes.first() {
        let id = relationship_attribute(drawing).ok_or_else(|| {
            part_error(
                "use.office.spreadsheet_drawing_invalid",
                "Spreadsheet drawing element has no relationship ID.",
            )
        })?;
        let relationship = single_relationship(&relationships, id)?;
        let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
            return Err(part_error(
                "use.office.spreadsheet_drawing_invalid",
                "Spreadsheet drawing relationship must target an internal part.",
            ));
        };
        require_drawing_root(package, part_name, dialect)?;
        return Ok(part_name.clone());
    }
    if !relationships.is_empty() {
        return Err(part_error(
            "use.office.spreadsheet_drawing_invalid",
            "Spreadsheet worksheet has a drawing relationship without a drawing element.",
        ));
    }

    let drawing_part = allocate_part(package, "xl/drawings", "drawing")?;
    let xml = drawing_xml(dialect);
    crate::opc_edit::add_content_type_override(package, &drawing_part, DRAWING_CONTENT_TYPE)?;
    package.set_part(&drawing_part, xml.into_bytes())?;
    let relationship_id = crate::opc_edit::add_relationship(
        package,
        &relationship_part(worksheet_part),
        &dialect.relationship_type("drawing"),
        &relative_target(worksheet_part, &drawing_part),
    )?;
    insert_worksheet_drawing(package, worksheet_part, &relationship_id, dialect)?;
    Ok(drawing_part)
}

fn insert_worksheet_drawing(
    package: &mut NativeOfficePackage,
    worksheet_part: &str,
    relationship_id: &str,
    dialect: OfficeDialect,
) -> UseResult<()> {
    let worksheet = package.xml_part(worksheet_part)?;
    let index = index_xml(&worksheet)?;
    let tag = qualified(prefix(&index.qualified_name), "drawing");
    let fragment = format!(
        "<{tag} xmlns:r=\"{}\" r:id=\"{}\"/>",
        dialect.relationship_namespace(),
        escape_attribute(relationship_id)
    );
    let later = [
        "legacyDrawing",
        "legacyDrawingHF",
        "picture",
        "oleObjects",
        "controls",
        "webPublishItems",
        "tableParts",
        "extLst",
    ];
    let edited = if let Some(child) = index
        .children
        .iter()
        .find(|child| later.contains(&child.local_name.as_str()))
    {
        apply_patches(
            &worksheet,
            vec![XmlPatch::new(
                child.full_range.start..child.full_range.start,
                fragment,
            )],
        )?
    } else {
        insert_child(&worksheet, &index, fragment)?
    };
    package.set_part(worksheet_part, edited)
}

fn single_relationship<'a>(
    relationships: &'a [Relationship],
    id: &str,
) -> UseResult<&'a Relationship> {
    let mut matches = relationships
        .iter()
        .filter(|relationship| relationship.id == id);
    let relationship = matches.next().ok_or_else(|| {
        part_error(
            "use.office.spreadsheet_drawing_invalid",
            format!("Spreadsheet drawing relationship '{id}' does not exist."),
        )
    })?;
    if matches.next().is_some() {
        return Err(part_error(
            "use.office.spreadsheet_drawing_invalid",
            format!("Spreadsheet drawing relationship '{id}' is ambiguous."),
        ));
    }
    Ok(relationship)
}

fn relationship_attribute(element: &crate::xml_edit::IndexedXmlElement) -> Option<&str> {
    element
        .qualified_attributes
        .iter()
        .find(|(name, _)| name.ends_with(":id"))
        .map(|(_, value)| value.as_str())
}

fn require_drawing_root(
    package: &NativeOfficePackage,
    part_name: &str,
    dialect: OfficeDialect,
) -> UseResult<()> {
    let part = package.xml_part(part_name)?;
    if part.root().local_name == "wsDr"
        && part.root().namespace.as_deref() == Some(dialect.spreadsheet_drawing_namespace())
    {
        return Ok(());
    }
    Err(part_error(
        "use.office.spreadsheet_drawing_invalid",
        format!("Spreadsheet drawing part '/{part_name}' has an unexpected root QName."),
    ))
}
