use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{semantic_error, DocumentNode, OfficeNodeType};
use crate::xml_tree::{parse_xml_tree, XmlElement, XmlNode};
use crate::{NativeOfficePackage, OpcPackageModel, RelationshipSource, RelationshipTarget};

const PRESENTATION_NAMESPACE: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PRESENTATION_NAMESPACE: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const DRAWING_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWING_NAMESPACE: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";

pub(super) fn read(
    package: &NativeOfficePackage,
    opc: &OpcPackageModel,
) -> UseResult<DocumentNode> {
    let part = package.xml_part("ppt/presentation.xml")?;
    let presentation = parse_xml_tree(&part)?;
    require_presentation_element(&presentation, "presentation", part.name())?;
    let mut root = DocumentNode::new("/", "presentation", OfficeNodeType::Presentation);
    if let Some(size) = presentation.child("sldSz") {
        copy_emu_attribute(size, "cx", "slideWidth", &mut root);
        copy_emu_attribute(size, "cy", "slideHeight", &mut root);
    }
    let source = RelationshipSource::Part {
        part_name: "ppt/presentation.xml".to_string(),
    };
    let slide_list = presentation.child("sldIdLst").ok_or_else(|| {
        semantic_error(
            "use.office.presentation_slides_missing",
            "Presentation has no slide ID list.",
        )
    })?;
    let mut slides = Vec::new();
    for (offset, slide_id) in slide_list.children_named("sldId").enumerate() {
        let index = offset + 1;
        let relationship_id = relationship_attribute(slide_id, "id").ok_or_else(|| {
            semantic_error(
                "use.office.presentation_slide_invalid",
                format!("Presentation slide {index} has no relationship ID."),
            )
        })?;
        let relationship = opc
            .relationships()
            .relationship(&source, relationship_id)
            .ok_or_else(|| {
                semantic_error(
                    "use.office.presentation_slide_missing",
                    format!(
                        "Presentation slide {index} references missing relationship '{relationship_id}'."
                    ),
                )
            })?;
        let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
            return Err(semantic_error(
                "use.office.presentation_slide_invalid",
                format!("Presentation slide {index} cannot use an external relationship."),
            ));
        };
        slides.push((index, slide_id, part_name.clone()));
    }
    let slide_paths = slides
        .iter()
        .map(|(index, _, part_name)| (part_name.clone(), format!("/slide[{index}]")))
        .collect::<BTreeMap<_, _>>();
    for (index, slide_id, part_name) in slides {
        let mut slide = read_slide(package, opc, &part_name, index, &slide_paths)?;
        if let Some(id) = slide_id.attribute("id") {
            slide.format.insert("slideId".into(), id.into());
        }
        root.children.push(slide);
    }
    root.text = root
        .children
        .iter()
        .map(|slide| slide.text.as_str())
        .collect::<Vec<_>>()
        .join("\n");
    Ok(root)
}

pub(super) fn virtual_get(
    root: &DocumentNode,
    path: &str,
    depth: usize,
) -> UseResult<Option<DocumentNode>> {
    if let Some(column) = virtual_table_column(root, path)? {
        return Ok(Some(column));
    }
    let Some((slide_path, requested)) = path.rsplit_once("/placeholder[") else {
        return Ok(None);
    };
    let Some(requested) = requested.strip_suffix(']') else {
        return Err(semantic_error(
            "use.office.path_invalid",
            format!("Presentation placeholder path '{path}' is missing ']'."),
        ));
    };
    let Some(slide) = root.children.iter().find(|node| node.path == slide_path) else {
        return Ok(None);
    };
    let mut placeholders = Vec::new();
    collect_placeholders(slide, &mut placeholders);
    let found = if let Ok(position) = requested.parse::<usize>() {
        if position == 0 {
            return Err(semantic_error(
                "use.office.path_invalid",
                "Presentation placeholder positions are one-based.",
            ));
        }
        placeholders.get(position - 1).copied()
    } else {
        let requested = normalize_placeholder_type(requested);
        placeholders.into_iter().find(|placeholder| {
            placeholder
                .format
                .get("placeholderType")
                .is_some_and(|value| normalize_placeholder_type(value) == requested)
        })
    };
    Ok(found.map(|node| node.clone_to_depth(depth)))
}

fn virtual_table_column(root: &DocumentNode, path: &str) -> UseResult<Option<DocumentNode>> {
    let Some((table_path, requested)) = path.rsplit_once("/col[") else {
        return Ok(None);
    };
    let Some(requested) = requested.strip_suffix(']') else {
        return Err(semantic_error(
            "use.office.path_invalid",
            format!("Presentation table-column path '{path}' is missing ']'."),
        ));
    };
    let position = requested
        .parse::<usize>()
        .ok()
        .filter(|value| *value > 0)
        .ok_or_else(|| {
            semantic_error(
                "use.office.path_invalid",
                "Presentation table-column positions are one-based positive integers.",
            )
        })?;
    let Some(table) =
        find_node(root, table_path).filter(|node| node.node_type == OfficeNodeType::Table)
    else {
        return Ok(None);
    };
    let widths = table
        .format
        .get("columnWidthsEmu")
        .map(|value| value.split(',').collect::<Vec<_>>())
        .unwrap_or_default();
    let Some(width) = widths.get(position - 1) else {
        return Ok(None);
    };
    let mut column = DocumentNode::new(path, "col", OfficeNodeType::TableColumn);
    column.format.insert("index".into(), position.to_string());
    column.format.insert("widthEmu".into(), (*width).into());
    Ok(Some(column))
}

fn find_node<'a>(node: &'a DocumentNode, path: &str) -> Option<&'a DocumentNode> {
    if node.path == path {
        return Some(node);
    }
    node.children
        .iter()
        .find_map(|child| find_node(child, path))
}

fn collect_placeholders<'a>(node: &'a DocumentNode, output: &mut Vec<&'a DocumentNode>) {
    for child in &node.children {
        if child.node_type == OfficeNodeType::Placeholder {
            output.push(child);
        }
        collect_placeholders(child, output);
    }
}

fn normalize_placeholder_type(value: &str) -> String {
    match value.to_ascii_lowercase().as_str() {
        "centertitle" => "ctrtitle".to_string(),
        "slidenum" => "sldnum".to_string(),
        value => value.to_string(),
    }
}

fn read_slide(
    package: &NativeOfficePackage,
    opc: &OpcPackageModel,
    part_name: &str,
    index: usize,
    slide_paths: &BTreeMap<String, String>,
) -> UseResult<DocumentNode> {
    let part = package.xml_part(part_name)?;
    let slide = parse_xml_tree(&part)?;
    require_presentation_element(&slide, "sld", part.name())?;
    let path = format!("/slide[{index}]");
    let mut node = DocumentNode::new(&path, "slide", OfficeNodeType::Slide);
    node.format.insert("part".into(), part_name.into());
    if let Some(show) = slide.attribute("show") {
        node.format.insert("visible".into(), show.into());
    }
    if let Some(shape_tree) = slide
        .child("cSld")
        .and_then(|common_slide| common_slide.child("spTree"))
    {
        read_shape_tree(
            shape_tree,
            &path,
            opc,
            part_name,
            slide_paths,
            &mut node.children,
        );
    }
    append_comments(package, opc, part_name, &path, &mut node)?;
    append_notes(package, opc, part_name, &path, &mut node)?;
    node.text = node
        .children
        .iter()
        .filter(|child| {
            !matches!(
                child.node_type,
                OfficeNodeType::Notes | OfficeNodeType::Comment
            )
        })
        .map(|child| child.text.as_str())
        .filter(|text| !text.is_empty())
        .collect::<Vec<_>>()
        .join("\n");
    Ok(node)
}

fn append_comments(
    package: &NativeOfficePackage,
    opc: &OpcPackageModel,
    slide_part: &str,
    slide_path: &str,
    slide: &mut DocumentNode,
) -> UseResult<()> {
    let source = RelationshipSource::Part {
        part_name: slide_part.to_string(),
    };
    let Some(relationship) = opc
        .relationships()
        .relationships_from(&source)
        .iter()
        .find(|relationship| relationship.relationship_type.ends_with("/comments"))
    else {
        return Ok(());
    };
    let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
        return Err(semantic_error(
            "use.office.comment_relationship_invalid",
            format!("Presentation slide '{slide_path}' comments relationship must be internal."),
        ));
    };
    let authors = read_comment_authors(package, opc)?;
    let part = package.xml_part(part_name)?;
    let comments = parse_xml_tree(&part)?;
    require_presentation_element(&comments, "cmLst", part.name())?;
    for (offset, comment) in comments.children_named("cm").enumerate() {
        let mut node = DocumentNode::new(
            format!("{slide_path}/comment[{}]", offset + 1),
            "comment",
            OfficeNodeType::Comment,
        );
        node.text = comment.child("text").map(direct_text).unwrap_or_default();
        node.format.insert("part".into(), part_name.clone());
        node.format
            .insert("ownerPart".into(), slide_part.to_string());
        for (attribute, key) in [("authorId", "authorId"), ("idx", "index"), ("dt", "date")] {
            if let Some(value) = comment.attribute(attribute) {
                node.format.insert(key.into(), value.into());
            }
        }
        if let Some(author_id) = comment
            .attribute("authorId")
            .and_then(|value| value.parse::<u32>().ok())
        {
            if let Some((name, initials)) = authors.get(&author_id) {
                node.format.insert("author".into(), name.clone());
                node.format.insert("initials".into(), initials.clone());
            }
        }
        if let Some(position) = comment.child("pos") {
            if let Some(x) = position.attribute("x") {
                node.format.insert("xEmu".into(), x.into());
            }
            if let Some(y) = position.attribute("y") {
                node.format.insert("yEmu".into(), y.into());
            }
        }
        slide.children.push(node);
    }
    Ok(())
}

fn read_comment_authors(
    package: &NativeOfficePackage,
    opc: &OpcPackageModel,
) -> UseResult<std::collections::BTreeMap<u32, (String, String)>> {
    let source = RelationshipSource::Part {
        part_name: "ppt/presentation.xml".to_string(),
    };
    let Some(relationship) = opc
        .relationships()
        .relationships_from(&source)
        .iter()
        .find(|relationship| relationship.relationship_type.ends_with("/commentAuthors"))
    else {
        return Ok(std::collections::BTreeMap::new());
    };
    let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
        return Err(semantic_error(
            "use.office.comment_relationship_invalid",
            "Presentation comment authors relationship must be internal.",
        ));
    };
    let part = package.xml_part(part_name)?;
    let root = parse_xml_tree(&part)?;
    require_presentation_element(&root, "cmAuthorLst", part.name())?;
    Ok(root
        .children_named("cmAuthor")
        .filter_map(|author| {
            author
                .attribute("id")
                .and_then(|id| id.parse::<u32>().ok())
                .map(|id| {
                    (
                        id,
                        (
                            author.attribute("name").unwrap_or_default().to_string(),
                            author.attribute("initials").unwrap_or_default().to_string(),
                        ),
                    )
                })
        })
        .collect())
}

fn append_notes(
    package: &NativeOfficePackage,
    opc: &OpcPackageModel,
    slide_part: &str,
    slide_path: &str,
    slide: &mut DocumentNode,
) -> UseResult<()> {
    let source = RelationshipSource::Part {
        part_name: slide_part.to_string(),
    };
    let Some(relationship) = opc
        .relationships()
        .relationships_from(&source)
        .iter()
        .find(|relationship| relationship.relationship_type.ends_with("/notesSlide"))
    else {
        return Ok(());
    };
    let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
        return Ok(());
    };
    if !package.contains_part(part_name) {
        return Ok(());
    }
    let part = package.xml_part(part_name)?;
    let notes = parse_xml_tree(&part)?;
    require_presentation_element(&notes, "notes", part.name())?;
    let mut node = DocumentNode::new(
        format!("{slide_path}/notes"),
        "notes",
        OfficeNodeType::Notes,
    );
    node.format.insert("part".into(), part_name.clone());
    let mut paragraphs = Vec::new();
    let mut descendants = Vec::new();
    notes.descendants(&mut descendants);
    for paragraph in descendants
        .into_iter()
        .filter(|element| element.local_name == "p")
    {
        let text = presentation_text(paragraph);
        if !text.is_empty() {
            paragraphs.push(text);
        }
    }
    node.text = paragraphs.join("\n");
    slide.children.push(node);
    Ok(())
}

fn read_shape_tree(
    tree: &XmlElement,
    parent_path: &str,
    opc: &OpcPackageModel,
    owner_part: &str,
    slide_paths: &BTreeMap<String, String>,
    output: &mut Vec<DocumentNode>,
) {
    let mut shape_index = 0_usize;
    let mut picture_index = 0_usize;
    let mut table_index = 0_usize;
    let mut chart_index = 0_usize;
    let mut connector_index = 0_usize;
    let mut group_index = 0_usize;
    for element in tree.child_elements() {
        match element.local_name.as_str() {
            "sp" => {
                shape_index += 1;
                output.push(read_shape(
                    element,
                    &format!("{parent_path}/shape[{shape_index}]"),
                    opc,
                    owner_part,
                    slide_paths,
                ));
            }
            "pic" => {
                picture_index += 1;
                output.push(read_picture(
                    element,
                    &format!("{parent_path}/picture[{picture_index}]"),
                ));
            }
            "graphicFrame" => {
                if find_descendant(element, "tbl").is_some() {
                    table_index += 1;
                    output.push(read_table(
                        element,
                        &format!("{parent_path}/table[{table_index}]"),
                    ));
                } else if find_descendant(element, "chart").is_some() {
                    chart_index += 1;
                    output.push(read_chart(
                        element,
                        &format!("{parent_path}/chart[{chart_index}]"),
                    ));
                } else {
                    shape_index += 1;
                    output.push(read_graphic_frame(
                        element,
                        &format!("{parent_path}/shape[{shape_index}]"),
                    ));
                }
            }
            "cxnSp" => {
                connector_index += 1;
                output.push(read_connector(
                    element,
                    &format!("{parent_path}/connector[{connector_index}]"),
                ));
            }
            "grpSp" => {
                group_index += 1;
                let path = format!("{parent_path}/group[{group_index}]");
                let mut group = DocumentNode::new(&path, "group", OfficeNodeType::Group);
                apply_non_visual_properties(element, &mut group);
                read_shape_tree(
                    element,
                    &path,
                    opc,
                    owner_part,
                    slide_paths,
                    &mut group.children,
                );
                group.text = group
                    .children
                    .iter()
                    .map(|child| child.text.as_str())
                    .filter(|text| !text.is_empty())
                    .collect::<Vec<_>>()
                    .join("\n");
                output.push(group);
            }
            _ => {}
        }
    }
}

fn read_shape(
    shape: &XmlElement,
    path: &str,
    opc: &OpcPackageModel,
    owner_part: &str,
    slide_paths: &BTreeMap<String, String>,
) -> DocumentNode {
    let placeholder = find_descendant(shape, "ph");
    let node_type = if placeholder.is_some() {
        OfficeNodeType::Placeholder
    } else {
        OfficeNodeType::Shape
    };
    let mut node = DocumentNode::new(path, "shape", node_type);
    apply_non_visual_properties(shape, &mut node);
    apply_transform(shape, &mut node);
    apply_fill(shape, &mut node);
    if let Some(placeholder) = placeholder {
        let placeholder_type = placeholder.attribute("type").unwrap_or("body");
        node.format
            .insert("placeholderType".into(), placeholder_type.into());
        if let Some(index) = placeholder.attribute("idx") {
            node.format.insert("placeholderIndex".into(), index.into());
        }
        if matches!(placeholder_type, "title" | "ctrTitle") {
            node.format.insert("title".into(), "true".into());
        }
    }
    if let Some(text_body) = shape.child("txBody") {
        read_text_body(text_body, path, &mut node.children);
        node.text = node
            .children
            .iter()
            .map(|paragraph| paragraph.text.as_str())
            .collect::<Vec<_>>()
            .join("\n");
    }
    append_shape_hyperlink(shape, path, opc, owner_part, slide_paths, &mut node);
    node
}

fn append_shape_hyperlink(
    shape: &XmlElement,
    path: &str,
    opc: &OpcPackageModel,
    owner_part: &str,
    slide_paths: &BTreeMap<String, String>,
    shape_node: &mut DocumentNode,
) {
    let Some(hyperlink) =
        find_descendant(shape, "cNvPr").and_then(|properties| properties.child("hlinkClick"))
    else {
        return;
    };
    let mut node = DocumentNode::new(
        format!("{path}/hyperlink"),
        "hyperlink",
        OfficeNodeType::Hyperlink,
    );
    if let Some(tooltip) = hyperlink.attribute("tooltip") {
        node.format.insert("tooltip".into(), tooltip.into());
    }
    if let Some(action) = hyperlink.attribute("action") {
        node.format.insert("action".into(), action.into());
    }
    let Some(id) = hyperlink.attribute("id") else {
        shape_node.children.push(node);
        return;
    };
    node.format.insert("relationshipId".into(), id.into());
    let source = RelationshipSource::Part {
        part_name: owner_part.to_string(),
    };
    if let Some(relationship) = opc.relationships().relationship(&source, id) {
        match (&*relationship.relationship_type, &relationship.target) {
            (relationship_type, RelationshipTarget::External { uri })
                if relationship_type.ends_with("/hyperlink") =>
            {
                node.format.insert("targetKind".into(), "external".into());
                node.format.insert("target".into(), uri.clone());
            }
            (
                relationship_type,
                RelationshipTarget::Internal {
                    part_name,
                    fragment,
                },
            ) if relationship_type.ends_with("/hyperlink") => {
                node.format.insert("targetKind".into(), "internal".into());
                let target = fragment.as_ref().map_or_else(
                    || part_name.clone(),
                    |fragment| format!("{part_name}#{fragment}"),
                );
                node.format.insert("target".into(), target);
            }
            (
                relationship_type,
                RelationshipTarget::Internal {
                    part_name,
                    fragment: None,
                },
            ) if relationship_type.ends_with("/slide") => {
                node.format.insert("targetKind".into(), "internal".into());
                node.format.insert(
                    "target".into(),
                    slide_paths
                        .get(part_name)
                        .cloned()
                        .unwrap_or_else(|| part_name.clone()),
                );
            }
            _ => {}
        }
    }
    shape_node.children.push(node);
}

fn read_picture(picture: &XmlElement, path: &str) -> DocumentNode {
    let mut node = DocumentNode::new(path, "picture", OfficeNodeType::Picture);
    apply_non_visual_properties(picture, &mut node);
    apply_transform(picture, &mut node);
    if let Some(blip) = find_descendant(picture, "blip") {
        if let Some(embed) = blip.attribute("embed") {
            node.format.insert("relationshipId".into(), embed.into());
        }
        if let Some(link) = blip.attribute("link") {
            node.format.insert("linkRelationshipId".into(), link.into());
        }
    }
    if let Some(transform) = find_descendant(picture, "xfrm") {
        if let Some(extent) = transform.child("ext") {
            copy_pixel_extent(extent, "cx", "widthPx", &mut node);
            copy_pixel_extent(extent, "cy", "heightPx", &mut node);
        }
    }
    node
}

fn copy_pixel_extent(element: &XmlElement, attribute: &str, key: &str, node: &mut DocumentNode) {
    if let Some(value) = element
        .attribute(attribute)
        .and_then(|value| value.parse::<u64>().ok())
    {
        node.format
            .insert(key.into(), ((value + 4_762) / 9_525).to_string());
    }
}

fn read_table(frame: &XmlElement, path: &str) -> DocumentNode {
    let mut node = DocumentNode::new(path, "table", OfficeNodeType::Table);
    apply_non_visual_properties(frame, &mut node);
    apply_transform(frame, &mut node);
    let Some(table) = find_descendant(frame, "tbl") else {
        return node;
    };
    if let Some(grid) = table.child("tblGrid") {
        let widths = grid
            .children_named("gridCol")
            .map(|column| column.attribute("w").unwrap_or("0"))
            .collect::<Vec<_>>();
        node.format
            .insert("columns".into(), widths.len().to_string());
        node.format
            .insert("columnWidthsEmu".into(), widths.join(","));
    }
    for (row_offset, row) in table.children_named("tr").enumerate() {
        let row_path = format!("{path}/tr[{}]", row_offset + 1);
        let mut row_node = DocumentNode::new(&row_path, "tr", OfficeNodeType::TableRow);
        for (cell_offset, cell) in row.children_named("tc").enumerate() {
            let cell_path = format!("{row_path}/tc[{}]", cell_offset + 1);
            let mut cell_node = DocumentNode::new(&cell_path, "tc", OfficeNodeType::TableCell);
            if let Some(text_body) = cell.child("txBody") {
                read_text_body(text_body, &cell_path, &mut cell_node.children);
                cell_node.text = cell_node
                    .children
                    .iter()
                    .map(|paragraph| paragraph.text.as_str())
                    .collect::<Vec<_>>()
                    .join("\n");
            }
            row_node.children.push(cell_node);
        }
        row_node.text = row_node
            .children
            .iter()
            .map(|cell| cell.text.as_str())
            .collect::<Vec<_>>()
            .join("\t");
        node.children.push(row_node);
    }
    node.text = node
        .children
        .iter()
        .map(|row| row.text.as_str())
        .collect::<Vec<_>>()
        .join("\n");
    node
}

fn read_chart(frame: &XmlElement, path: &str) -> DocumentNode {
    let mut node = DocumentNode::new(path, "chart", OfficeNodeType::Chart);
    apply_non_visual_properties(frame, &mut node);
    apply_transform(frame, &mut node);
    if let Some(chart) = find_descendant(frame, "chart") {
        if let Some(id) = chart.attribute("id") {
            node.format.insert("relationshipId".into(), id.into());
        }
    }
    node
}

fn read_graphic_frame(frame: &XmlElement, path: &str) -> DocumentNode {
    let mut node = DocumentNode::new(path, "shape", OfficeNodeType::Shape);
    apply_non_visual_properties(frame, &mut node);
    apply_transform(frame, &mut node);
    node
}

fn read_connector(connector: &XmlElement, path: &str) -> DocumentNode {
    let mut node = DocumentNode::new(path, "connector", OfficeNodeType::Connector);
    apply_non_visual_properties(connector, &mut node);
    apply_transform(connector, &mut node);
    node
}

fn read_text_body(body: &XmlElement, parent_path: &str, output: &mut Vec<DocumentNode>) {
    for (paragraph_offset, paragraph) in body.children_named("p").enumerate() {
        let paragraph_path = format!("{parent_path}/paragraph[{}]", paragraph_offset + 1);
        let mut paragraph_node =
            DocumentNode::new(&paragraph_path, "paragraph", OfficeNodeType::Paragraph);
        if let Some(properties) = paragraph.child("pPr") {
            if let Some(alignment) = properties.attribute("algn") {
                paragraph_node
                    .format
                    .insert("alignment".into(), alignment.into());
            }
            if let Some(level) = properties.attribute("lvl") {
                paragraph_node.format.insert("level".into(), level.into());
            }
        }
        let mut run_index = 0_usize;
        for child in paragraph.child_elements() {
            if !matches!(child.local_name.as_str(), "r" | "fld") {
                continue;
            }
            run_index += 1;
            let mut run_node = DocumentNode::new(
                format!("{paragraph_path}/run[{run_index}]"),
                "run",
                OfficeNodeType::Run,
            );
            run_node.text = presentation_text(child);
            if let Some(properties) = child.child("rPr") {
                apply_run_properties(properties, &mut run_node);
            }
            paragraph_node.children.push(run_node);
        }
        if let Some(end_properties) = paragraph.child("endParaRPr") {
            if paragraph_node.children.is_empty() {
                apply_run_properties(end_properties, &mut paragraph_node);
            }
        }
        paragraph_node.text = paragraph_node
            .children
            .iter()
            .map(|run| run.text.as_str())
            .collect();
        output.push(paragraph_node);
    }
}

fn apply_run_properties(properties: &XmlElement, node: &mut DocumentNode) {
    for (attribute, key) in [
        ("b", "bold"),
        ("i", "italic"),
        ("lang", "language"),
        ("kumimoji", "kumimoji"),
    ] {
        if let Some(value) = properties.attribute(attribute) {
            node.format.insert(key.into(), value.into());
        }
    }
    if let Some(underline) = properties.attribute("u") {
        let normalized = match underline {
            "sng" => "single",
            "dbl" => "double",
            value => value,
        };
        node.format.insert("underline".into(), normalized.into());
    }
    if let Some(value) = properties.attribute("cap") {
        let text_case = match value {
            "small" => "small-caps",
            "all" => "all-caps",
            _ => "none",
        };
        node.format.insert("textCase".into(), text_case.into());
    }
    if let Some(baseline) = properties
        .attribute("baseline")
        .and_then(|value| value.parse::<i32>().ok())
    {
        let script = match baseline.cmp(&0) {
            std::cmp::Ordering::Greater => "superscript",
            std::cmp::Ordering::Less => "subscript",
            std::cmp::Ordering::Equal => "baseline",
        };
        node.format.insert("script".into(), script.into());
    }
    if let Some(size) = properties
        .attribute("sz")
        .and_then(|value| value.parse::<f64>().ok())
    {
        node.format
            .insert("size".into(), format!("{}pt", size / 100.0));
    }
    if let Some(font) = properties
        .child("latin")
        .and_then(|font| font.attribute("typeface"))
    {
        node.format.insert("font".into(), font.into());
    }
    if let Some(color) = properties
        .child("solidFill")
        .and_then(|fill| fill.child("srgbClr"))
        .and_then(|color| color.attribute("val"))
    {
        node.format.insert("color".into(), color.into());
    }
    if let Some(highlight) = properties.child("highlight") {
        if let Some(rgb) = highlight
            .child("srgbClr")
            .and_then(|color| color.attribute("val"))
        {
            node.format
                .insert("highlight".into(), highlight_name(rgb).into());
        } else if let Some(color) = highlight.child_elements().next() {
            if let Some(value) = color.attribute("val") {
                node.format.insert("highlight".into(), value.into());
            }
        }
    }
}

fn highlight_name(rgb: &str) -> &str {
    match rgb.to_ascii_uppercase().as_str() {
        "000000" => "black",
        "0000FF" => "blue",
        "00FFFF" => "cyan",
        "000080" => "dark-blue",
        "008080" => "dark-cyan",
        "808080" => "dark-gray",
        "008000" => "dark-green",
        "800080" => "dark-magenta",
        "800000" => "dark-red",
        "808000" => "dark-yellow",
        "00FF00" => "green",
        "C0C0C0" => "light-gray",
        "FF00FF" => "magenta",
        "FF0000" => "red",
        "FFFFFF" => "white",
        "FFFF00" => "yellow",
        _ => rgb,
    }
}

fn apply_non_visual_properties(element: &XmlElement, node: &mut DocumentNode) {
    let Some(properties) = find_descendant(element, "cNvPr") else {
        return;
    };
    for (attribute, key) in [
        ("id", "id"),
        ("name", "name"),
        ("title", "titleText"),
        ("hidden", "hidden"),
    ] {
        if let Some(value) = properties.attribute(attribute) {
            node.format.insert(key.into(), value.into());
        }
    }
    if let Some(description) = properties.attribute("descr") {
        node.format.insert("alt".into(), description.into());
    } else if let Some(title) = properties.attribute("title") {
        node.format.insert("alt".into(), title.into());
    }
}

fn apply_transform(element: &XmlElement, node: &mut DocumentNode) {
    let Some(transform) = find_descendant(element, "xfrm") else {
        return;
    };
    if let Some(offset) = transform.child("off") {
        copy_emu_attribute(offset, "x", "x", node);
        copy_emu_attribute(offset, "y", "y", node);
    }
    if let Some(extent) = transform.child("ext") {
        copy_emu_attribute(extent, "cx", "width", node);
        copy_emu_attribute(extent, "cy", "height", node);
    }
    if let Some(rotation) = transform
        .attribute("rot")
        .and_then(|value| value.parse::<f64>().ok())
    {
        node.format
            .insert("rotation".into(), format!("{}deg", rotation / 60_000.0));
    }
}

fn apply_fill(element: &XmlElement, node: &mut DocumentNode) {
    if let Some(color) = element
        .child("spPr")
        .and_then(|properties| properties.child("solidFill"))
        .and_then(|fill| fill.child("srgbClr"))
        .and_then(|color| color.attribute("val"))
    {
        node.format.insert("fill".into(), color.into());
    }
}

fn copy_emu_attribute(element: &XmlElement, attribute: &str, key: &str, node: &mut DocumentNode) {
    if let Some(value) = element
        .attribute(attribute)
        .and_then(|value| value.parse::<f64>().ok())
    {
        node.format
            .insert(key.into(), format!("{:.4}cm", value / 360_000.0));
    }
}

fn presentation_text(element: &XmlElement) -> String {
    let mut output = String::new();
    append_presentation_text(element, &mut output);
    output
}

fn append_presentation_text(element: &XmlElement, output: &mut String) {
    let is_drawing = matches!(
        element.namespace.as_deref(),
        Some(DRAWING_NAMESPACE | STRICT_DRAWING_NAMESPACE)
    );
    if is_drawing && element.local_name == "t" {
        output.push_str(&direct_text(element));
        return;
    }
    if is_drawing && element.local_name == "br" {
        output.push('\n');
        return;
    }
    for child in element.child_elements() {
        append_presentation_text(child, output);
    }
}

fn direct_text(element: &XmlElement) -> String {
    element
        .children
        .iter()
        .filter_map(|child| match child {
            XmlNode::Text(text) => Some(text.as_str()),
            XmlNode::Element(_) => None,
        })
        .collect()
}

fn find_descendant<'a>(element: &'a XmlElement, local_name: &str) -> Option<&'a XmlElement> {
    for child in element.child_elements() {
        if child.local_name == local_name {
            return Some(child);
        }
        if let Some(found) = find_descendant(child, local_name) {
            return Some(found);
        }
    }
    None
}

fn require_presentation_element(
    element: &XmlElement,
    local_name: &str,
    part: &str,
) -> UseResult<()> {
    if element.local_name == local_name
        && matches!(
            element.namespace.as_deref(),
            Some(PRESENTATION_NAMESPACE | STRICT_PRESENTATION_NAMESPACE)
        )
    {
        return Ok(());
    }
    Err(semantic_error(
        "use.office.presentation_xml_invalid",
        format!("Presentation part '{part}' has an unexpected root element."),
    ))
}

fn relationship_attribute<'a>(element: &'a XmlElement, local_name: &str) -> Option<&'a str> {
    element
        .attribute_ns(RELATIONSHIPS_NAMESPACE, local_name)
        .or_else(|| element.attribute_ns(STRICT_RELATIONSHIPS_NAMESPACE, local_name))
}
