use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{semantic_error, DocumentNode, OfficeNodeType};
use crate::xml_tree::{parse_xml_tree, XmlElement, XmlNode};
use crate::{NativeOfficePackage, OpcPackageModel, RelationshipSource, RelationshipTarget};

const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD_NAMESPACE: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

pub(super) fn read(
    package: &NativeOfficePackage,
    opc: &OpcPackageModel,
) -> UseResult<DocumentNode> {
    let part = package.xml_part("word/document.xml")?;
    let document = parse_xml_tree(&part)?;
    require_word_element(&document, "document", part.name())?;
    let body = document.child("body").ok_or_else(|| {
        semantic_error(
            "use.office.word_body_missing",
            "Word document has no w:body element.",
        )
    })?;
    let styles = read_styles(package)?;
    let mut comment_anchors = BTreeMap::new();

    let mut root = DocumentNode::new("/", "document", OfficeNodeType::Document);
    let mut body_node = DocumentNode::new("/body", "body", OfficeNodeType::Body);
    body_node
        .format
        .insert("part".into(), "word/document.xml".into());
    read_block_children(
        body,
        "/body",
        &styles,
        opc,
        "word/document.xml",
        &mut comment_anchors,
        &mut body_node.children,
    );
    body_node.text = join_block_text(&body_node.children);
    root.text = body_node.text.clone();
    root.children.push(body_node);
    append_comments(package, opc, &comment_anchors, &mut root)?;
    read_headers_and_footers(package, opc, &styles, &mut root)?;
    Ok(root)
}

fn read_headers_and_footers(
    package: &NativeOfficePackage,
    opc: &OpcPackageModel,
    styles: &BTreeMap<String, String>,
    root: &mut DocumentNode,
) -> UseResult<()> {
    let source = RelationshipSource::Part {
        part_name: "word/document.xml".to_string(),
    };
    let mut header_index = 0_usize;
    let mut footer_index = 0_usize;
    for relationship in opc.relationships().relationships_from(&source) {
        let (tag, node_type, index) = if relationship.relationship_type.ends_with("/header") {
            header_index += 1;
            ("header", OfficeNodeType::Header, header_index)
        } else if relationship.relationship_type.ends_with("/footer") {
            footer_index += 1;
            ("footer", OfficeNodeType::Footer, footer_index)
        } else {
            continue;
        };
        let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
            continue;
        };
        if !package.contains_part(part_name) {
            continue;
        }
        let part = package.xml_part(part_name)?;
        let container = parse_xml_tree(&part)?;
        let path = format!("/{tag}[{index}]");
        let mut node = DocumentNode::new(&path, tag, node_type);
        node.format.insert("part".into(), part_name.clone());
        let mut ignored_comment_anchors = BTreeMap::new();
        read_block_children(
            &container,
            &path,
            styles,
            opc,
            part_name,
            &mut ignored_comment_anchors,
            &mut node.children,
        );
        node.text = join_block_text(&node.children);
        root.children.push(node);
    }
    Ok(())
}

fn append_comments(
    package: &NativeOfficePackage,
    opc: &OpcPackageModel,
    anchors: &BTreeMap<String, String>,
    root: &mut DocumentNode,
) -> UseResult<()> {
    let source = RelationshipSource::Part {
        part_name: "word/document.xml".to_string(),
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
            "Word comments relationship must be internal.",
        ));
    };
    let part = package.xml_part(part_name)?;
    let comments = parse_xml_tree(&part)?;
    require_word_element(&comments, "comments", part.name())?;
    for (offset, comment) in comments.children_named("comment").enumerate() {
        let id = comment.attribute("id").ok_or_else(|| {
            semantic_error(
                "use.office.comment_id_missing",
                format!("Word comment {} has no ID.", offset + 1),
            )
        })?;
        let mut node = DocumentNode::new(
            format!("/comments/comment[{}]", offset + 1),
            "comment",
            OfficeNodeType::Comment,
        );
        node.text = word_text(comment);
        node.format.insert("id".into(), id.into());
        node.format.insert("part".into(), part_name.clone());
        for (attribute, key) in [
            ("author", "author"),
            ("initials", "initials"),
            ("date", "date"),
        ] {
            if let Some(value) = comment.attribute(attribute) {
                node.format.insert(key.into(), value.into());
            }
        }
        if let Some(anchor) = anchors.get(id) {
            node.format.insert("anchoredTo".into(), anchor.clone());
        }
        root.children.push(node);
    }
    Ok(())
}

fn collect_paragraph_comment_anchors(
    paragraph: &XmlElement,
    path: &str,
    anchors: &mut BTreeMap<String, String>,
) {
    let mut descendants = Vec::new();
    paragraph.descendants(&mut descendants);
    for marker_name in ["commentRangeStart", "commentReference"] {
        for marker in descendants
            .iter()
            .copied()
            .filter(|element| element.local_name == marker_name)
        {
            if let Some(id) = marker.attribute("id") {
                anchors
                    .entry(id.to_string())
                    .or_insert_with(|| path.to_string());
            }
        }
    }
}

fn read_styles(package: &NativeOfficePackage) -> UseResult<BTreeMap<String, String>> {
    if !package.contains_part("word/styles.xml") {
        return Ok(BTreeMap::new());
    }
    let part = package.xml_part("word/styles.xml")?;
    let root = parse_xml_tree(&part)?;
    require_word_element(&root, "styles", part.name())?;
    let mut styles = BTreeMap::new();
    for style in root.children_named("style") {
        let Some(style_id) = style.attribute("styleId") else {
            continue;
        };
        let name = style
            .child("name")
            .and_then(|name| name.attribute("val"))
            .unwrap_or(style_id);
        styles.insert(style_id.to_string(), name.to_string());
    }
    Ok(styles)
}

fn read_block_children(
    container: &XmlElement,
    parent_path: &str,
    styles: &BTreeMap<String, String>,
    opc: &OpcPackageModel,
    owner_part: &str,
    comment_anchors: &mut BTreeMap<String, String>,
    output: &mut Vec<DocumentNode>,
) {
    let mut paragraph_index = output
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Paragraph)
        .count();
    let mut table_index = output
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Table)
        .count();
    for element in container.child_elements() {
        match element.local_name.as_str() {
            "p" => {
                paragraph_index += 1;
                output.push(read_paragraph(
                    element,
                    &format!("{parent_path}/p[{paragraph_index}]"),
                    styles,
                    opc,
                    owner_part,
                    comment_anchors,
                ));
            }
            "tbl" => {
                table_index += 1;
                output.push(read_table(
                    element,
                    &format!("{parent_path}/tbl[{table_index}]"),
                    styles,
                    opc,
                    owner_part,
                    comment_anchors,
                ));
            }
            "sdt" | "customXml" | "ins" | "del" | "moveFrom" | "moveTo" => {
                let nested = element.child("sdtContent").unwrap_or(element);
                read_block_children(
                    nested,
                    parent_path,
                    styles,
                    opc,
                    owner_part,
                    comment_anchors,
                    output,
                );
                paragraph_index = output
                    .iter()
                    .filter(|node| node.node_type == OfficeNodeType::Paragraph)
                    .count();
                table_index = output
                    .iter()
                    .filter(|node| node.node_type == OfficeNodeType::Table)
                    .count();
            }
            _ => {}
        }
    }
}

fn read_paragraph(
    paragraph: &XmlElement,
    path: &str,
    styles: &BTreeMap<String, String>,
    opc: &OpcPackageModel,
    owner_part: &str,
    comment_anchors: &mut BTreeMap<String, String>,
) -> DocumentNode {
    let mut node = DocumentNode::new(path, "p", OfficeNodeType::Paragraph);
    collect_paragraph_comment_anchors(paragraph, path, comment_anchors);
    if let Some(properties) = paragraph.child("pPr") {
        apply_paragraph_properties(properties, styles, &mut node);
    }
    let mut run_index = 0_usize;
    let mut hyperlink_index = 0_usize;
    for child in paragraph.child_elements() {
        match child.local_name.as_str() {
            "r" => {
                if is_comment_reference_run(child) {
                    continue;
                }
                run_index += 1;
                node.children
                    .push(read_run(child, &format!("{path}/r[{run_index}]")));
            }
            "hyperlink" => {
                hyperlink_index += 1;
                let hyperlink_path = format!("{path}/hyperlink[{hyperlink_index}]");
                let mut hyperlink =
                    DocumentNode::new(&hyperlink_path, "hyperlink", OfficeNodeType::Hyperlink);
                apply_hyperlink_target(child, opc, owner_part, &mut hyperlink);
                for run in child.children_named("r") {
                    run_index += 1;
                    hyperlink
                        .children
                        .push(read_run(run, &format!("{hyperlink_path}/r[{run_index}]")));
                }
                hyperlink.text = hyperlink
                    .children
                    .iter()
                    .map(|child| child.text.as_str())
                    .collect();
                node.children.push(hyperlink);
            }
            "fldSimple" | "sdt" | "smartTag" | "ins" | "del" => {
                append_nested_runs(child, path, &mut run_index, &mut node.children);
            }
            _ => {}
        }
    }
    node.text = node
        .children
        .iter()
        .map(|child| child.text.as_str())
        .collect();
    node
}

fn apply_hyperlink_target(
    element: &XmlElement,
    opc: &OpcPackageModel,
    owner_part: &str,
    node: &mut DocumentNode,
) {
    for (attribute, key) in [
        ("tooltip", "tooltip"),
        ("history", "history"),
        ("docLocation", "docLocation"),
        ("tgtFrame", "targetFrame"),
    ] {
        if let Some(value) = element.attribute(attribute) {
            node.format.insert(key.into(), value.into());
        }
    }
    if let Some(anchor) = element.attribute("anchor") {
        node.format.insert("targetKind".into(), "internal".into());
        node.format.insert("target".into(), anchor.into());
        return;
    }
    let Some(id) = element.attribute("id") else {
        return;
    };
    node.format.insert("relationshipId".into(), id.into());
    let source = RelationshipSource::Part {
        part_name: owner_part.to_string(),
    };
    let Some(relationship) = opc
        .relationships()
        .relationship(&source, id)
        .filter(|relationship| relationship.relationship_type.ends_with("/hyperlink"))
    else {
        return;
    };
    match &relationship.target {
        RelationshipTarget::External { uri } => {
            node.format.insert("targetKind".into(), "external".into());
            node.format.insert("target".into(), uri.clone());
        }
        RelationshipTarget::Internal {
            part_name,
            fragment,
        } => {
            node.format.insert("targetKind".into(), "internal".into());
            let target = fragment.as_ref().map_or_else(
                || part_name.clone(),
                |fragment| format!("{part_name}#{fragment}"),
            );
            node.format.insert("target".into(), target);
        }
    }
}

fn append_nested_runs(
    element: &XmlElement,
    paragraph_path: &str,
    run_index: &mut usize,
    output: &mut Vec<DocumentNode>,
) {
    for child in element.child_elements() {
        if child.local_name == "r" {
            if is_comment_reference_run(child) {
                continue;
            }
            *run_index += 1;
            output.push(read_run(child, &format!("{paragraph_path}/r[{run_index}]")));
        } else {
            append_nested_runs(child, paragraph_path, run_index, output);
        }
    }
}

fn is_comment_reference_run(run: &XmlElement) -> bool {
    let meaningful = run
        .child_elements()
        .filter(|child| child.local_name != "rPr")
        .collect::<Vec<_>>();
    meaningful.len() == 1 && meaningful[0].local_name == "commentReference"
}

fn read_run(run: &XmlElement, path: &str) -> DocumentNode {
    let mut node = DocumentNode::new(path, "r", OfficeNodeType::Run);
    if let Some(properties) = run.child("rPr") {
        apply_run_properties(properties, &mut node);
    }
    node.text = word_text(run);
    if let Some(blip) = find_descendant(run, "blip") {
        node.tag = "picture".into();
        node.node_type = OfficeNodeType::Picture;
        if let Some(embed) = blip.attribute("embed") {
            node.format.insert("relationshipId".into(), embed.into());
        }
        if let Some(properties) = find_descendant(run, "docPr") {
            copy_picture_attribute(properties, "name", "name", &mut node);
            if let Some(alt) = properties
                .attribute("descr")
                .or_else(|| properties.attribute("title"))
            {
                node.format.insert("alt".into(), alt.into());
            }
        }
        if let Some(extent) = find_descendant(run, "extent") {
            copy_picture_extent(extent, "cx", "widthPx", &mut node);
            copy_picture_extent(extent, "cy", "heightPx", &mut node);
        }
    }
    node
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

fn copy_picture_attribute(
    element: &XmlElement,
    attribute: &str,
    key: &str,
    node: &mut DocumentNode,
) {
    if let Some(value) = element.attribute(attribute) {
        node.format.insert(key.into(), value.into());
    }
}

fn copy_picture_extent(element: &XmlElement, attribute: &str, key: &str, node: &mut DocumentNode) {
    if let Some(value) = element
        .attribute(attribute)
        .and_then(|value| value.parse::<u64>().ok())
    {
        node.format
            .insert(key.into(), ((value + 4_762) / 9_525).to_string());
    }
}

fn read_table(
    table: &XmlElement,
    path: &str,
    styles: &BTreeMap<String, String>,
    opc: &OpcPackageModel,
    owner_part: &str,
    comment_anchors: &mut BTreeMap<String, String>,
) -> DocumentNode {
    let mut node = DocumentNode::new(path, "tbl", OfficeNodeType::Table);
    if let Some(style) = table
        .child("tblPr")
        .and_then(|properties| properties.child("tblStyle"))
        .and_then(|style| style.attribute("val"))
    {
        node.style = Some(style.to_string());
        if let Some(name) = styles.get(style) {
            node.format.insert("styleName".into(), name.clone());
        }
    }
    for (row_offset, row) in table.children_named("tr").enumerate() {
        let row_path = format!("{path}/tr[{}]", row_offset + 1);
        let mut row_node = DocumentNode::new(&row_path, "tr", OfficeNodeType::TableRow);
        for (cell_offset, cell) in row.children_named("tc").enumerate() {
            let cell_path = format!("{row_path}/tc[{}]", cell_offset + 1);
            let mut cell_node = DocumentNode::new(&cell_path, "tc", OfficeNodeType::TableCell);
            read_block_children(
                cell,
                &cell_path,
                styles,
                opc,
                owner_part,
                comment_anchors,
                &mut cell_node.children,
            );
            cell_node.text = join_block_text(&cell_node.children);
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

fn apply_paragraph_properties(
    properties: &XmlElement,
    styles: &BTreeMap<String, String>,
    node: &mut DocumentNode,
) {
    if let Some(style_id) = properties
        .child("pStyle")
        .and_then(|style| style.attribute("val"))
    {
        node.style = Some(style_id.to_string());
        if let Some(name) = styles.get(style_id) {
            node.format.insert("styleName".into(), name.clone());
        }
    }
    copy_child_value(properties, "jc", "alignment", node);
    if let Some(numbering) = properties.child("numPr") {
        copy_child_value(numbering, "numId", "numId", node);
        copy_child_value(numbering, "ilvl", "numLevel", node);
    }
    if let Some(direction) = properties.child("bidi") {
        let direction = if bool_value(direction) { "rtl" } else { "ltr" };
        node.format.insert("direction".into(), direction.into());
    }
}

fn apply_run_properties(properties: &XmlElement, node: &mut DocumentNode) {
    for (child, key) in [
        ("b", "bold"),
        ("i", "italic"),
        ("strike", "strike"),
        ("dstrike", "doubleStrike"),
        ("vanish", "hidden"),
        ("rtl", "rtl"),
    ] {
        if let Some(property) = properties.child(child) {
            node.format
                .insert(key.into(), bool_value(property).to_string());
        }
    }
    let caps = properties.child("caps").map(bool_value);
    let small_caps = properties.child("smallCaps").map(bool_value);
    if caps.is_some() || small_caps.is_some() {
        let text_case = if caps == Some(true) {
            "all-caps"
        } else if small_caps == Some(true) {
            "small-caps"
        } else {
            "none"
        };
        node.format.insert("textCase".into(), text_case.into());
    }
    if let Some(fonts) = properties.child("rFonts") {
        for (attribute, key) in [
            ("ascii", "font"),
            ("eastAsia", "font.eastAsia"),
            ("cs", "font.complexScript"),
        ] {
            if let Some(value) = fonts.attribute(attribute) {
                node.format.insert(key.into(), value.into());
            }
        }
    }
    if let Some(size) = properties
        .child("sz")
        .and_then(|size| size.attribute("val"))
        .and_then(|value| value.parse::<f64>().ok())
    {
        node.format
            .insert("size".into(), format!("{}pt", size / 2.0));
    }
    copy_child_value(properties, "color", "color", node);
    if let Some(highlight) = properties
        .child("highlight")
        .and_then(|property| property.attribute("val"))
    {
        node.format.insert(
            "highlight".into(),
            normalize_highlight_name(highlight).into(),
        );
    }
    if let Some(underline) = properties.child("u") {
        node.format.insert(
            "underline".into(),
            underline.attribute("val").unwrap_or("single").into(),
        );
    }
    copy_child_value(properties, "vertAlign", "script", node);
    if let Some(language) = properties.child("lang") {
        for (attribute, key) in [
            ("val", "language"),
            ("eastAsia", "language.eastAsia"),
            ("bidi", "language.complexScript"),
        ] {
            if let Some(value) = language.attribute(attribute) {
                node.format.insert(key.into(), value.into());
            }
        }
    }
}

fn normalize_highlight_name(value: &str) -> &str {
    match value {
        "darkBlue" => "dark-blue",
        "darkCyan" => "dark-cyan",
        "darkGray" => "dark-gray",
        "darkGreen" => "dark-green",
        "darkMagenta" => "dark-magenta",
        "darkRed" => "dark-red",
        "darkYellow" => "dark-yellow",
        "lightGray" => "light-gray",
        _ => value,
    }
}

fn copy_child_value(parent: &XmlElement, child_name: &str, key: &str, node: &mut DocumentNode) {
    if let Some(value) = parent
        .child(child_name)
        .and_then(|child| child.attribute("val"))
    {
        node.format.insert(key.to_string(), value.to_string());
    }
}

fn bool_value(element: &XmlElement) -> bool {
    !matches!(
        element
            .attribute("val")
            .map(str::to_ascii_lowercase)
            .as_deref(),
        Some("false" | "0" | "off" | "no")
    )
}

fn word_text(element: &XmlElement) -> String {
    let mut output = String::new();
    append_word_text(element, &mut output);
    output
}

fn append_word_text(element: &XmlElement, output: &mut String) {
    match (
        is_word_namespace(element.namespace.as_deref()),
        element.local_name.as_str(),
    ) {
        (true, "t" | "delText" | "instrText") => output.push_str(&direct_text(element)),
        (true, "tab") => output.push('\t'),
        (true, "br" | "cr") => output.push('\n'),
        (true, "noBreakHyphen") => output.push('\u{2011}'),
        (true, "softHyphen") => output.push('\u{00ad}'),
        _ => {
            for child in element.child_elements() {
                append_word_text(child, output);
            }
        }
    }
}

fn is_word_namespace(namespace: Option<&str>) -> bool {
    matches!(namespace, Some(WORD_NAMESPACE | STRICT_WORD_NAMESPACE))
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

fn join_block_text(children: &[DocumentNode]) -> String {
    children
        .iter()
        .map(|child| child.text.as_str())
        .collect::<Vec<_>>()
        .join("\n")
}

fn require_word_element(element: &XmlElement, local_name: &str, part: &str) -> UseResult<()> {
    if element.local_name == local_name
        && matches!(
            element.namespace.as_deref(),
            Some(WORD_NAMESPACE | STRICT_WORD_NAMESPACE)
        )
    {
        return Ok(());
    }
    Err(semantic_error(
        "use.office.word_xml_invalid",
        format!("Word part '{part}' has an unexpected root element."),
    ))
}
