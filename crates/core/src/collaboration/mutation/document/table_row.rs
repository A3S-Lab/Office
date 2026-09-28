use std::collections::HashMap;
use std::sync::Arc;

use a3s_use_core::UseResult;
use yrs::types::xml::XmlIn;
use yrs::{
    GetString, ReadTxn, Transact, Xml, XmlElementPrelim, XmlElementRef, XmlFragment, XmlOut,
    XmlTextPrelim,
};

use super::super::super::{
    collaboration_error, NativeOfficeCollaborationManifest,
    NativeOfficeCollaborationParagraphPosition,
};
use super::super::utf16_len;
use super::identity::{
    ancestor_table_rows, document_identity_attribute, table_row_text_id_rotations,
    validate_paragraph_id_input, PARAGRAPH_ID_ATTRIBUTE, ROW_ID_ATTRIBUTE, ROW_TEXT_ID_ATTRIBUTE,
    TEXT_ID_ATTRIBUTE,
};

const MAX_DOCUMENT_PARAGRAPH_TEXT_UTF16: u32 = 1_048_576;
const MAX_DOCUMENT_ANCESTOR_DEPTH: usize = 64;

pub(super) fn validate_insert_table_row(
    anchor_row_id: &str,
    expected_row_text_id: &str,
    row_id: &str,
    row_text_id: &str,
    paragraph_id: &str,
    text_id: &str,
    text: &str,
) -> UseResult<()> {
    validate_paragraph_id_input(anchor_row_id, "anchorRowId")?;
    validate_paragraph_id_input(expected_row_text_id, "expectedRowTextId")?;
    validate_paragraph_id_input(row_id, "rowId")?;
    validate_paragraph_id_input(row_text_id, "rowTextId")?;
    validate_paragraph_id_input(paragraph_id, "paragraphId")?;
    validate_paragraph_id_input(text_id, "textId")?;
    if anchor_row_id == row_id {
        return Err(invalid_table_row(
            "A new Document table row ID must differ from its anchor row ID.",
        ));
    }
    if utf16_len(text)? > MAX_DOCUMENT_PARAGRAPH_TEXT_UTF16 {
        return Err(collaboration_error(
            "office.collaboration.mutation_too_large",
            "The new Document table row text exceeds the supported UTF-16 length.",
        ));
    }
    Ok(())
}

pub(super) fn validate_delete_table_row(
    row_id: &str,
    expected_row_text_id: &str,
    paragraph_id: &str,
    expected_text_id: &str,
    expected_text: &str,
) -> UseResult<()> {
    validate_paragraph_id_input(row_id, "rowId")?;
    validate_paragraph_id_input(expected_row_text_id, "expectedRowTextId")?;
    validate_paragraph_id_input(paragraph_id, "paragraphId")?;
    validate_paragraph_id_input(expected_text_id, "expectedTextId")?;
    if utf16_len(expected_text)? > MAX_DOCUMENT_PARAGRAPH_TEXT_UTF16 {
        return Err(collaboration_error(
            "office.collaboration.mutation_too_large",
            "The observed Document table row text exceeds the supported UTF-16 length.",
        ));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
pub(super) fn insert_table_row(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    anchor_row_id: &str,
    expected_row_text_id: &str,
    position: NativeOfficeCollaborationParagraphPosition,
    row_id: &str,
    row_text_id: &str,
    paragraph_id: &str,
    text_id: &str,
    text: &str,
) -> UseResult<()> {
    let root = format!("{}.document.content", manifest.namespace);
    let fragment = doc.get_or_insert_xml_fragment(root);
    let transaction = doc.transact();
    let anchor = table_row(&fragment, &transaction, anchor_row_id)?;
    let current_row_text_id = required_row_text_id(&anchor, &transaction)?;
    if current_row_text_id != expected_row_text_id {
        return Err(collaboration_error(
            "office.collaboration.mutation_match_conflict",
            "The target Document table row identity changed after it was inspected.",
        )
        .with_detail("currentRowTextId", current_row_text_id)
        .with_detail("expectedRowTextId", expected_row_text_id));
    }
    let table = parent_table(&anchor)?;
    assert_table_reaches_section(&fragment, &transaction, &table)?;
    ensure_identity_available(&fragment, &transaction, ROW_ID_ATTRIBUTE, row_id)?;
    ensure_identity_available(
        &fragment,
        &transaction,
        PARAGRAPH_ID_ATTRIBUTE,
        paragraph_id,
    )?;
    let rotations = table_row_text_id_rotations(
        &ancestor_table_rows(std::slice::from_ref(&anchor))?,
        &transaction,
    )?;
    let anchor_index = child_index(&table, &transaction, &anchor)?;
    let insert_index = match position {
        NativeOfficeCollaborationParagraphPosition::Before => anchor_index,
        NativeOfficeCollaborationParagraphPosition::After => {
            anchor_index.checked_add(1).ok_or_else(table_too_large)?
        }
    };
    drop(transaction);

    let row = new_row(row_id, row_text_id, paragraph_id, text_id, text);
    let mut transaction = doc.transact_mut();
    table.insert(&mut transaction, insert_index, row);
    for (ancestor, next_text_id) in rotations {
        ancestor.insert_attribute(&mut transaction, ROW_TEXT_ID_ATTRIBUTE, next_text_id);
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
pub(super) fn delete_table_row(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    row_id: &str,
    expected_row_text_id: &str,
    paragraph_id: &str,
    expected_text_id: &str,
    expected_text: &str,
) -> UseResult<()> {
    let root = format!("{}.document.content", manifest.namespace);
    let fragment = doc.get_or_insert_xml_fragment(root);
    let transaction = doc.transact();
    let row = table_row(&fragment, &transaction, row_id)?;
    let current_row_text_id = required_row_text_id(&row, &transaction)?;
    if current_row_text_id != expected_row_text_id {
        return Err(collaboration_error(
            "office.collaboration.mutation_match_conflict",
            "The target Document table row identity changed after it was inspected.",
        )
        .with_detail("currentRowTextId", current_row_text_id)
        .with_detail("expectedRowTextId", expected_row_text_id));
    }
    let table = parent_table(&row)?;
    assert_table_reaches_section(&fragment, &transaction, &table)?;
    let row_count = table_row_count(&table, &transaction)?;
    if row_count <= 1 {
        return Err(structure_conflict(
            "A Document table must retain at least one row.",
        ));
    }
    prove_single_paragraph_row(
        &row,
        &transaction,
        paragraph_id,
        expected_text_id,
        expected_text,
    )?;
    let rotations = table_row_text_id_rotations(
        &ancestor_table_rows(std::slice::from_ref(&row))?,
        &transaction,
    )?;
    let index = child_index(&table, &transaction, &row)?;
    drop(transaction);

    let mut transaction = doc.transact_mut();
    table.remove(&mut transaction, index);
    for (ancestor, next_text_id) in rotations {
        ancestor.insert_attribute(&mut transaction, ROW_TEXT_ID_ATTRIBUTE, next_text_id);
    }
    Ok(())
}

fn new_row(
    row_id: &str,
    row_text_id: &str,
    paragraph_id: &str,
    text_id: &str,
    text: &str,
) -> XmlElementPrelim {
    let mut paragraph_attributes = HashMap::new();
    paragraph_attributes.insert(Arc::from(PARAGRAPH_ID_ATTRIBUTE), paragraph_id.to_owned());
    paragraph_attributes.insert(Arc::from(TEXT_ID_ATTRIBUTE), text_id.to_owned());
    let paragraph_children = if text.is_empty() {
        Vec::new()
    } else {
        vec![XmlIn::from(XmlTextPrelim::new(text))]
    };
    let paragraph = XmlElementPrelim {
        tag: Arc::from("paragraph"),
        attributes: paragraph_attributes,
        children: paragraph_children,
    };
    let cell = XmlElementPrelim {
        tag: Arc::from("tableCell"),
        attributes: HashMap::new(),
        children: vec![XmlIn::from(paragraph)],
    };
    let mut row_attributes = HashMap::new();
    row_attributes.insert(Arc::from(ROW_ID_ATTRIBUTE), row_id.to_owned());
    row_attributes.insert(Arc::from(ROW_TEXT_ID_ATTRIBUTE), row_text_id.to_owned());
    XmlElementPrelim {
        tag: Arc::from("tableRow"),
        attributes: row_attributes,
        children: vec![XmlIn::from(cell)],
    }
}

fn table_row<T: ReadTxn>(
    fragment: &impl XmlFragment,
    transaction: &T,
    row_id: &str,
) -> UseResult<XmlElementRef> {
    let mut matches = Vec::new();
    for node in fragment.successors(transaction) {
        let XmlOut::Element(element) = node else {
            continue;
        };
        if element.tag().as_ref() != "tableRow" {
            continue;
        }
        if document_identity_attribute(&element, transaction, ROW_ID_ATTRIBUTE)?.as_deref()
            == Some(row_id)
        {
            matches.push(element);
        }
    }
    match matches.len() {
        0 => Err(collaboration_error(
            "office.collaboration.mutation_target_missing",
            format!("Document table row '{row_id}' does not exist."),
        )),
        1 => Ok(matches.pop().expect("length checked")),
        count => Err(collaboration_error(
            "office.collaboration.mutation_identity_conflict",
            format!("Document table row ID '{row_id}' is assigned to {count} live nodes."),
        )),
    }
}

fn required_row_text_id<T: ReadTxn>(row: &XmlElementRef, transaction: &T) -> UseResult<String> {
    document_identity_attribute(row, transaction, ROW_TEXT_ID_ATTRIBUTE)?.ok_or_else(|| {
        collaboration_error(
            "office.collaboration.content_invalid",
            "The target Document table row does not have a valid Word row identity.",
        )
    })
}

fn parent_table(row: &XmlElementRef) -> UseResult<XmlElementRef> {
    match row.parent() {
        Some(XmlOut::Element(parent)) if parent.tag().as_ref() == "table" => Ok(parent),
        _ => Err(structure_conflict(
            "A Document table row must be a direct child of a table.",
        )),
    }
}

fn assert_table_reaches_section<T: ReadTxn>(
    fragment: &impl XmlFragment,
    transaction: &T,
    table: &XmlElementRef,
) -> UseResult<()> {
    let mut current = match table.parent() {
        Some(XmlOut::Element(parent)) => parent,
        _ => {
            return Err(structure_conflict(
                "The target Document table is detached from a supported block container.",
            ))
        }
    };
    for _ in 0..MAX_DOCUMENT_ANCESTOR_DEPTH {
        match current.tag().as_ref() {
            "documentSection" => {
                if fragment
                    .children(transaction)
                    .any(|node| matches!(node, XmlOut::Element(candidate) if candidate == current))
                {
                    return Ok(());
                }
                return Err(structure_conflict(
                    "The target Document table is not contained by a top-level section.",
                ));
            }
            "listItem" => {
                let list = parent_element(&current, "A Document list item is detached.")?;
                if !matches!(list.tag().as_ref(), "bulletList" | "orderedList") {
                    return Err(structure_conflict(
                        "A Document list item must be a direct child of a bullet or ordered list.",
                    ));
                }
                current = parent_element(&list, "The target Document list is detached.")?;
            }
            "tableCell" | "tableHeader" => {
                let row = parent_element(&current, "A Document table cell is detached.")?;
                if row.tag().as_ref() != "tableRow" {
                    return Err(structure_conflict(
                        "A Document table cell must be a direct child of a table row.",
                    ));
                }
                let table = parent_element(&row, "A Document table row is detached.")?;
                if table.tag().as_ref() != "table" {
                    return Err(structure_conflict(
                        "A Document table row must be a direct child of a table.",
                    ));
                }
                current = parent_element(&table, "The target Document table is detached.")?;
            }
            "blockquote" => {
                current = parent_element(&current, "The target Document blockquote is detached.")?;
            }
            _ => {
                return Err(structure_conflict(
                    "The target Document table is not inside a supported section, list item, table cell, or blockquote.",
                ))
            }
        }
    }
    Err(collaboration_error(
        "office.collaboration.mutation_too_large",
        "The target Document table exceeds the supported structural depth.",
    ))
}

fn prove_single_paragraph_row<T: ReadTxn>(
    row: &XmlElementRef,
    transaction: &T,
    paragraph_id: &str,
    expected_text_id: &str,
    expected_text: &str,
) -> UseResult<()> {
    let mut cells = Vec::new();
    for child in row.children(transaction) {
        let XmlOut::Element(element) = child else {
            return Err(structure_conflict(
                "A Document table row contains a child that is not a cell.",
            ));
        };
        if !matches!(element.tag().as_ref(), "tableCell" | "tableHeader") {
            return Err(structure_conflict(
                "A Document table row contains a child that is not a cell.",
            ));
        }
        cells.push(element);
    }
    if cells.len() != 1 {
        return Err(structure_conflict(
            "A Document table row with more than one cell cannot be deleted by this frame.",
        ));
    }
    let mut blocks = Vec::new();
    for child in cells[0].children(transaction) {
        let XmlOut::Element(element) = child else {
            return Err(structure_conflict(
                "A Document table cell contains a child that is not a block.",
            ));
        };
        blocks.push(element);
    }
    if blocks.len() != 1 || blocks[0].tag().as_ref() != "paragraph" {
        return Err(structure_conflict(
            "A Document table row can be deleted only when its single cell holds one plain paragraph.",
        ));
    }
    let paragraph = &blocks[0];
    let current_paragraph_id =
        document_identity_attribute(paragraph, transaction, PARAGRAPH_ID_ATTRIBUTE)?.ok_or_else(
            || {
                structure_conflict(
            "The Document table row paragraph does not have a valid Word paragraph identity.",
        )
            },
        )?;
    if current_paragraph_id != paragraph_id {
        return Err(collaboration_error(
            "office.collaboration.mutation_match_conflict",
            "The target Document table row paragraph changed after it was inspected.",
        ));
    }
    let current_text_id = document_identity_attribute(paragraph, transaction, TEXT_ID_ATTRIBUTE)?
        .ok_or_else(|| {
        structure_conflict(
            "The Document table row paragraph does not have a valid Word text identity.",
        )
    })?;
    if current_text_id != expected_text_id {
        return Err(collaboration_error(
            "office.collaboration.mutation_match_conflict",
            "The target Document table row paragraph text identity changed after it was inspected.",
        ));
    }
    let mut text = String::new();
    for child in paragraph.children(transaction) {
        let XmlOut::Text(child) = child else {
            return Err(structure_conflict(
                "A Document table row paragraph containing inline atoms cannot be deleted by this frame.",
            ));
        };
        text.push_str(&child.get_string(transaction));
    }
    if text != expected_text {
        return Err(collaboration_error(
            "office.collaboration.mutation_match_conflict",
            "The target Document table row text changed after it was inspected.",
        ));
    }
    Ok(())
}

fn table_row_count<T: ReadTxn>(table: &XmlElementRef, transaction: &T) -> UseResult<u32> {
    let mut count = 0_u32;
    for child in table.children(transaction) {
        let XmlOut::Element(element) = child else {
            return Err(structure_conflict(
                "A Document table contains a child that is not a row.",
            ));
        };
        if element.tag().as_ref() != "tableRow" {
            return Err(structure_conflict(
                "A Document table contains a child that is not a row.",
            ));
        }
        count = count.checked_add(1).ok_or_else(table_too_large)?;
    }
    Ok(count)
}

fn ensure_identity_available<T: ReadTxn>(
    fragment: &impl XmlFragment,
    transaction: &T,
    attribute: &str,
    value: &str,
) -> UseResult<()> {
    for node in fragment.successors(transaction) {
        let XmlOut::Element(element) = node else {
            continue;
        };
        if document_identity_attribute(&element, transaction, attribute)?.as_deref() == Some(value)
        {
            return Err(collaboration_error(
                "office.collaboration.mutation_identity_conflict",
                format!("Document identity '{value}' is already in use."),
            ));
        }
    }
    Ok(())
}

fn child_index<T: ReadTxn>(
    container: &XmlElementRef,
    transaction: &T,
    target: &XmlElementRef,
) -> UseResult<u32> {
    container
        .children(transaction)
        .enumerate()
        .find_map(|(index, node)| {
            matches!(node, XmlOut::Element(candidate) if candidate == *target)
                .then(|| u32::try_from(index).ok())
                .flatten()
        })
        .ok_or_else(|| {
            structure_conflict("The target Document table row cannot be located inside its table.")
        })
}

fn parent_element(element: &XmlElementRef, message: &str) -> UseResult<XmlElementRef> {
    match element.parent() {
        Some(XmlOut::Element(parent)) => Ok(parent),
        _ => Err(structure_conflict(message)),
    }
}

fn invalid_table_row(message: &str) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_invalid", message)
}

fn structure_conflict(message: &str) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_structure_conflict", message)
}

fn table_too_large() -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.mutation_too_large",
        "The target Document table contains too many rows for a structural mutation.",
    )
}
