use a3s_use_core::UseResult;
use yrs::{ReadTxn, Xml, XmlElementRef, XmlFragment, XmlOut};

use super::super::super::collaboration_error;
use super::identity::{
    document_identity_attribute, is_identity_paragraph_tag, PARAGRAPH_ID_ATTRIBUTE,
};

const MAX_DOCUMENT_CONTAINER_DEPTH: usize = 64;
const MAX_DOCUMENT_CONTAINER_BLOCKS: u32 = 1_048_576;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum ParagraphContainerKind {
    Section,
    ListItem,
    TableCell,
    TableHeader,
    Blockquote,
}

pub(super) struct StructuralParagraph {
    pub container: XmlElementRef,
    pub paragraph: XmlElementRef,
    pub index: u32,
    pub container_kind: ParagraphContainerKind,
}

pub(super) fn structural_paragraph<T: ReadTxn>(
    fragment: &impl XmlFragment,
    transaction: &T,
    paragraph_id: &str,
) -> UseResult<StructuralParagraph> {
    let mut matches = Vec::new();
    for node in fragment.successors(transaction) {
        let XmlOut::Element(element) = node else {
            continue;
        };
        if document_identity_attribute(&element, transaction, PARAGRAPH_ID_ATTRIBUTE)?.as_deref()
            == Some(paragraph_id)
        {
            matches.push(element);
        }
    }
    let paragraph = match matches.len() {
        0 => {
            return Err(collaboration_error(
                "office.collaboration.mutation_target_missing",
                format!("Document paragraph '{paragraph_id}' does not exist."),
            ))
        }
        1 => matches.pop().ok_or_else(|| {
            collaboration_error(
                "office.collaboration.mutation_target_missing",
                format!("Document paragraph '{paragraph_id}' does not exist."),
            )
        })?,
        count => {
            return Err(collaboration_error(
                "office.collaboration.mutation_identity_conflict",
                format!(
                    "Document paragraph ID '{paragraph_id}' is assigned to {count} live nodes."
                ),
            ))
        }
    };
    if !is_identity_paragraph_tag(paragraph.tag()) {
        return Err(structure_conflict(
            "The target Word paragraph identity belongs to an unsupported node type.",
        ));
    }
    let container = parent_element(
        &paragraph,
        "The target paragraph does not have a supported structural parent.",
    )?;
    let container_kind = container_kind(&container)?;
    validate_container_ancestry(fragment, transaction, &container)?;
    validate_container_shape(&container, transaction, container_kind, None)?;
    let index = child_index(&container, transaction, &paragraph)?;
    Ok(StructuralParagraph {
        container,
        paragraph,
        index,
        container_kind,
    })
}

pub(super) fn ensure_paragraph_id_available<T: ReadTxn>(
    fragment: &impl XmlFragment,
    transaction: &T,
    paragraph_id: &str,
) -> UseResult<()> {
    for node in fragment.successors(transaction) {
        let XmlOut::Element(element) = node else {
            continue;
        };
        if document_identity_attribute(&element, transaction, PARAGRAPH_ID_ATTRIBUTE)?.as_deref()
            == Some(paragraph_id)
        {
            return Err(collaboration_error(
                "office.collaboration.mutation_identity_conflict",
                format!("Document paragraph ID '{paragraph_id}' is already in use."),
            ));
        }
    }
    Ok(())
}

pub(super) fn validate_paragraph_deletion<T: ReadTxn>(
    target: &StructuralParagraph,
    transaction: &T,
) -> UseResult<()> {
    validate_container_shape(
        &target.container,
        transaction,
        target.container_kind,
        Some(target.index),
    )
}

pub(super) fn validate_paragraph_insertion<T: ReadTxn>(
    target: &StructuralParagraph,
    transaction: &T,
) -> UseResult<()> {
    if target.container.len(transaction) >= MAX_DOCUMENT_CONTAINER_BLOCKS {
        return Err(container_too_large());
    }
    Ok(())
}

fn validate_container_ancestry<T: ReadTxn>(
    fragment: &impl XmlFragment,
    transaction: &T,
    container: &XmlElementRef,
) -> UseResult<()> {
    let mut current = container.clone();
    for _ in 0..MAX_DOCUMENT_CONTAINER_DEPTH {
        match current.tag().as_ref() {
            "documentSection" => {
                if fragment.children(transaction).any(
                    |node| matches!(node, XmlOut::Element(candidate) if candidate == current),
                ) {
                    return Ok(());
                }
                return Err(structure_conflict(
                    "The target paragraph is not contained by a top-level Document section.",
                ));
            }
            "listItem" => {
                let list = parent_element(
                    &current,
                    "A Document list item is detached from its list container.",
                )?;
                if !matches!(list.tag().as_ref(), "bulletList" | "orderedList") {
                    return Err(structure_conflict(
                        "A Document list item must be a direct child of a bullet or ordered list.",
                    ));
                }
                current = parent_element(
                    &list,
                    "The target Document list is detached from a supported block container.",
                )?;
            }
            "tableCell" | "tableHeader" => {
                let row = parent_element(
                    &current,
                    "A Document table cell is detached from its row.",
                )?;
                if row.tag().as_ref() != "tableRow" {
                    return Err(structure_conflict(
                        "A Document table cell must be a direct child of a table row.",
                    ));
                }
                let table = parent_element(&row, "A Document table row is detached from its table.")?;
                if table.tag().as_ref() != "table" {
                    return Err(structure_conflict(
                        "A Document table row must be a direct child of a table.",
                    ));
                }
                current = parent_element(
                    &table,
                    "The target Document table is detached from a supported block container.",
                )?;
            }
            "blockquote" => {
                current = parent_element(
                    &current,
                    "The target Document blockquote is detached from a supported block container.",
                )?;
            }
            _ => {
                return Err(structure_conflict(
                    "The target paragraph is not inside a supported Document section, list item, table cell, or blockquote container.",
                ))
            }
        }
    }
    Err(collaboration_error(
        "office.collaboration.mutation_too_large",
        "The target Document paragraph exceeds the supported structural depth.",
    ))
}

fn validate_container_shape<T: ReadTxn>(
    container: &XmlElementRef,
    transaction: &T,
    kind: ParagraphContainerKind,
    removed_index: Option<u32>,
) -> UseResult<()> {
    let mut remaining = 0_u32;
    let mut first_tag = None;
    for (index, node) in container.children(transaction).enumerate() {
        let index = u32::try_from(index).map_err(|_| container_too_large())?;
        let XmlOut::Element(element) = node else {
            return Err(structure_conflict(
                "The target paragraph container contains a non-block child.",
            ));
        };
        if removed_index == Some(index) {
            continue;
        }
        remaining = remaining.checked_add(1).ok_or_else(container_too_large)?;
        if remaining > MAX_DOCUMENT_CONTAINER_BLOCKS {
            return Err(container_too_large());
        }
        if first_tag.is_none() {
            first_tag = Some(element.tag().to_string());
        }
    }
    if remaining == 0 {
        return Err(structure_conflict(match kind {
            ParagraphContainerKind::Section => "A Document section must retain at least one block.",
            ParagraphContainerKind::ListItem => {
                "A Document list item must retain its leading paragraph."
            }
            ParagraphContainerKind::TableCell | ParagraphContainerKind::TableHeader => {
                "A Document table cell must retain at least one block."
            }
            ParagraphContainerKind::Blockquote => {
                "A Document blockquote must retain at least one block."
            }
        }));
    }
    if kind == ParagraphContainerKind::ListItem && first_tag.as_deref() != Some("paragraph") {
        return Err(structure_conflict(
            "A Document list item must begin with a plain paragraph.",
        ));
    }
    Ok(())
}

fn container_kind(container: &XmlElementRef) -> UseResult<ParagraphContainerKind> {
    match container.tag().as_ref() {
        "documentSection" => Ok(ParagraphContainerKind::Section),
        "listItem" => Ok(ParagraphContainerKind::ListItem),
        "tableCell" => Ok(ParagraphContainerKind::TableCell),
        "tableHeader" => Ok(ParagraphContainerKind::TableHeader),
        "blockquote" => Ok(ParagraphContainerKind::Blockquote),
        _ => Err(structure_conflict(
            "The target paragraph is not a direct child of a supported Document block container.",
        )),
    }
}

fn parent_element(element: &XmlElementRef, message: &str) -> UseResult<XmlElementRef> {
    match element.parent() {
        Some(XmlOut::Element(parent)) => Ok(parent),
        _ => Err(structure_conflict(message)),
    }
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
            structure_conflict(
                "The target paragraph cannot be located inside its Document container.",
            )
        })
}

fn structure_conflict(message: &str) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_structure_conflict", message)
}

fn container_too_large() -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.mutation_too_large",
        "The target Document container contains too many blocks for a structural mutation.",
    )
}
