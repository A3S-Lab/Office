use std::collections::HashSet;

use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;
use yrs::{Any, Array, Map, MapRef, Out, ReadTxn, Transact};

use super::json::{json_to_any, validate_presentation_identifier};
use super::state::read_element_state;
use super::{
    invalid_presentation_mutation, invalid_shared_presentation, presentation_match_conflict,
    NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation,
    NativeOfficeCollaborationPresentationContainerKind,
};
use crate::collaboration::NativeOfficeCollaborationPresentationGroupMember;

const MAX_GROUP_MEMBERS: usize = 64;
const MAX_GROUP_DEPTH: usize = 8;

pub(super) fn validate_deck_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::PresentationMoveSlide {
            slide_id,
            expected_after_slide_id,
            after_slide_id,
        } => {
            validate_presentation_identifier(slide_id, "slide", false)?;
            validate_slide_anchor(
                expected_after_slide_id.as_deref(),
                slide_id,
                "expected predecessor",
            )?;
            validate_slide_anchor(
                after_slide_id.as_deref(),
                slide_id,
                "destination predecessor",
            )?;
            Ok(())
        }
        NativeOfficeCollaborationMutation::PresentationSetGroup {
            container_id,
            members,
            ..
        } => {
            validate_presentation_identifier(container_id, "container", false)?;
            if members.is_empty() {
                return Err(invalid_presentation_mutation(
                    "A Presentation group frame must name at least one scene element.",
                ));
            }
            if members.len() > MAX_GROUP_MEMBERS {
                return Err(invalid_presentation_mutation(
                    "A Presentation group frame names more than 64 scene elements.",
                ));
            }
            let mut seen = HashSet::new();
            for member in members {
                validate_presentation_identifier(&member.element_id, "scene element", false)?;
                if !seen.insert(member.element_id.as_str()) {
                    return Err(invalid_presentation_mutation(format!(
                        "A Presentation group frame names scene element ID '{}' more than once.",
                        member.element_id
                    )));
                }
                validate_group_path(&member.expected_group_ids)?;
                validate_group_path(&member.next_group_ids)?;
            }
            Ok(())
        }
        NativeOfficeCollaborationMutation::PresentationSetBackground {
            container_id,
            expected_background,
            next_background,
            ..
        } => {
            validate_presentation_identifier(container_id, "container", false)?;
            validate_background(expected_background.as_deref())?;
            validate_background(next_background.as_deref())?;
            Ok(())
        }
        _ => Err(invalid_presentation_mutation(
            "The supplied mutation is not a Presentation deck mutation.",
        )),
    }
}

pub(super) fn apply_deck_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::PresentationMoveSlide {
            slide_id,
            expected_after_slide_id,
            after_slide_id,
        } => move_slide(
            doc,
            manifest,
            slide_id,
            expected_after_slide_id.as_deref(),
            after_slide_id.as_deref(),
        ),
        NativeOfficeCollaborationMutation::PresentationSetGroup {
            container_kind,
            container_id,
            members,
        } => set_group(doc, manifest, *container_kind, container_id, members),
        NativeOfficeCollaborationMutation::PresentationSetBackground {
            container_kind,
            container_id,
            expected_background,
            next_background,
        } => set_background(
            doc,
            manifest,
            *container_kind,
            container_id,
            expected_background,
            next_background,
        ),
        _ => Err(invalid_presentation_mutation(
            "The supplied mutation is not a Presentation deck mutation.",
        )),
    }
}

fn move_slide(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    slide_id: &str,
    expected_after_slide_id: Option<&str>,
    after_slide_id: Option<&str>,
) -> UseResult<()> {
    let deck = read_slide_deck(doc, manifest)?;
    if !deck.active.iter().any(|id| id == slide_id) {
        return Err(presentation_match_conflict(format!(
            "Presentation slide ID '{slide_id}' is not an active slide."
        )));
    }
    if let Some(anchor_id) = after_slide_id {
        if !deck.active.iter().any(|id| id == anchor_id) {
            return Err(presentation_match_conflict(format!(
                "Presentation destination predecessor slide ID '{anchor_id}' is not an active slide."
            )));
        }
    }
    let current = predecessor(&deck.active, slide_id)?;
    if current.as_deref() != expected_after_slide_id {
        return Err(presentation_match_conflict(format!(
            "Presentation slide ID '{slide_id}' moved after it was observed."
        )));
    }
    if current.as_deref() == after_slide_id {
        return Ok(());
    }

    let remaining = deck
        .raw
        .iter()
        .filter(|id| id.as_str() != slide_id)
        .collect::<Vec<_>>();
    let insertion_index = match after_slide_id {
        Some(anchor_id) => remaining
            .iter()
            .position(|id| id.as_str() == anchor_id)
            .map(|index| index + 1)
            .ok_or_else(|| {
                invalid_shared_presentation(
                    "The shared Presentation slide order omits an active destination predecessor.",
                )
            })?,
        None => 0,
    };
    let insertion_index = u32::try_from(insertion_index)
        .map_err(|_| invalid_presentation_mutation("The Presentation slide order is too long."))?;
    let mut removals = Vec::new();
    for index in (0..deck.raw.len()).rev() {
        if deck.raw[index] == slide_id {
            removals.push(u32::try_from(index).map_err(|_| {
                invalid_presentation_mutation("The Presentation slide order is too long.")
            })?);
        }
    }

    let mut transaction = doc.transact_mut();
    for index in removals {
        deck.order.remove_range(&mut transaction, index, 1);
    }
    deck.order
        .insert(&mut transaction, insertion_index, slide_id);
    drop(transaction);

    let next = read_slide_deck(doc, manifest)?;
    if predecessor(&next.active, slide_id)?.as_deref() != after_slide_id {
        return Err(invalid_shared_presentation(
            "The Presentation slide move did not produce the requested order.",
        ));
    }
    Ok(())
}

fn set_group(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    container_kind: NativeOfficeCollaborationPresentationContainerKind,
    container_id: &str,
    members: &[NativeOfficeCollaborationPresentationGroupMember],
) -> UseResult<()> {
    let state = read_element_state(doc, manifest, container_kind, container_id)?;
    let mut prepared = Vec::with_capacity(members.len());
    let mut changed = false;
    for member in members {
        let record = state.records.get(&member.element_id).ok_or_else(|| {
            presentation_match_conflict(format!(
                "Presentation scene element ID '{}' does not exist.",
                member.element_id
            ))
        })?;
        if record.tombstoned {
            return Err(presentation_match_conflict(format!(
                "Presentation scene element ID '{}' was deleted.",
                member.element_id
            )));
        }
        let current = group_path(&record.value)?;
        if current != member.expected_group_ids {
            return Err(presentation_match_conflict(format!(
                "Presentation scene element ID '{}' group path changed after it was observed.",
                member.element_id
            )));
        }
        if current != member.next_group_ids {
            changed = true;
        }
        let next = if member.next_group_ids.is_empty() {
            None
        } else {
            let values = member
                .next_group_ids
                .iter()
                .cloned()
                .map(JsonValue::String)
                .collect();
            Some(json_to_any(&JsonValue::Array(values), 0)?)
        };
        prepared.push((record.map.clone(), next));
    }
    if !changed {
        return Ok(());
    }

    let mut transaction = doc.transact_mut();
    for (map, next) in prepared {
        match next {
            Some(value) => {
                map.insert(&mut transaction, "groupIds", value);
            }
            None => {
                map.remove(&mut transaction, "groupIds");
            }
        }
    }
    drop(transaction);
    Ok(())
}

fn set_background(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    container_kind: NativeOfficeCollaborationPresentationContainerKind,
    container_id: &str,
    expected_background: &Option<String>,
    next_background: &Option<String>,
) -> UseResult<()> {
    let (container, current) = read_container(doc, manifest, container_kind, container_id)?;
    if &current != expected_background {
        return Err(presentation_match_conflict(format!(
            "Presentation {} ID '{container_id}' background changed after it was observed.",
            container_kind.as_str()
        )));
    }
    if &current == next_background {
        return Ok(());
    }

    let mut transaction = doc.transact_mut();
    match next_background {
        Some(value) => {
            container.insert(&mut transaction, "background", value.as_str());
        }
        None => {
            container.remove(&mut transaction, "background");
        }
    }
    drop(transaction);
    Ok(())
}

struct SlideDeck {
    order: yrs::ArrayRef,
    raw: Vec<String>,
    active: Vec<String>,
}

fn read_slide_deck(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<SlideDeck> {
    let slides = doc.get_or_insert_map(format!("{}.presentation.slides", manifest.namespace));
    let order = doc.get_or_insert_array(format!("{}.presentation.slide-order", manifest.namespace));
    let transaction = doc.transact();
    let mut raw = Vec::new();
    let mut active = Vec::new();
    let mut seen = HashSet::new();
    for index in 0..order.len(&transaction) {
        let id = match order.get(&transaction, index) {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            _ => {
                return Err(invalid_shared_presentation(
                    "The shared Presentation slide order contains a non-string identity.",
                ))
            }
        };
        validate_presentation_identifier(&id, "slide", true)?;
        if seen.insert(id.clone()) {
            if let Some(Out::YMap(record)) = slides.get(&transaction, id.as_str()) {
                if !slide_deleted(&record, &transaction) {
                    active.push(id.clone());
                }
            }
        }
        raw.push(id);
    }
    Ok(SlideDeck { order, raw, active })
}

fn read_container(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    container_kind: NativeOfficeCollaborationPresentationContainerKind,
    container_id: &str,
) -> UseResult<(MapRef, Option<String>)> {
    let collection = match container_kind {
        NativeOfficeCollaborationPresentationContainerKind::Slide => "slides",
        NativeOfficeCollaborationPresentationContainerKind::Master => "masters",
        NativeOfficeCollaborationPresentationContainerKind::Layout => "layouts",
    };
    let containers =
        doc.get_or_insert_map(format!("{}.presentation.{collection}", manifest.namespace));
    let transaction = doc.transact();
    let container = match containers.get(&transaction, container_id) {
        Some(Out::YMap(value)) => value,
        Some(_) => {
            return Err(invalid_shared_presentation(format!(
                "The shared Presentation {} '{container_id}' is not a typed map.",
                container_kind.as_str()
            )))
        }
        None => {
            return Err(presentation_match_conflict(format!(
                "Presentation {} ID '{container_id}' does not exist.",
                container_kind.as_str()
            )))
        }
    };
    match container.get(&transaction, "id") {
        Some(Out::Any(Any::String(value))) if value.as_ref() == container_id => {}
        _ => {
            return Err(invalid_shared_presentation(format!(
                "The shared Presentation {} identity does not match its collection key.",
                container_kind.as_str()
            )))
        }
    }
    let background = match container.get(&transaction, "background") {
        None => None,
        Some(Out::Any(Any::String(value))) => Some(value.to_string()),
        Some(_) => {
            return Err(invalid_shared_presentation(
                "The shared Presentation background is not a string.",
            ))
        }
    };
    Ok((container, background))
}

fn slide_deleted(record: &MapRef, transaction: &impl ReadTxn) -> bool {
    matches!(
        record.get(transaction, "tombstone"),
        Some(Out::Any(Any::Bool(true)))
    )
}

fn predecessor(order: &[String], slide_id: &str) -> UseResult<Option<String>> {
    let index = order.iter().position(|id| id == slide_id).ok_or_else(|| {
        invalid_shared_presentation("The shared Presentation slide order omits an active slide.")
    })?;
    Ok(index
        .checked_sub(1)
        .and_then(|index| order.get(index))
        .cloned())
}

fn group_path(value: &JsonValue) -> UseResult<Vec<String>> {
    match value.get("groupIds") {
        None | Some(JsonValue::Null) => Ok(Vec::new()),
        Some(JsonValue::Array(items)) => items
            .iter()
            .map(|item| {
                item.as_str().map(str::to_owned).ok_or_else(|| {
                    invalid_shared_presentation(
                        "A shared Presentation group path entry is not a string.",
                    )
                })
            })
            .collect(),
        Some(_) => Err(invalid_shared_presentation(
            "A shared Presentation group path is not an array.",
        )),
    }
}

fn validate_slide_anchor(value: Option<&str>, slide_id: &str, label: &str) -> UseResult<()> {
    let Some(value) = value else {
        return Ok(());
    };
    validate_presentation_identifier(value, label, false)?;
    if value == slide_id {
        return Err(invalid_presentation_mutation(format!(
            "A Presentation slide cannot use itself as its {label}."
        )));
    }
    Ok(())
}

fn validate_group_path(path: &[String]) -> UseResult<()> {
    if path.len() > MAX_GROUP_DEPTH {
        return Err(invalid_presentation_mutation(
            "A Presentation group path is deeper than 8 groups.",
        ));
    }
    for group_id in path {
        validate_presentation_identifier(group_id, "group", false)?;
    }
    Ok(())
}

fn validate_background(value: Option<&str>) -> UseResult<()> {
    let Some(value) = value else {
        return Ok(());
    };
    let length = value.encode_utf16().count();
    let trimmed = value
        .chars()
        .next()
        .is_some_and(|character| character.is_whitespace())
        || value
            .chars()
            .next_back()
            .is_some_and(|character| character.is_whitespace());
    let control = value.chars().any(|character| character.is_control());
    if (1..=256).contains(&length) && !trimmed && !control {
        return Ok(());
    }
    Err(invalid_presentation_mutation(
        "A Presentation background must contain 1 to 256 UTF-16 code units without surrounding whitespace or control characters.",
    ))
}
