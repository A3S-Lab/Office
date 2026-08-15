use a3s_use_core::UseResult;
use yrs::{Array, Transact};

use super::json::validate_presentation_identifier;
use super::state::read_element_state;
use super::{
    invalid_presentation_mutation, invalid_shared_presentation, presentation_match_conflict,
    NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation,
    NativeOfficeCollaborationPresentationContainerKind,
};

pub(super) fn validate_element_order_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    let NativeOfficeCollaborationMutation::PresentationMoveElement {
        container_id,
        element_id,
        expected_after_element_id,
        after_element_id,
        ..
    } = mutation
    else {
        return Err(invalid_presentation_mutation(
            "The supplied mutation is not a Presentation scene-element order mutation.",
        ));
    };
    validate_presentation_identifier(container_id, "container", false)?;
    validate_presentation_identifier(element_id, "scene element", false)?;
    validate_anchor(
        expected_after_element_id.as_deref(),
        element_id,
        "expected predecessor",
    )?;
    validate_anchor(
        after_element_id.as_deref(),
        element_id,
        "destination predecessor",
    )?;
    Ok(())
}

pub(super) fn apply_element_order_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    let NativeOfficeCollaborationMutation::PresentationMoveElement {
        container_kind,
        container_id,
        element_id,
        expected_after_element_id,
        after_element_id,
    } = mutation
    else {
        return Err(invalid_presentation_mutation(
            "The supplied mutation is not a Presentation scene-element order mutation.",
        ));
    };
    move_element(
        doc,
        manifest,
        *container_kind,
        container_id,
        element_id,
        expected_after_element_id.as_deref(),
        after_element_id.as_deref(),
    )
}

fn move_element(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    container_kind: NativeOfficeCollaborationPresentationContainerKind,
    container_id: &str,
    element_id: &str,
    expected_after_element_id: Option<&str>,
    after_element_id: Option<&str>,
) -> UseResult<()> {
    let state = read_element_state(doc, manifest, container_kind, container_id)?;
    let record = state.records.get(element_id).ok_or_else(|| {
        presentation_match_conflict(format!(
            "Presentation scene element ID '{element_id}' does not exist."
        ))
    })?;
    if record.tombstoned {
        return Err(presentation_match_conflict(format!(
            "Presentation scene element ID '{element_id}' was deleted."
        )));
    }
    if let Some(anchor_id) = after_element_id {
        let anchor = state.records.get(anchor_id).ok_or_else(|| {
            presentation_match_conflict(format!(
                "Presentation destination predecessor scene element ID '{anchor_id}' does not exist."
            ))
        })?;
        if anchor.tombstoned {
            return Err(presentation_match_conflict(format!(
                "Presentation destination predecessor scene element ID '{anchor_id}' was deleted."
            )));
        }
    }

    let current_after_element_id = predecessor(&state.active_order, element_id)?;
    if current_after_element_id == after_element_id {
        return Ok(());
    }
    if current_after_element_id != expected_after_element_id {
        return Err(presentation_match_conflict(format!(
            "Presentation scene element ID '{element_id}' moved after it was observed."
        )));
    }

    let remaining_order = state
        .raw_order
        .iter()
        .filter(|id| id.as_str() != element_id)
        .collect::<Vec<_>>();
    let insertion_index = match after_element_id {
        Some(anchor_id) => remaining_order
            .iter()
            .position(|id| id.as_str() == anchor_id)
            .map(|index| index + 1)
            .ok_or_else(|| {
                invalid_shared_presentation(
                    "The shared Presentation element order omits an active destination predecessor.",
                )
            })?,
        None => 0,
    };

    let mut transaction = doc.transact_mut();
    for index in (0..state.raw_order.len()).rev() {
        if state.raw_order[index] == element_id {
            state.order.remove_range(&mut transaction, index as u32, 1);
        }
    }
    state
        .order
        .insert(&mut transaction, insertion_index as u32, element_id);
    drop(transaction);

    let next = read_element_state(doc, manifest, container_kind, container_id)?;
    if predecessor(&next.active_order, element_id)? != after_element_id {
        return Err(invalid_shared_presentation(
            "The Presentation scene-element move did not produce the requested z-order.",
        ));
    }
    Ok(())
}

fn predecessor<'a>(order: &'a [String], element_id: &str) -> UseResult<Option<&'a str>> {
    let index = order
        .iter()
        .position(|id| id == element_id)
        .ok_or_else(|| {
            invalid_shared_presentation(
                "The shared Presentation element order omits an active scene element.",
            )
        })?;
    Ok(index
        .checked_sub(1)
        .and_then(|index| order.get(index))
        .map(String::as_str))
}

fn validate_anchor(value: Option<&str>, element_id: &str, label: &str) -> UseResult<()> {
    let Some(value) = value else {
        return Ok(());
    };
    validate_presentation_identifier(value, label, false)?;
    if value == element_id {
        return Err(invalid_presentation_mutation(format!(
            "A Presentation scene element cannot use itself as its {label}."
        )));
    }
    Ok(())
}
