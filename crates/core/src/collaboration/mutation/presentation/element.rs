use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;
use yrs::{Array, Map, MapPrelim, Transact};

use super::json::{
    assert_compatible_element, canonical_json, element_field_patches, json_equal, json_to_any,
    validate_element_input, validate_presentation_identifier, PresentationElementPatch,
};
use super::state::{encoded_element_claim, read_element_state};
use super::{
    invalid_presentation_mutation, invalid_shared_presentation, presentation_match_conflict,
    NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation,
    NativeOfficeCollaborationPresentationContainerKind,
};

pub(super) fn validate_element_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::PresentationCreateElement {
            container_id,
            element,
            after_element_id,
            ..
        } => {
            validate_container_id(container_id)?;
            if let Some(after_element_id) = after_element_id {
                validate_presentation_identifier(after_element_id, "anchor scene element", false)?;
            }
            validate_element_input(element, "scene element")?;
            Ok(())
        }
        NativeOfficeCollaborationMutation::PresentationUpdateElement {
            container_id,
            element_id,
            expected_element,
            next_element,
            ..
        } => {
            validate_container_id(container_id)?;
            validate_presentation_identifier(element_id, "scene element", false)?;
            let expected = validate_element_input(expected_element, "expected scene element")?;
            let next = validate_element_input(next_element, "next scene element")?;
            if expected.id != element_id.as_str() || next.id != element_id.as_str() {
                return Err(invalid_presentation_mutation(
                    "A Presentation scene-element update must retain its requested element ID.",
                ));
            }
            if expected.element_type != next.element_type {
                return Err(invalid_presentation_mutation(
                    "A Presentation scene-element update cannot change its type identity.",
                ));
            }
            Ok(())
        }
        NativeOfficeCollaborationMutation::PresentationDeleteElement {
            container_id,
            expected_element,
            ..
        } => {
            validate_container_id(container_id)?;
            validate_element_input(expected_element, "expected scene element")?;
            Ok(())
        }
        _ => Err(invalid_presentation_mutation(
            "The supplied mutation is not a Presentation scene-element mutation.",
        )),
    }
}

pub(super) fn apply_element_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::PresentationCreateElement {
            container_kind,
            container_id,
            element,
            after_element_id,
        } => create_element(
            doc,
            manifest,
            *container_kind,
            container_id,
            element,
            after_element_id.as_deref(),
        ),
        NativeOfficeCollaborationMutation::PresentationUpdateElement {
            container_kind,
            container_id,
            element_id,
            expected_element,
            next_element,
        } => update_element(
            doc,
            manifest,
            *container_kind,
            container_id,
            element_id,
            expected_element,
            next_element,
        ),
        NativeOfficeCollaborationMutation::PresentationDeleteElement {
            container_kind,
            container_id,
            expected_element,
        } => delete_element(
            doc,
            manifest,
            *container_kind,
            container_id,
            expected_element,
        ),
        _ => Err(invalid_presentation_mutation(
            "The supplied mutation is not a Presentation scene-element mutation.",
        )),
    }
}

fn create_element(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    container_kind: NativeOfficeCollaborationPresentationContainerKind,
    container_id: &str,
    element: &JsonValue,
    after_element_id: Option<&str>,
) -> UseResult<()> {
    let identity = validate_element_input(element, "scene element")?;
    let state = read_element_state(doc, manifest, container_kind, container_id)?;
    let fingerprint = canonical_json(element)?;
    let existing_claim = state
        .claims
        .claim_for(container_kind, container_id, &identity.id);
    if let Some(record) = state.records.get(&identity.id) {
        if record.tombstoned {
            return Err(presentation_match_conflict(format!(
                "Deleted Presentation scene element ID '{}' cannot be reused.",
                identity.id
            )));
        }
        if json_equal(&record.value, element)
            && existing_claim.is_none_or(|claim| claim == fingerprint)
        {
            return Ok(());
        }
        return Err(presentation_match_conflict(format!(
            "Presentation scene element ID '{}' already belongs to a different record.",
            identity.id
        )));
    }
    if existing_claim.is_some_and(|claim| claim != fingerprint) {
        return Err(presentation_match_conflict(format!(
            "Presentation scene element ID '{}' was already claimed by a different record.",
            identity.id
        )));
    }
    let insertion_index = match after_element_id {
        Some(anchor_id) => {
            let anchor = state
                .records
                .get(anchor_id)
                .filter(|record| !record.tombstoned);
            if anchor.is_none() {
                return Err(presentation_match_conflict(format!(
                    "Presentation anchor scene element ID '{anchor_id}' does not exist."
                )));
            }
            state
                .raw_order
                .iter()
                .position(|id| id == anchor_id)
                .map(|index| index as u32 + 1)
                .ok_or_else(|| {
                    invalid_shared_presentation(
                        "The shared Presentation element order omits an active anchor.",
                    )
                })?
        }
        None => state.order.len(&doc.transact()),
    };
    let fields = element
        .as_object()
        .expect("validated scene element")
        .iter()
        .map(|(key, value)| Ok((key.clone(), json_to_any(value, 0)?)))
        .collect::<UseResult<Vec<_>>>()?;
    let claim = encoded_element_claim(container_kind, container_id, element)?;
    let append_claim = existing_claim.is_none();

    let mut transaction = doc.transact_mut();
    let record =
        state
            .elements
            .insert(&mut transaction, identity.id.as_str(), MapPrelim::default());
    for (key, value) in fields {
        record.insert(&mut transaction, key, value);
    }
    state
        .order
        .insert(&mut transaction, insertion_index, identity.id.as_str());
    if append_claim {
        state.claims.root.push_back(&mut transaction, claim);
    }
    drop(transaction);

    let next = read_element_state(doc, manifest, container_kind, container_id)?;
    if next
        .records
        .get(&identity.id)
        .is_none_or(|record| record.tombstoned || !json_equal(&record.value, element))
    {
        return Err(invalid_shared_presentation(
            "The Presentation scene-element creation did not produce the requested record.",
        ));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn update_element(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    container_kind: NativeOfficeCollaborationPresentationContainerKind,
    container_id: &str,
    element_id: &str,
    expected_element: &JsonValue,
    next_element: &JsonValue,
) -> UseResult<()> {
    let expected_identity = validate_element_input(expected_element, "expected scene element")?;
    let next_identity = validate_element_input(next_element, "next scene element")?;
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
    let shared_identity = validate_element_input(&record.value, "shared scene element")?;
    if expected_identity.id != element_id
        || next_identity.id != element_id
        || expected_identity.element_type != next_identity.element_type
        || expected_identity.element_type != shared_identity.element_type
    {
        return Err(presentation_match_conflict(format!(
            "Presentation scene element ID '{element_id}' no longer matches its expected immutable identity."
        )));
    }
    assert_compatible_element(
        expected_element,
        next_element,
        &record.value,
        &format!("Presentation scene element '{element_id}'"),
    )?;
    let patches = element_field_patches(expected_element, &record.value, next_element)?;
    if patches.is_empty() {
        return Ok(());
    }
    let record_map = record.map.clone();
    let mut transaction = doc.transact_mut();
    for patch in patches {
        match patch {
            PresentationElementPatch::Remove(key) => {
                record_map.remove(&mut transaction, key.as_str());
            }
            PresentationElementPatch::Set(key, value) => {
                record_map.insert(&mut transaction, key, value);
            }
        }
    }
    drop(transaction);
    read_element_state(doc, manifest, container_kind, container_id)?;
    Ok(())
}

fn delete_element(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    container_kind: NativeOfficeCollaborationPresentationContainerKind,
    container_id: &str,
    expected_element: &JsonValue,
) -> UseResult<()> {
    let expected_identity = validate_element_input(expected_element, "expected scene element")?;
    let state = read_element_state(doc, manifest, container_kind, container_id)?;
    let record = state.records.get(&expected_identity.id).ok_or_else(|| {
        presentation_match_conflict(format!(
            "Presentation scene element ID '{}' does not exist.",
            expected_identity.id
        ))
    })?;
    if !json_equal(&record.value, expected_element) {
        return Err(presentation_match_conflict(format!(
            "Presentation scene element ID '{}' changed before it could be deleted.",
            expected_identity.id
        )));
    }
    if record.tombstoned {
        return Ok(());
    }
    let record_map = record.map.clone();
    let mut transaction = doc.transact_mut();
    record_map.insert(&mut transaction, "tombstone", true);
    for index in (0..state.raw_order.len()).rev() {
        if state.raw_order[index] == expected_identity.id {
            state.order.remove_range(&mut transaction, index as u32, 1);
        }
    }
    drop(transaction);
    let next = read_element_state(doc, manifest, container_kind, container_id)?;
    if next
        .records
        .get(&expected_identity.id)
        .is_none_or(|record| !record.tombstoned)
    {
        return Err(invalid_shared_presentation(
            "The Presentation scene-element deletion did not leave a tombstone.",
        ));
    }
    Ok(())
}

fn validate_container_id(value: &str) -> UseResult<()> {
    validate_presentation_identifier(value, "container", false)
}
