use std::collections::{HashMap, HashSet};

use a3s_use_core::UseResult;
use serde_json::{Map as JsonMap, Value as JsonValue};
use yrs::{Any, Array, ArrayRef, Map, MapRef, Out, Transact};

use super::json::{
    any_to_json, canonical_json, validate_presentation_identifier, validate_shared_element,
};
use super::{
    invalid_shared_presentation, NativeOfficeCollaborationManifest,
    NativeOfficeCollaborationPresentationContainerKind,
};

const MAX_PRESENTATION_ELEMENTS: u32 = 1_000_000;

#[derive(Debug)]
pub(super) struct PresentationElementRecord {
    pub(super) map: MapRef,
    pub(super) value: JsonValue,
    pub(super) tombstoned: bool,
}

pub(super) struct PresentationClaims {
    pub(super) root: ArrayRef,
    by_identity: HashMap<(String, String, String), String>,
}

impl PresentationClaims {
    pub(super) fn claim_for(
        &self,
        container_kind: NativeOfficeCollaborationPresentationContainerKind,
        container_id: &str,
        element_id: &str,
    ) -> Option<&str> {
        self.by_identity
            .get(&(
                container_kind.as_str().to_owned(),
                container_id.to_owned(),
                element_id.to_owned(),
            ))
            .map(String::as_str)
    }
}

pub(super) struct PresentationElementState {
    pub(super) claims: PresentationClaims,
    pub(super) elements: MapRef,
    pub(super) order: ArrayRef,
    pub(super) raw_order: Vec<String>,
    pub(super) records: HashMap<String, PresentationElementRecord>,
}

pub(super) fn read_element_state(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    container_kind: NativeOfficeCollaborationPresentationContainerKind,
    container_id: &str,
) -> UseResult<PresentationElementState> {
    let claims = read_presentation_claims(doc, manifest)?;
    let containers = doc.get_or_insert_map(format!(
        "{}.presentation.{}",
        manifest.namespace,
        container_collection(container_kind)
    ));
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
            return Err(super::presentation_match_conflict(format!(
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
    let elements = match container.get(&transaction, "elements") {
        Some(Out::YMap(value)) => value,
        _ => {
            return Err(invalid_shared_presentation(format!(
                "The shared Presentation {} '{container_id}' elements root is invalid.",
                container_kind.as_str()
            )))
        }
    };
    let order = match container.get(&transaction, "elementOrder") {
        Some(Out::YArray(value)) => value,
        _ => {
            return Err(invalid_shared_presentation(format!(
                "The shared Presentation {} '{container_id}' element order is invalid.",
                container_kind.as_str()
            )))
        }
    };
    if elements.len(&transaction) > MAX_PRESENTATION_ELEMENTS {
        return Err(invalid_shared_presentation(format!(
            "The shared Presentation {} '{container_id}' contains too many scene elements.",
            container_kind.as_str()
        )));
    }

    let mut raw_order = Vec::new();
    let mut ordered_ids = HashSet::new();
    for index in 0..order.len(&transaction) {
        let id = match order.get(&transaction, index) {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            _ => {
                return Err(invalid_shared_presentation(
                    "The shared Presentation element order contains a non-string identity.",
                ))
            }
        };
        validate_presentation_identifier(&id, "scene element", true)?;
        ordered_ids.insert(id.clone());
        raw_order.push(id);
    }

    let mut records = HashMap::new();
    for (id, output) in elements.iter(&transaction) {
        validate_presentation_identifier(id, "scene element", true)?;
        let map = match output {
            Out::YMap(value) => value,
            _ => {
                return Err(invalid_shared_presentation(
                    "A shared Presentation scene element is not a typed map.",
                ))
            }
        };
        let record = read_element_record(&map, &transaction, id)?;
        if record.tombstoned == ordered_ids.contains(id) {
            return Err(invalid_shared_presentation(
                "The shared Presentation element order and record set disagree.",
            ));
        }
        assert_claimed_identity(&claims, container_kind, container_id, id, &record.value)?;
        records.insert(id.to_owned(), record);
    }
    if ordered_ids.iter().any(|id| !records.contains_key(id)) {
        return Err(invalid_shared_presentation(
            "The shared Presentation element order and record set disagree.",
        ));
    }
    drop(transaction);
    Ok(PresentationElementState {
        claims,
        elements,
        order,
        raw_order,
        records,
    })
}

pub(super) fn encoded_element_claim(
    container_kind: NativeOfficeCollaborationPresentationContainerKind,
    container_id: &str,
    element: &JsonValue,
) -> UseResult<String> {
    let element_id = element
        .get("id")
        .and_then(JsonValue::as_str)
        .expect("validated scene element ID");
    canonical_json(&serde_json::json!({
        "containerId": container_id,
        "containerKind": container_kind.as_str(),
        "fingerprint": canonical_json(element)?,
        "id": element_id,
        "kind": "element",
    }))
}

fn read_element_record<T: yrs::ReadTxn>(
    map: &MapRef,
    transaction: &T,
    expected_id: &str,
) -> UseResult<PresentationElementRecord> {
    let mut object = JsonMap::new();
    let mut tombstoned = false;
    for (key, output) in map.iter(transaction) {
        if key == "tombstone" {
            if !matches!(output, Out::Any(Any::Bool(true))) {
                return Err(invalid_shared_presentation(
                    "A shared Presentation scene-element tombstone is invalid.",
                ));
            }
            tombstoned = true;
            continue;
        }
        let value = match output {
            Out::Any(value) => any_to_json(value, 0, "scene element")?,
            _ => {
                return Err(invalid_shared_presentation(format!(
                    "The shared Presentation scene element field '{key}' contains a shared type."
                )))
            }
        };
        object.insert(key.to_owned(), value);
    }
    let value = JsonValue::Object(object);
    let identity = validate_shared_element(&value, "scene element")?;
    if identity.id != expected_id {
        return Err(invalid_shared_presentation(
            "A shared Presentation scene-element identity does not match its collection key.",
        ));
    }
    Ok(PresentationElementRecord {
        map: map.clone(),
        value,
        tombstoned,
    })
}

fn read_presentation_claims(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<PresentationClaims> {
    let root =
        doc.get_or_insert_array(format!("{}.presentation.record-claims", manifest.namespace));
    let transaction = doc.transact();
    let mut by_identity = HashMap::new();
    for index in 0..root.len(&transaction) {
        let raw = match root.get(&transaction, index) {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            _ => {
                return Err(invalid_shared_presentation(
                    "A shared Presentation record claim is not a string.",
                ))
            }
        };
        let value: JsonValue = serde_json::from_str(&raw).map_err(|_| {
            invalid_shared_presentation("A shared Presentation record claim is not valid JSON.")
        })?;
        if canonical_json(&value)? != raw {
            return Err(invalid_shared_presentation(
                "A shared Presentation record claim is not canonical JSON.",
            ));
        }
        let object = value.as_object().ok_or_else(|| {
            invalid_shared_presentation("A shared Presentation record claim is not an object.")
        })?;
        let expected_keys = ["containerId", "containerKind", "fingerprint", "id", "kind"];
        if object.len() != expected_keys.len()
            || object
                .keys()
                .any(|key| !expected_keys.contains(&key.as_str()))
        {
            return Err(invalid_shared_presentation(
                "A shared Presentation record claim contains unsupported fields.",
            ));
        }
        if object.get("kind").and_then(JsonValue::as_str) != Some("element") {
            return Err(invalid_shared_presentation(
                "A shared Presentation record claim kind is unsupported.",
            ));
        }
        let container_kind =
            parsed_container_kind(object.get("containerKind").and_then(JsonValue::as_str))?;
        let container_id = required_claim_string(object, "containerId")?;
        let id = required_claim_string(object, "id")?;
        validate_presentation_identifier(container_id, "record claim container", true)?;
        validate_presentation_identifier(id, "record claim", true)?;
        let fingerprint = required_claim_string(object, "fingerprint")?;
        let claimed: JsonValue = serde_json::from_str(fingerprint).map_err(|_| {
            invalid_shared_presentation(
                "A shared Presentation record claim fingerprint is not valid JSON.",
            )
        })?;
        if canonical_json(&claimed)? != fingerprint {
            return Err(invalid_shared_presentation(
                "A shared Presentation record claim fingerprint is not canonical JSON.",
            ));
        }
        let claimed_identity = validate_shared_element(&claimed, "record claim fingerprint")?;
        if claimed_identity.id != id {
            return Err(invalid_shared_presentation(
                "A shared Presentation record claim fingerprint has the wrong identity.",
            ));
        }
        let identity = (
            container_kind.as_str().to_owned(),
            container_id.to_owned(),
            id.to_owned(),
        );
        if let Some(existing) = by_identity.insert(identity, fingerprint.to_owned()) {
            if existing != fingerprint {
                return Err(invalid_shared_presentation(format!(
                    "The scene element in {} '{container_id}' ID '{id}' was concurrently assigned to different records.",
                    container_kind.as_str()
                )));
            }
        }
    }
    drop(transaction);
    Ok(PresentationClaims { root, by_identity })
}

fn assert_claimed_identity(
    claims: &PresentationClaims,
    container_kind: NativeOfficeCollaborationPresentationContainerKind,
    container_id: &str,
    element_id: &str,
    element: &JsonValue,
) -> UseResult<()> {
    let Some(fingerprint) = claims.claim_for(container_kind, container_id, element_id) else {
        // The claim root was added after the original Presentation model.
        return Ok(());
    };
    let claimed: JsonValue = serde_json::from_str(fingerprint).map_err(|_| {
        invalid_shared_presentation(
            "A shared Presentation record claim fingerprint is not valid JSON.",
        )
    })?;
    let claimed_identity = validate_shared_element(&claimed, "record claim fingerprint")?;
    let element_identity = validate_shared_element(element, "scene element")?;
    if claimed_identity.id != element_identity.id
        || claimed_identity.element_type != element_identity.element_type
    {
        return Err(invalid_shared_presentation(
            "A shared Presentation scene element violates its immutable identity claim.",
        ));
    }
    Ok(())
}

fn required_claim_string<'a>(
    object: &'a JsonMap<String, JsonValue>,
    key: &str,
) -> UseResult<&'a str> {
    object.get(key).and_then(JsonValue::as_str).ok_or_else(|| {
        invalid_shared_presentation(format!(
            "A shared Presentation record claim '{key}' field is not a string."
        ))
    })
}

fn parsed_container_kind(
    value: Option<&str>,
) -> UseResult<NativeOfficeCollaborationPresentationContainerKind> {
    match value {
        Some("slide") => Ok(NativeOfficeCollaborationPresentationContainerKind::Slide),
        Some("master") => Ok(NativeOfficeCollaborationPresentationContainerKind::Master),
        Some("layout") => Ok(NativeOfficeCollaborationPresentationContainerKind::Layout),
        _ => Err(invalid_shared_presentation(
            "A shared Presentation record claim container kind is unsupported.",
        )),
    }
}

fn container_collection(value: NativeOfficeCollaborationPresentationContainerKind) -> &'static str {
    match value {
        NativeOfficeCollaborationPresentationContainerKind::Slide => "slides",
        NativeOfficeCollaborationPresentationContainerKind::Master => "masters",
        NativeOfficeCollaborationPresentationContainerKind::Layout => "layouts",
    }
}
