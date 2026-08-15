use serde_json::{Map as JsonMap, Value as JsonValue};
use yrs::updates::decoder::Decode;
use yrs::{Any, Array, Doc, Map, Out, Transact, Update};

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;

const PRESENTATION_COLLABORATION_FIXTURE_BASE64: &str = include_str!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../tests/fixtures/browser-presentation-collaboration-update.base64"
));

pub fn presentation_collaboration_fixture() -> Vec<u8> {
    STANDARD
        .decode(PRESENTATION_COLLABORATION_FIXTURE_BASE64.trim())
        .unwrap()
}

pub fn presentation_slide_title_element() -> JsonValue {
    serde_json::json!({
        "id": "element-title",
        "type": "text",
        "x": 10,
        "y": 10,
        "width": 80,
        "height": 20,
        "text": "Shared presentation",
        "fontSize": 32,
        "color": "#172033",
        "fill": "transparent",
        "bold": true,
        "align": "center",
    })
}

pub fn presentation_slide_body_element() -> JsonValue {
    serde_json::json!({
        "id": "element-body",
        "type": "shape",
        "x": 20,
        "y": 25,
        "width": 60,
        "height": 40,
        "text": "Body",
        "fontSize": 18,
        "color": "#172033",
        "fill": "#DCE6FB",
        "bold": false,
        "align": "left",
    })
}

pub fn presentation_scene_element(id: &str, text: &str, element_type: &str) -> JsonValue {
    serde_json::json!({
        "id": id,
        "type": element_type,
        "x": 24,
        "y": 24,
        "width": 40,
        "height": 20,
        "text": text,
        "fontSize": 18,
        "color": "#172033",
        "fill": "#F8FAFC",
        "bold": false,
        "align": "left",
    })
}

pub fn presentation_element(
    update: &[u8],
    collection: &str,
    container_id: &str,
    element_id: &str,
) -> Option<JsonValue> {
    let document = presentation_document(update);
    let element = element_map(&document, collection, container_id, element_id)?;
    let transaction = document.transact();
    if matches!(
        element.get(&transaction, "tombstone"),
        Some(Out::Any(Any::Bool(true)))
    ) {
        return None;
    }
    let mut object = JsonMap::new();
    for (key, value) in element.iter(&transaction) {
        let value = match value {
            Out::Any(value) => serde_json::to_value(value).unwrap(),
            value => panic!("unexpected Presentation scene-element field: {value:?}"),
        };
        object.insert(key.to_owned(), value);
    }
    Some(JsonValue::Object(object))
}

pub fn presentation_element_tombstoned(
    update: &[u8],
    collection: &str,
    container_id: &str,
    element_id: &str,
) -> bool {
    let document = presentation_document(update);
    let Some(element) = element_map(&document, collection, container_id, element_id) else {
        return false;
    };
    let tombstoned = matches!(
        element.get(&document.transact(), "tombstone"),
        Some(Out::Any(Any::Bool(true)))
    );
    tombstoned
}

pub fn presentation_element_order(
    update: &[u8],
    collection: &str,
    container_id: &str,
) -> Vec<String> {
    let document = presentation_document(update);
    let containers = document.get_or_insert_map(format!("a3s.office.presentation.{collection}"));
    let transaction = document.transact();
    let container = match containers.get(&transaction, container_id) {
        Some(Out::YMap(value)) => value,
        value => panic!("unexpected Presentation container: {value:?}"),
    };
    let order = match container.get(&transaction, "elementOrder") {
        Some(Out::YArray(value)) => value,
        value => panic!("unexpected Presentation element order: {value:?}"),
    };
    (0..order.len(&transaction))
        .map(|index| match order.get(&transaction, index) {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            value => panic!("unexpected Presentation order entry: {value:?}"),
        })
        .collect()
}

fn presentation_document(update: &[u8]) -> Doc {
    let document = Doc::with_client_id(710_015);
    document
        .transact_mut()
        .apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
    document
}

fn element_map(
    document: &Doc,
    collection: &str,
    container_id: &str,
    element_id: &str,
) -> Option<yrs::MapRef> {
    let containers = document.get_or_insert_map(format!("a3s.office.presentation.{collection}"));
    let transaction = document.transact();
    let container = match containers.get(&transaction, container_id) {
        Some(Out::YMap(value)) => value,
        None => return None,
        value => panic!("unexpected Presentation container: {value:?}"),
    };
    let elements = match container.get(&transaction, "elements") {
        Some(Out::YMap(value)) => value,
        value => panic!("unexpected Presentation elements root: {value:?}"),
    };
    match elements.get(&transaction, element_id) {
        Some(Out::YMap(value)) => Some(value),
        None => None,
        value => panic!("unexpected Presentation scene element: {value:?}"),
    }
}
