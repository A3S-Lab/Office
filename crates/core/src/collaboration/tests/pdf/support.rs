use super::*;

pub(super) fn pdf_create_request(root: &std::path::Path) -> NativeOfficeCollaborationCreateRequest {
    NativeOfficeCollaborationCreateRequest {
        store: root.to_path_buf(),
        artifact_id: "fixture-pdf".to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Pdf,
        actor_id: "agent-alpha".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Agent,
        mode: NativeOfficeCollaborationMode::Edit,
        operation_id: "create-pdf-1".to_owned(),
        namespace: None,
        client_id: Some(900_003),
        initial_update: None,
    }
}

pub(super) fn pdf_apply_request(
    operation_id: &str,
    update: Vec<u8>,
) -> NativeOfficeCollaborationApplyRequest {
    NativeOfficeCollaborationApplyRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-alpha".to_owned(),
        mode: NativeOfficeCollaborationMode::Edit,
        expected_artifact_id: "fixture-pdf".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Pdf,
        update,
        if_state_vector: None,
        origin: None,
    }
}

pub(super) fn pdf_mutation_request(
    operation_id: &str,
    mutation: NativeOfficeCollaborationMutation,
) -> NativeOfficeCollaborationMutationRequest {
    NativeOfficeCollaborationMutationRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-alpha".to_owned(),
        mode: NativeOfficeCollaborationMode::Edit,
        expected_artifact_id: "fixture-pdf".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Pdf,
        mutation,
        if_state_vector: None,
    }
}

pub(super) fn pdf_form_value(
    store: &NativeOfficeCollaborationStore,
    field_id: &str,
) -> Option<String> {
    let peer = pdf_peer(store);
    let presence = peer.get_or_insert_map("a3s.office.pdf.form-values.presence");
    let fields = peer.get_or_insert_map("a3s.office.pdf.form-values.fields");
    let transaction = peer.transact();
    assert!(matches!(
        presence.get(&transaction, field_id),
        Some(Out::Any(Any::Bool(true)))
    ));
    let key = pdf_form_field_key(field_id, "value");
    match fields.get(&transaction, key.as_str()) {
        Some(Out::Any(Any::String(value))) => Some(value.to_string()),
        None => None,
        value => panic!("unexpected PDF form value: {value:?}"),
    }
}

pub(super) fn pdf_form_order(store: &NativeOfficeCollaborationStore) -> Vec<String> {
    let peer = pdf_peer(store);
    let order = peer.get_or_insert_array("a3s.office.pdf.form-values.order");
    let transaction = peer.transact();
    (0..order.len(&transaction))
        .map(|index| match order.get(&transaction, index) {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            value => panic!("unexpected PDF form order entry: {value:?}"),
        })
        .collect()
}

pub(super) fn pdf_peer(store: &NativeOfficeCollaborationStore) -> Doc {
    let exported = store.synchronize(None).unwrap();
    let peer = Doc::with_client_id(818_183);
    peer.transact_mut()
        .apply_update(Update::decode_v1(&exported.update).unwrap())
        .unwrap();
    peer
}

pub(super) fn pdf_form_field_key(field_id: &str, property: &str) -> String {
    let field = serde_json::to_string(&["value", property]).unwrap();
    serde_json::to_string(&[field_id, field.as_str()]).unwrap()
}

pub(super) fn portable_annotation(id: &str, color: &str, contents: &str) -> serde_json::Value {
    serde_json::json!({
        "id": id,
        "pageIndex": 1,
        "type": 9,
        "rect": {
            "origin": { "x": 68, "y": 78 },
            "size": { "width": 300, "height": 28 },
        },
        "segmentRects": [{
            "origin": { "x": 68, "y": 78 },
            "size": { "width": 300, "height": 28 },
        }],
        "strokeColor": color,
        "color": color,
        "opacity": 0.48,
        "contents": contents,
        "author": "A3S Agent",
        "created": "2026-08-15T08:00:00.000Z",
    })
}

pub(super) fn valid_pdf_rect() -> NativeOfficeCollaborationPdfRect {
    NativeOfficeCollaborationPdfRect {
        left: 10.0,
        top: 10.0,
        right: 20.0,
        bottom: 20.0,
    }
}

pub(super) fn pdf_record(
    store: &NativeOfficeCollaborationStore,
    collection: &str,
    record_id: &str,
) -> serde_json::Value {
    let peer = pdf_peer(store);
    let presence = peer.get_or_insert_map(format!("a3s.office.pdf.{collection}.presence"));
    let fields = peer.get_or_insert_map(format!("a3s.office.pdf.{collection}.fields"));
    let transaction = peer.transact();
    assert!(matches!(
        presence.get(&transaction, record_id),
        Some(Out::Any(Any::Bool(true)))
    ));
    let mut entries = Vec::new();
    for (encoded, value) in fields.iter(&transaction) {
        let identity = serde_json::from_str::<Vec<String>>(encoded).unwrap();
        if identity.first().map(String::as_str) != Some(record_id) {
            continue;
        }
        let field = serde_json::from_str::<Vec<String>>(&identity[1]).unwrap();
        assert!(field.len() >= 2);
        let kind = field[0].clone();
        let value = match (kind.as_str(), value) {
            ("object", Out::Any(Any::Bool(true))) => serde_json::Value::Bool(true),
            ("value", Out::Any(value)) => serde_json::to_value(value).unwrap(),
            _ => panic!("unexpected shared PDF record field"),
        };
        entries.push((kind, field[1..].to_vec(), value));
    }
    drop(transaction);
    entries.sort_by(|left, right| {
        left.1.len().cmp(&right.1.len()).then_with(|| {
            if left.0 == right.0 {
                std::cmp::Ordering::Equal
            } else if left.0 == "object" {
                std::cmp::Ordering::Less
            } else {
                std::cmp::Ordering::Greater
            }
        })
    });
    let object_paths = entries
        .iter()
        .filter(|(kind, _, _)| kind == "object")
        .map(|(_, path, _)| path.clone())
        .collect::<HashSet<_>>();
    for (_, path, _) in &entries {
        for depth in 1..path.len() {
            assert!(object_paths.contains(&path[..depth]));
        }
    }
    let mut record = serde_json::Map::new();
    for (kind, path, value) in entries {
        if kind == "object" {
            ensure_pdf_json_object_path(&mut record, &path);
        } else {
            set_pdf_json_value(&mut record, &path, value);
        }
    }
    serde_json::Value::Object(record)
}

pub(super) fn ensure_pdf_json_object_path(
    object: &mut serde_json::Map<String, serde_json::Value>,
    path: &[String],
) {
    let (key, rest) = path.split_first().expect("non-empty PDF field path");
    let child = object
        .entry(key.clone())
        .or_insert_with(|| serde_json::Value::Object(serde_json::Map::new()))
        .as_object_mut()
        .expect("PDF object marker must not overlap a value");
    if !rest.is_empty() {
        ensure_pdf_json_object_path(child, rest);
    }
}

pub(super) fn set_pdf_json_value(
    object: &mut serde_json::Map<String, serde_json::Value>,
    path: &[String],
    value: serde_json::Value,
) {
    let (key, rest) = path.split_first().expect("non-empty PDF field path");
    if rest.is_empty() {
        assert!(object.insert(key.clone(), value).is_none());
        return;
    }
    let child = object
        .get_mut(key)
        .and_then(serde_json::Value::as_object_mut)
        .expect("PDF value must have an object marker parent");
    set_pdf_json_value(child, rest, value);
}
