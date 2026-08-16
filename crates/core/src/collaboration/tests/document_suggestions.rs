use std::collections::HashMap;
use std::sync::Arc;

use base64::Engine as _;
use yrs::types::Attrs;
use yrs::{Any, Doc, Map, Text, Transact, Update, XmlFragment, XmlOut};

use super::*;

#[test]
fn authenticated_transport_accepts_attributed_suggestions_and_rejects_canonical_edits() {
    let temp = tempfile::tempdir().unwrap();
    let server = NativeOfficeCollaborationStore::create(NativeOfficeCollaborationCreateRequest {
        store: temp.path().join("authorized-suggestion-server"),
        artifact_id: "fixture-document".to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Document,
        actor_id: "collaboration-server".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::System,
        mode: NativeOfficeCollaborationMode::Edit,
        operation_id: "create-authorized-suggestion-server".to_owned(),
        namespace: None,
        client_id: Some(900_301),
        initial_update: None,
    })
    .unwrap();
    server
        .apply(NativeOfficeCollaborationApplyRequest {
            operation_id: "bootstrap-authorized-suggestion-server".to_owned(),
            actor_id: "collaboration-server".to_owned(),
            mode: NativeOfficeCollaborationMode::Edit,
            expected_artifact_id: "fixture-document".to_owned(),
            expected_kind: NativeOfficeCollaborationArtifactKind::Document,
            update: STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap(),
            if_state_vector: None,
            origin: None,
        })
        .unwrap();

    let authorization = NativeOfficeCollaborationTransportAuthorization {
        actor_id: "suggester-1".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Human,
        actor_name: "Ada Suggester".to_owned(),
        mode: NativeOfficeCollaborationMode::Suggest,
    };
    let mut transport = NativeOfficeCollaborationTransportSession::attach(server.clone()).unwrap();
    let suggestion = suggestion_update(&server, suggestion_attributes("Ada Suggester"), true);
    let operation_id = "authorized-suggestion-delivery";
    let accepted = transport
        .receive_authorized(
            NativeOfficeCollaborationTransportReceiveRequest {
                message: suggestion_message(suggestion, operation_id),
                operation_id: Some(operation_id.to_owned()),
                if_state_vector: None,
            },
            authorization.clone(),
        )
        .unwrap();
    assert!(accepted.apply.unwrap().state_changed);
    let NativeOfficeCollaborationProjectedContent::Document { plain_text, .. } =
        server.project().unwrap().content
    else {
        panic!("expected Document projection");
    };
    assert!(plain_text.ends_with(" proposed"));
    let durable = server
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(1),
            limit: 1,
        })
        .unwrap();
    assert_eq!(durable.updates[0].actor_id, "suggester-1");
    assert_eq!(
        durable.updates[0].mode,
        NativeOfficeCollaborationMode::Suggest
    );

    let before_rejection = server.inspect().unwrap().document_state_sha256;
    let forged = suggestion_update(&server, Attrs::new(), true);
    let forged_operation = "forged-canonical-suggestion-delivery";
    let rejected = transport
        .receive_authorized(
            NativeOfficeCollaborationTransportReceiveRequest {
                message: suggestion_message(forged, forged_operation),
                operation_id: Some(forged_operation.to_owned()),
                if_state_vector: None,
            },
            authorization.clone(),
        )
        .unwrap_err();
    assert_eq!(rejected.code, "office.collaboration.permission_denied");
    assert_eq!(
        server.inspect().unwrap().document_state_sha256,
        before_rejection
    );

    let forged_author =
        suggestion_update(&server, suggestion_attributes("Forged Display Name"), true);
    let forged_author_operation = "forged-suggestion-author-delivery";
    let rejected = transport
        .receive_authorized(
            NativeOfficeCollaborationTransportReceiveRequest {
                message: suggestion_message(forged_author, forged_author_operation),
                operation_id: Some(forged_author_operation.to_owned()),
                if_state_vector: None,
            },
            authorization,
        )
        .unwrap_err();
    assert_eq!(rejected.code, "office.collaboration.permission_denied");
    assert_eq!(
        server.inspect().unwrap().document_state_sha256,
        before_rejection
    );
}

#[test]
fn suggestion_transport_protects_foreign_proposals_and_allows_owner_withdrawal() {
    let temp = tempfile::tempdir().unwrap();
    let server = suggestion_server(temp.path().join("suggestion-ownership-server"));
    let ada = suggestion_authorization("suggester-1", "Ada Suggester");
    let mut transport = NativeOfficeCollaborationTransportSession::attach(server.clone()).unwrap();
    transport
        .receive_authorized(
            NativeOfficeCollaborationTransportReceiveRequest {
                message: suggestion_message(
                    suggestion_update(&server, suggestion_attributes("Ada Suggester"), true),
                    "create-owned-suggestion",
                ),
                operation_id: Some("create-owned-suggestion".to_owned()),
                if_state_vector: None,
            },
            ada.clone(),
        )
        .unwrap();

    let before_rejection = server.inspect().unwrap().document_state_sha256;
    let bob = suggestion_authorization("suggester-2", "Bob Reviewer");
    let rejected = transport
        .receive_authorized(
            NativeOfficeCollaborationTransportReceiveRequest {
                message: suggestion_message_for_actor(
                    remove_trailing_suggestion_update(&server),
                    "withdraw-foreign-suggestion",
                    "suggester-2",
                ),
                operation_id: Some("withdraw-foreign-suggestion".to_owned()),
                if_state_vector: None,
            },
            bob,
        )
        .unwrap_err();
    assert_eq!(rejected.code, "office.collaboration.permission_denied");
    assert_eq!(
        server.inspect().unwrap().document_state_sha256,
        before_rejection
    );

    transport
        .receive_authorized(
            NativeOfficeCollaborationTransportReceiveRequest {
                message: suggestion_message(
                    remove_trailing_suggestion_update(&server),
                    "withdraw-owned-suggestion",
                ),
                operation_id: Some("withdraw-owned-suggestion".to_owned()),
                if_state_vector: None,
            },
            ada,
        )
        .unwrap();
    let NativeOfficeCollaborationProjectedContent::Document { plain_text, .. } =
        server.project().unwrap().content
    else {
        panic!("expected Document projection");
    };
    assert!(!plain_text.ends_with(" proposed"));
}

#[test]
fn suggestion_transport_accepts_deletion_marks_but_rejects_sidecar_changes() {
    let temp = tempfile::tempdir().unwrap();
    let server = suggestion_server(temp.path().join("suggestion-deletion-server"));
    let authorization = suggestion_authorization("suggester-1", "Ada Suggester");
    let initial_text = match server.project().unwrap().content {
        NativeOfficeCollaborationProjectedContent::Document { plain_text, .. } => plain_text,
        _ => panic!("expected Document projection"),
    };
    let mut transport = NativeOfficeCollaborationTransportSession::attach(server.clone()).unwrap();
    transport
        .receive_authorized(
            NativeOfficeCollaborationTransportReceiveRequest {
                message: suggestion_message(
                    deletion_suggestion_update(&server),
                    "create-deletion-suggestion",
                ),
                operation_id: Some("create-deletion-suggestion".to_owned()),
                if_state_vector: None,
            },
            authorization.clone(),
        )
        .unwrap();
    let NativeOfficeCollaborationProjectedContent::Document { plain_text, .. } =
        server.project().unwrap().content
    else {
        panic!("expected Document projection");
    };
    assert_eq!(plain_text, initial_text);

    let before_rejection = server.inspect().unwrap().document_state_sha256;
    let rejected = transport
        .receive_authorized(
            NativeOfficeCollaborationTransportReceiveRequest {
                message: suggestion_message(
                    document_option_update(&server),
                    "forged-suggestion-sidecar",
                ),
                operation_id: Some("forged-suggestion-sidecar".to_owned()),
                if_state_vector: None,
            },
            authorization,
        )
        .unwrap_err();
    assert_eq!(rejected.code, "office.collaboration.permission_denied");
    assert_eq!(
        server.inspect().unwrap().document_state_sha256,
        before_rejection
    );
}

fn suggestion_server(store: std::path::PathBuf) -> NativeOfficeCollaborationStore {
    let server = NativeOfficeCollaborationStore::create(NativeOfficeCollaborationCreateRequest {
        store,
        artifact_id: "fixture-document".to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Document,
        actor_id: "collaboration-server".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::System,
        mode: NativeOfficeCollaborationMode::Edit,
        operation_id: "create-suggestion-server".to_owned(),
        namespace: None,
        client_id: Some(900_302),
        initial_update: None,
    })
    .unwrap();
    server
        .apply(NativeOfficeCollaborationApplyRequest {
            operation_id: "bootstrap-suggestion-server".to_owned(),
            actor_id: "collaboration-server".to_owned(),
            mode: NativeOfficeCollaborationMode::Edit,
            expected_artifact_id: "fixture-document".to_owned(),
            expected_kind: NativeOfficeCollaborationArtifactKind::Document,
            update: STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap(),
            if_state_vector: None,
            origin: None,
        })
        .unwrap();
    server
}

fn suggestion_authorization(
    actor_id: &str,
    actor_name: &str,
) -> NativeOfficeCollaborationTransportAuthorization {
    NativeOfficeCollaborationTransportAuthorization {
        actor_id: actor_id.to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Human,
        actor_name: actor_name.to_owned(),
        mode: NativeOfficeCollaborationMode::Suggest,
    }
}

fn suggestion_update(
    server: &NativeOfficeCollaborationStore,
    attributes: Attrs,
    insert: bool,
) -> Vec<u8> {
    let producer = Doc::with_client_id(818_585);
    producer
        .transact_mut()
        .apply_update(Update::decode_v1(&server.synchronize(None).unwrap().update).unwrap())
        .unwrap();
    let before = producer.transact().state_vector();
    let fragment = producer.get_or_insert_xml_fragment("a3s.office.document.content");
    let text = fragment
        .successors(&producer.transact())
        .find_map(|node| match node {
            XmlOut::Text(text) => Some(text),
            _ => None,
        })
        .unwrap();
    let mut transaction = producer.transact_mut();
    let position = text.len(&transaction);
    if insert && attributes.is_empty() {
        text.insert_with_attributes(&mut transaction, position, " FORGED", Attrs::new());
    } else {
        text.insert_with_attributes(&mut transaction, position, " proposed", attributes);
    }
    drop(transaction);
    let update = producer.transact().encode_state_as_update_v1(&before);
    update
}

fn remove_trailing_suggestion_update(server: &NativeOfficeCollaborationStore) -> Vec<u8> {
    let producer = synchronized_producer(server, 818_586);
    let before = producer.transact().state_vector();
    let text = first_document_text(&producer);
    let mut transaction = producer.transact_mut();
    let start = text.len(&transaction) - u32::try_from(" proposed".encode_utf16().count()).unwrap();
    text.remove_range(&mut transaction, start, 9);
    drop(transaction);
    let update = producer.transact().encode_state_as_update_v1(&before);
    update
}

fn deletion_suggestion_update(server: &NativeOfficeCollaborationStore) -> Vec<u8> {
    let producer = synchronized_producer(server, 818_587);
    let before = producer.transact().state_vector();
    let text = first_document_text(&producer);
    let mut attributes = suggestion_attributes("Ada Suggester");
    let value = attributes.remove("documentChange").unwrap();
    let Any::Map(fields) = value else {
        unreachable!("suggestion attributes are a map")
    };
    let mut fields = (*fields).clone();
    fields.insert("id".to_owned(), Any::String("suggestion-delete-1".into()));
    fields.insert("kind".to_owned(), Any::String("deletion".into()));
    attributes.insert("documentChange".into(), Any::Map(Arc::new(fields)));
    text.format(&mut producer.transact_mut(), 0, 1, attributes);
    let update = producer.transact().encode_state_as_update_v1(&before);
    update
}

fn document_option_update(server: &NativeOfficeCollaborationStore) -> Vec<u8> {
    let producer = synchronized_producer(server, 818_588);
    let before = producer.transact().state_vector();
    producer
        .get_or_insert_map("a3s.office.document.options")
        .insert(&mut producer.transact_mut(), "pageColor", "#ff0000");
    let update = producer.transact().encode_state_as_update_v1(&before);
    update
}

fn synchronized_producer(server: &NativeOfficeCollaborationStore, client_id: u64) -> Doc {
    let producer = Doc::with_client_id(client_id);
    producer
        .transact_mut()
        .apply_update(Update::decode_v1(&server.synchronize(None).unwrap().update).unwrap())
        .unwrap();
    producer
}

fn first_document_text(producer: &Doc) -> yrs::XmlTextRef {
    producer
        .get_or_insert_xml_fragment("a3s.office.document.content")
        .successors(&producer.transact())
        .find_map(|node| match node {
            XmlOut::Text(text) => Some(text),
            _ => None,
        })
        .unwrap()
}

fn suggestion_attributes(author: &str) -> Attrs {
    let fields = HashMap::from([
        ("actorId".to_owned(), Any::String("suggester-1".into())),
        ("author".to_owned(), Any::String(author.into())),
        (
            "date".to_owned(),
            Any::String("2026-08-17T08:00:00.000Z".into()),
        ),
        ("id".to_owned(), Any::String("suggestion-native-1".into())),
        ("kind".to_owned(), Any::String("insertion".into())),
    ]);
    Attrs::from([("documentChange".into(), Any::Map(Arc::new(fields)))])
}

fn suggestion_message(
    update: Vec<u8>,
    operation_id: &str,
) -> NativeOfficeCollaborationTransportMessage {
    suggestion_message_for_actor(update, operation_id, "suggester-1")
}

fn suggestion_message_for_actor(
    update: Vec<u8>,
    operation_id: &str,
    actor_id: &str,
) -> NativeOfficeCollaborationTransportMessage {
    NativeOfficeCollaborationTransportMessage {
        protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
        version: NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
        artifact_id: "fixture-document".to_owned(),
        artifact_kind: NativeOfficeCollaborationArtifactKind::Document,
        namespace: NATIVE_OFFICE_COLLABORATION_NAMESPACE.to_owned(),
        sender_client_id: 818_585,
        message_type: NativeOfficeCollaborationTransportMessageType::Update,
        payload: update,
        origin: Some(NativeOfficeCollaborationOrigin {
            protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
            kind: NativeOfficeCollaborationOriginKind::Editor,
            actor_id: Some(actor_id.to_owned()),
            operation_id: Some(operation_id.to_owned()),
        }),
    }
}
