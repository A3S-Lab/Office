use base64::Engine as _;
use yrs::types::Attrs;
use yrs::types::ToJson;
use yrs::{Any, Array, Doc, Map, Out, Text, Transact, Update, XmlFragment, XmlOut};

use super::*;

const COMMENT_ID: &str = "comment-native-1";
const REPLY_ID: &str = "reply-native-1";
const CREATED_AT: &str = "2026-08-17T00:00:00.000Z";

#[test]
fn native_document_comments_are_browser_readable_durable_and_projected() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("document-comment-replica");
    let store = NativeOfficeCollaborationStore::create(comment_create_request(&root)).unwrap();
    store
        .apply(comment_apply_request(
            "comment-bootstrap",
            STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();

    let create = comment_mutation_request(
        "comment-create-1",
        NativeOfficeCollaborationMutation::DocumentCommentCreate {
            comment_id: COMMENT_ID.to_owned(),
            paragraph_id: "00000001".to_owned(),
            expected_text_id: "00000002".to_owned(),
            start_utf16: 6,
            end_utf16: 8,
            expected_text: "😀".to_owned(),
            author: "Ada".to_owned(),
            created_at: CREATED_AT.to_owned(),
            text: "Keep the browser-compatible anchor.".to_owned(),
        },
    );
    let created = store.mutate(create.clone()).unwrap();
    assert!(created.state_changed);
    assert_eq!(created.sequence, Some(2));

    let replay = store.mutate(create).unwrap();
    assert!(replay.duplicate);
    assert_eq!(replay.sequence, Some(2));

    assert_browser_comment_state(&store, true, false, false);
    let projection = store.project().unwrap();
    let NativeOfficeCollaborationProjectedContent::Document { comments, .. } = projection.content
    else {
        panic!("expected Document projection");
    };
    assert_eq!(comments.len(), 1);
    let comment = &comments[0];
    assert_eq!(comment.id, COMMENT_ID);
    assert_eq!(comment.actor_id.as_deref(), Some("agent-alpha"));
    assert_eq!(comment.author, "Ada");
    assert!(!comment.resolved);
    assert!(!comment.detached);
    assert!(comment.replies.is_empty());
    assert_eq!(comment.anchors.len(), 1);
    assert_eq!(comment.anchors[0].paragraph_id.as_deref(), Some("00000001"));
    assert_eq!(comment.anchors[0].text_id.as_deref(), Some("00000002"));
    assert_eq!(comment.anchors[0].start_utf16, 6);
    assert_eq!(comment.anchors[0].end_utf16, 8);
    assert_eq!(comment.anchors[0].text, "😀");

    store
        .mutate(comment_mutation_request(
            "comment-reply-1",
            NativeOfficeCollaborationMutation::DocumentCommentReply {
                comment_id: COMMENT_ID.to_owned(),
                reply_id: REPLY_ID.to_owned(),
                author: "Ada".to_owned(),
                created_at: "2026-08-17T00:01:00.000Z".to_owned(),
                text: "Confirmed.".to_owned(),
            },
        ))
        .unwrap();
    store
        .mutate(comment_mutation_request(
            "comment-resolve-1",
            NativeOfficeCollaborationMutation::DocumentCommentSetResolved {
                comment_id: COMMENT_ID.to_owned(),
                resolved: true,
            },
        ))
        .unwrap();
    assert_browser_comment_state(&store, true, true, true);

    drop(store);
    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    let NativeOfficeCollaborationProjectedContent::Document { comments, .. } =
        reopened.project().unwrap().content
    else {
        panic!("expected Document projection");
    };
    assert!(comments[0].resolved);
    assert_eq!(comments[0].replies[0].id, REPLY_ID);

    // A compacted/full-state update has a different internal Yrs split
    // history. Deleting its interior comment mark must remain panic-free.
    let full_state_root = temp.path().join("document-comment-full-state");
    let full_state = reopened.synchronize(None).unwrap().update;
    let full_state_store =
        NativeOfficeCollaborationStore::create(NativeOfficeCollaborationCreateRequest {
            store: full_state_root,
            artifact_id: "fixture-document".to_owned(),
            kind: NativeOfficeCollaborationArtifactKind::Document,
            actor_id: "agent-alpha".to_owned(),
            actor_kind: NativeOfficeCollaborationActorKind::Agent,
            mode: NativeOfficeCollaborationMode::Comment,
            operation_id: "join-comment-full-state".to_owned(),
            namespace: None,
            client_id: Some(900_103),
            initial_update: Some(full_state),
        })
        .unwrap();
    full_state_store
        .mutate(comment_mutation_request(
            "comment-delete-full-state",
            NativeOfficeCollaborationMutation::DocumentCommentDelete {
                comment_id: COMMENT_ID.to_owned(),
                reply_id: None,
            },
        ))
        .unwrap();
    let NativeOfficeCollaborationProjectedContent::Document { comments, .. } =
        full_state_store.project().unwrap().content
    else {
        panic!("expected Document projection");
    };
    assert!(comments.is_empty());

    reopened
        .mutate(comment_mutation_request(
            "comment-delete-reply-1",
            NativeOfficeCollaborationMutation::DocumentCommentDelete {
                comment_id: COMMENT_ID.to_owned(),
                reply_id: Some(REPLY_ID.to_owned()),
            },
        ))
        .unwrap();
    assert_browser_comment_state(&reopened, true, true, false);
    reopened
        .mutate(comment_mutation_request(
            "comment-delete-thread-1",
            NativeOfficeCollaborationMutation::DocumentCommentDelete {
                comment_id: COMMENT_ID.to_owned(),
                reply_id: None,
            },
        ))
        .unwrap();
    assert_browser_comment_state(&reopened, false, false, false);
}

#[test]
fn document_comment_mutations_fail_closed_on_conflicts_and_modes() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("document-comment-conflicts");
    let store = NativeOfficeCollaborationStore::create(comment_create_request(&root)).unwrap();
    store
        .apply(comment_apply_request(
            "comment-conflict-bootstrap",
            STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    store
        .mutate(comment_mutation_request(
            "comment-conflict-create",
            create_comment_mutation("Original"),
        ))
        .unwrap();
    let before = store.inspect().unwrap();

    let collision = store
        .mutate(comment_mutation_request(
            "comment-conflict-collision",
            create_comment_mutation("Different stable-ID payload"),
        ))
        .unwrap_err();
    assert_eq!(
        collision.code,
        "office.collaboration.mutation_identity_conflict"
    );
    let stale_anchor = store
        .mutate(comment_mutation_request(
            "comment-conflict-anchor",
            create_comment_mutation_at("Original", 0, 5, "Hello"),
        ))
        .unwrap_err();
    assert_eq!(
        stale_anchor.code,
        "office.collaboration.mutation_identity_conflict"
    );
    let content_write = store
        .mutate(comment_mutation_request(
            "comment-mode-content-write",
            NativeOfficeCollaborationMutation::DocumentReplaceText {
                search: "Hello".to_owned(),
                replacement: "Changed".to_owned(),
                expected_matches: 1,
            },
        ))
        .unwrap_err();
    assert_eq!(
        content_write.code,
        "office.collaboration.mutation_forbidden"
    );
    let after = store.inspect().unwrap();
    assert_eq!(before.current_sequence, after.current_sequence);
    assert_eq!(before.document_state_sha256, after.document_state_sha256);

    for mode in [
        NativeOfficeCollaborationMode::View,
        NativeOfficeCollaborationMode::Suggest,
    ] {
        let mode_root = temp.path().join(mode.as_str());
        let mode_store =
            NativeOfficeCollaborationStore::create(mode_create_request(&mode_root, mode)).unwrap();
        mode_store
            .apply(mode_apply_request(
                "mode-bootstrap",
                mode,
                STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap(),
            ))
            .unwrap();
        let denied = mode_store
            .mutate(mode_mutation_request(
                "mode-comment-create",
                mode,
                create_comment_mutation("Denied"),
            ))
            .unwrap_err();
        assert_eq!(denied.code, "office.collaboration.mutation_forbidden");
    }
}

#[test]
fn detached_foreign_comments_remain_reviewable_but_not_deletable() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("document-comment-detached");
    let store = NativeOfficeCollaborationStore::create(comment_create_request(&root)).unwrap();
    store
        .apply(comment_apply_request(
            "comment-detached-bootstrap",
            STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    store
        .mutate(comment_mutation_request(
            "comment-detached-create",
            create_comment_mutation("Detached review"),
        ))
        .unwrap();

    let browser = Doc::with_client_id(818_383);
    browser
        .transact_mut()
        .apply_update(Update::decode_v1(&store.synchronize(None).unwrap().update).unwrap())
        .unwrap();
    let before_anchor_delete = browser.transact().state_vector();
    let fragment = browser.get_or_insert_xml_fragment("a3s.office.document.content");
    let (text, attribute) = {
        let transaction = browser.transact();
        fragment
            .successors(&transaction)
            .find_map(|node| match node {
                XmlOut::Text(text) => text.diff(&transaction, |_| ()).iter().find_map(|chunk| {
                    chunk.attributes.as_ref().and_then(|attributes| {
                        attributes
                            .keys()
                            .find(|key| key.starts_with("documentComment"))
                            .map(|key| (text.clone(), key.to_string()))
                    })
                }),
                _ => None,
            })
            .unwrap()
    };
    let text_length = text.len(&browser.transact());
    text.format(
        &mut browser.transact_mut(),
        0,
        text_length,
        Attrs::from([(attribute.into(), Any::Null)]),
    );
    let anchor_delete = browser
        .transact()
        .encode_state_as_update_v1(&before_anchor_delete);
    store
        .apply(comment_apply_request(
            "browser-delete-comment-anchor",
            anchor_delete,
        ))
        .unwrap();

    let NativeOfficeCollaborationProjectedContent::Document { comments, .. } =
        store.project().unwrap().content
    else {
        panic!("expected Document projection");
    };
    assert!(comments[0].detached);
    assert!(comments[0].anchors.is_empty());

    let before_actor_change = browser.transact().state_vector();
    let comments = browser.get_or_insert_map("a3s.office.document.comments");
    let comment = match comments.get(&browser.transact(), COMMENT_ID) {
        Some(Out::YMap(comment)) => comment,
        value => panic!("unexpected comment record: {value:?}"),
    };
    comment.insert(&mut browser.transact_mut(), "actorId", "foreign-reviewer");
    let actor_change = browser
        .transact()
        .encode_state_as_update_v1(&before_actor_change);
    store
        .apply(comment_apply_request(
            "browser-change-comment-owner",
            actor_change,
        ))
        .unwrap();

    store
        .mutate(comment_mutation_request(
            "detached-local-reply",
            NativeOfficeCollaborationMutation::DocumentCommentReply {
                comment_id: COMMENT_ID.to_owned(),
                reply_id: REPLY_ID.to_owned(),
                author: "Ada".to_owned(),
                created_at: "2026-08-17T00:02:00.000Z".to_owned(),
                text: "Replying to a detached comment.".to_owned(),
            },
        ))
        .unwrap();
    store
        .mutate(comment_mutation_request(
            "detached-resolve",
            NativeOfficeCollaborationMutation::DocumentCommentSetResolved {
                comment_id: COMMENT_ID.to_owned(),
                resolved: true,
            },
        ))
        .unwrap();
    let before_forbidden = store.inspect().unwrap();
    let forbidden = store
        .mutate(comment_mutation_request(
            "detached-delete-foreign",
            NativeOfficeCollaborationMutation::DocumentCommentDelete {
                comment_id: COMMENT_ID.to_owned(),
                reply_id: None,
            },
        ))
        .unwrap_err();
    assert_eq!(forbidden.code, "office.collaboration.permission_denied");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_forbidden.document_state_sha256
    );
    store
        .mutate(comment_mutation_request(
            "detached-delete-own-reply",
            NativeOfficeCollaborationMutation::DocumentCommentDelete {
                comment_id: COMMENT_ID.to_owned(),
                reply_id: Some(REPLY_ID.to_owned()),
            },
        ))
        .unwrap();
}

#[test]
fn authenticated_transport_accepts_review_updates_and_rejects_forged_content() {
    let temp = tempfile::tempdir().unwrap();
    let server_root = temp.path().join("authorized-document-server");
    let server = NativeOfficeCollaborationStore::create(NativeOfficeCollaborationCreateRequest {
        store: server_root,
        artifact_id: "fixture-document".to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Document,
        actor_id: "collaboration-server".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::System,
        mode: NativeOfficeCollaborationMode::Edit,
        operation_id: "create-authorized-server".to_owned(),
        namespace: None,
        client_id: Some(900_201),
        initial_update: None,
    })
    .unwrap();
    server
        .apply(NativeOfficeCollaborationApplyRequest {
            operation_id: "bootstrap-authorized-server".to_owned(),
            actor_id: "collaboration-server".to_owned(),
            mode: NativeOfficeCollaborationMode::Edit,
            expected_artifact_id: "fixture-document".to_owned(),
            expected_kind: NativeOfficeCollaborationArtifactKind::Document,
            update: STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap(),
            if_state_vector: None,
            origin: None,
        })
        .unwrap();

    let producer_root = temp.path().join("authorized-comment-producer");
    let producer = NativeOfficeCollaborationStore::create(NativeOfficeCollaborationCreateRequest {
        store: producer_root,
        artifact_id: "fixture-document".to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Document,
        actor_id: "reviewer-1".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Human,
        mode: NativeOfficeCollaborationMode::Comment,
        operation_id: "create-comment-producer".to_owned(),
        namespace: None,
        client_id: Some(900_202),
        initial_update: Some(STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap()),
    })
    .unwrap();
    producer
        .mutate(NativeOfficeCollaborationMutationRequest {
            operation_id: "produce-browser-comment".to_owned(),
            actor_id: "reviewer-1".to_owned(),
            mode: NativeOfficeCollaborationMode::Comment,
            expected_artifact_id: "fixture-document".to_owned(),
            expected_kind: NativeOfficeCollaborationArtifactKind::Document,
            mutation: NativeOfficeCollaborationMutation::DocumentCommentCreate {
                comment_id: "authorized-comment-1".to_owned(),
                paragraph_id: "00000001".to_owned(),
                expected_text_id: "00000002".to_owned(),
                start_utf16: 6,
                end_utf16: 8,
                expected_text: "😀".to_owned(),
                author: "Ada Reviewer".to_owned(),
                created_at: CREATED_AT.to_owned(),
                text: "This review update is ticket scoped.".to_owned(),
            },
            if_state_vector: None,
        })
        .unwrap();
    let comment_update = producer
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(0),
            limit: MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH,
        })
        .unwrap()
        .updates
        .into_iter()
        .find(|event| event.operation_id == "produce-browser-comment")
        .unwrap()
        .update;

    let authorization = NativeOfficeCollaborationTransportAuthorization {
        actor_id: "reviewer-1".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Human,
        actor_name: "Ada Reviewer".to_owned(),
        mode: NativeOfficeCollaborationMode::Comment,
    };
    let mut transport = NativeOfficeCollaborationTransportSession::attach(server.clone()).unwrap();
    let operation_id = "authorized-comment-delivery";
    let accepted = transport
        .receive_authorized(
            NativeOfficeCollaborationTransportReceiveRequest {
                message: authorized_document_message(comment_update, operation_id),
                operation_id: Some(operation_id.to_owned()),
                if_state_vector: None,
            },
            authorization.clone(),
        )
        .unwrap();
    assert!(accepted.apply.unwrap().state_changed);
    let NativeOfficeCollaborationProjectedContent::Document { comments, .. } =
        server.project().unwrap().content
    else {
        panic!("expected Document projection");
    };
    assert_eq!(comments.len(), 1);
    assert_eq!(comments[0].actor_id.as_deref(), Some("reviewer-1"));
    assert_eq!(comments[0].author, "Ada Reviewer");
    let durable = server
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(1),
            limit: 1,
        })
        .unwrap();
    assert_eq!(durable.updates[0].actor_id, "reviewer-1");
    assert_eq!(
        durable.updates[0].actor_kind,
        NativeOfficeCollaborationActorKind::Human
    );
    assert_eq!(
        durable.updates[0].mode,
        NativeOfficeCollaborationMode::Comment
    );

    let attacker = Doc::with_client_id(818_484);
    attacker
        .transact_mut()
        .apply_update(Update::decode_v1(&server.synchronize(None).unwrap().update).unwrap())
        .unwrap();
    let before_attack = attacker.transact().state_vector();
    let fragment = attacker.get_or_insert_xml_fragment("a3s.office.document.content");
    let text = fragment
        .successors(&attacker.transact())
        .find_map(|node| match node {
            XmlOut::Text(text) => Some(text),
            _ => None,
        })
        .unwrap();
    text.insert(&mut attacker.transact_mut(), 0, "FORGED ");
    let forged_update = attacker
        .transact()
        .encode_state_as_update_v1(&before_attack);
    let before_rejection = server.inspect().unwrap();
    let forged_operation = "forged-comment-content-delivery";
    let rejected = transport
        .receive_authorized(
            NativeOfficeCollaborationTransportReceiveRequest {
                message: authorized_document_message(forged_update.clone(), forged_operation),
                operation_id: Some(forged_operation.to_owned()),
                if_state_vector: None,
            },
            authorization,
        )
        .unwrap_err();
    assert_eq!(rejected.code, "office.collaboration.permission_denied");
    assert_eq!(
        server.inspect().unwrap().document_state_sha256,
        before_rejection.document_state_sha256
    );

    let view_operation = "forged-view-content-delivery";
    let view_rejected = transport
        .receive_authorized(
            NativeOfficeCollaborationTransportReceiveRequest {
                message: authorized_document_message(forged_update, view_operation),
                operation_id: Some(view_operation.to_owned()),
                if_state_vector: None,
            },
            NativeOfficeCollaborationTransportAuthorization {
                actor_id: "viewer-1".to_owned(),
                actor_kind: NativeOfficeCollaborationActorKind::Human,
                actor_name: "Vera Viewer".to_owned(),
                mode: NativeOfficeCollaborationMode::View,
            },
        )
        .unwrap_err();
    assert_eq!(view_rejected.code, "office.collaboration.permission_denied");
}

fn authorized_document_message(
    update: Vec<u8>,
    operation_id: &str,
) -> NativeOfficeCollaborationTransportMessage {
    NativeOfficeCollaborationTransportMessage {
        protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
        version: NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
        artifact_id: "fixture-document".to_owned(),
        artifact_kind: NativeOfficeCollaborationArtifactKind::Document,
        namespace: NATIVE_OFFICE_COLLABORATION_NAMESPACE.to_owned(),
        sender_client_id: 818_485,
        message_type: NativeOfficeCollaborationTransportMessageType::Update,
        payload: update,
        origin: Some(NativeOfficeCollaborationOrigin {
            protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
            kind: NativeOfficeCollaborationOriginKind::Editor,
            actor_id: Some("reviewer-1".to_owned()),
            operation_id: Some(operation_id.to_owned()),
        }),
    }
}

fn create_comment_mutation(text: &str) -> NativeOfficeCollaborationMutation {
    create_comment_mutation_at(text, 6, 8, "😀")
}

fn create_comment_mutation_at(
    text: &str,
    start_utf16: u32,
    end_utf16: u32,
    expected_text: &str,
) -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::DocumentCommentCreate {
        comment_id: COMMENT_ID.to_owned(),
        paragraph_id: "00000001".to_owned(),
        expected_text_id: "00000002".to_owned(),
        start_utf16,
        end_utf16,
        expected_text: expected_text.to_owned(),
        author: "Ada".to_owned(),
        created_at: CREATED_AT.to_owned(),
        text: text.to_owned(),
    }
}

fn comment_create_request(root: &std::path::Path) -> NativeOfficeCollaborationCreateRequest {
    mode_create_request(root, NativeOfficeCollaborationMode::Comment)
}

fn mode_create_request(
    root: &std::path::Path,
    mode: NativeOfficeCollaborationMode,
) -> NativeOfficeCollaborationCreateRequest {
    NativeOfficeCollaborationCreateRequest {
        store: root.to_path_buf(),
        artifact_id: "fixture-document".to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Document,
        actor_id: "agent-alpha".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Agent,
        mode,
        operation_id: format!("create-{}", mode.as_str()),
        namespace: None,
        client_id: Some(900_102 + mode as u64),
        initial_update: None,
    }
}

fn comment_apply_request(
    operation_id: &str,
    update: Vec<u8>,
) -> NativeOfficeCollaborationApplyRequest {
    mode_apply_request(operation_id, NativeOfficeCollaborationMode::Comment, update)
}

fn mode_apply_request(
    operation_id: &str,
    mode: NativeOfficeCollaborationMode,
    update: Vec<u8>,
) -> NativeOfficeCollaborationApplyRequest {
    NativeOfficeCollaborationApplyRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-alpha".to_owned(),
        mode,
        expected_artifact_id: "fixture-document".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Document,
        update,
        if_state_vector: None,
        origin: None,
    }
}

fn comment_mutation_request(
    operation_id: &str,
    mutation: NativeOfficeCollaborationMutation,
) -> NativeOfficeCollaborationMutationRequest {
    mode_mutation_request(
        operation_id,
        NativeOfficeCollaborationMode::Comment,
        mutation,
    )
}

fn mode_mutation_request(
    operation_id: &str,
    mode: NativeOfficeCollaborationMode,
    mutation: NativeOfficeCollaborationMutation,
) -> NativeOfficeCollaborationMutationRequest {
    NativeOfficeCollaborationMutationRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-alpha".to_owned(),
        mode,
        expected_artifact_id: "fixture-document".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Document,
        mutation,
        if_state_vector: None,
    }
}

fn assert_browser_comment_state(
    store: &NativeOfficeCollaborationStore,
    comment_present: bool,
    resolved: bool,
    reply_present: bool,
) {
    let exported = store.synchronize(None).unwrap();
    let browser = Doc::with_client_id(818_282);
    browser
        .transact_mut()
        .apply_update(Update::decode_v1(&exported.update).unwrap())
        .unwrap();
    let comments = browser.get_or_insert_map("a3s.office.document.comments");
    let order = browser.get_or_insert_array("a3s.office.document.comment-order");
    let claims = browser.get_or_insert_array("a3s.office.document.record-claims");
    let fragment = browser.get_or_insert_xml_fragment("a3s.office.document.content");
    let transaction = browser.transact();

    if !comment_present {
        assert_eq!(comments.len(&transaction), 0);
        assert_eq!(order.len(&transaction), 0);
        assert!(!fragment.successors(&transaction).any(|node| match node {
            XmlOut::Text(text) => text.diff(&transaction, |_| ()).iter().any(|chunk| {
                chunk.attributes.as_ref().is_some_and(|attributes| {
                    attributes
                        .keys()
                        .any(|key| key.starts_with("documentComment"))
                })
            }),
            _ => false,
        }));
        return;
    }

    assert!(claims
        .to_json(&transaction)
        .to_string()
        .contains(COMMENT_ID));
    assert_eq!(
        order.get(&transaction, 0),
        Some(Out::Any(Any::String(COMMENT_ID.into())))
    );
    let Some(Out::YMap(comment)) = comments.get(&transaction, COMMENT_ID) else {
        panic!("browser comment record missing");
    };
    assert_eq!(
        comment.get(&transaction, "actorId"),
        Some(Out::Any(Any::String("agent-alpha".into())))
    );
    assert_eq!(
        comment.get(&transaction, "resolved"),
        Some(Out::Any(Any::Bool(resolved)))
    );
    let Some(Out::YMap(replies)) = comment.get(&transaction, "replies") else {
        panic!("browser reply map missing");
    };
    let Some(Out::YArray(reply_order)) = comment.get(&transaction, "replyOrder") else {
        panic!("browser reply order missing");
    };
    assert_eq!(replies.len(&transaction), u32::from(reply_present));
    assert_eq!(reply_order.len(&transaction), u32::from(reply_present));
    let anchored = fragment.successors(&transaction).any(|node| match node {
        XmlOut::Text(text) => text.diff(&transaction, |_| ()).iter().any(|chunk| {
            chunk.attributes.as_ref().is_some_and(|attributes| {
                attributes.iter().any(|(key, value)| {
                    key.starts_with("documentComment")
                        && matches!(value, Any::Map(fields) if matches!(fields.get("id"), Some(Any::String(id)) if id.as_ref() == COMMENT_ID))
                })
            })
        }),
        _ => false,
    });
    assert!(anchored);
}
