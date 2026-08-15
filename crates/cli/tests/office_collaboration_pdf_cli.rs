mod support;

use std::fs;
use std::process::{Command, Output};

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use support::{pdf_collaboration_fixture, pdf_snapshot};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

fn execute(args: &[&str]) -> Output {
    Command::new(binary()).args(args).output().unwrap()
}

fn run(args: &[&str]) -> serde_json::Value {
    let output = execute(args);
    assert!(output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

fn run_failure(args: &[&str]) -> serde_json::Value {
    let output = execute(args);
    assert!(!output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

#[test]
fn cli_sets_existing_and_new_pdf_form_values() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("pdf.replica");
    let fixture = temp.path().join("browser-pdf.update");
    fs::write(&fixture, pdf_collaboration_fixture()).unwrap();
    run(&[
        "collab",
        "join",
        replica.to_str().unwrap(),
        "--artifact-id",
        "fixture-pdf",
        "--kind",
        "pdf",
        "--actor-id",
        "coding-agent-pdf-cli",
        "--actor-kind",
        "agent",
        "--mode",
        "edit",
        "--operation-id",
        "join-browser-pdf-cli",
        "--input",
        fixture.to_str().unwrap(),
        "--client-id",
        "900013",
        "--json",
    ]);

    let name = mutate(
        &replica,
        "pdf-set-name-cli",
        r#"{"type":"pdf-set-form-value","fieldId":"Applicant.Name","value":"Grace"}"#,
    );
    assert_eq!(name["data"]["sequence"], 1);
    assert_eq!(name["data"]["mutation"]["type"], "pdf-set-form-value");
    let email = mutate(
        &replica,
        "pdf-set-email-cli",
        r#"{"type":"pdf-set-form-value","fieldId":"Applicant.Email","value":"grace@example.test"}"#,
    );
    assert_eq!(email["data"]["sequence"], 2);

    let replay = mutate(
        &replica,
        "pdf-set-name-cli",
        r#"{"type":"pdf-set-form-value","fieldId":"Applicant.Name","value":"Grace"}"#,
    );
    assert_eq!(replay["data"]["duplicate"], true);
    assert_eq!(replay["data"]["sequence"], 1);

    let invalid = run_failure(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        r#"{"type":"pdf-set-form-value","fieldId":" Applicant.Name","value":"Rejected"}"#,
        "--actor-id",
        "coding-agent-pdf-cli",
        "--operation-id",
        "pdf-invalid-field-cli",
        "--artifact-id",
        "fixture-pdf",
        "--kind",
        "pdf",
        "--mode",
        "edit",
        "--json",
    ]);
    assert_eq!(
        invalid["error"]["code"],
        "office.collaboration.mutation_invalid"
    );

    let exported = run(&["collab", "diff", replica.to_str().unwrap(), "--json"]);
    let update = STANDARD
        .decode(exported["data"]["updateBase64"].as_str().unwrap())
        .unwrap();
    let snapshot = pdf_snapshot(&update);
    assert_eq!(snapshot.form_values["Applicant.Name"], "Grace");
    assert_eq!(
        snapshot.form_values["Applicant.Email"],
        "grace@example.test"
    );
    assert_eq!(
        snapshot.form_order,
        vec!["Applicant.Name".to_owned(), "Applicant.Email".to_owned()]
    );
    let inspected = run(&["collab", "inspect", replica.to_str().unwrap(), "--json"]);
    assert_eq!(inspected["data"]["currentSequence"], 2);
}

#[test]
fn cli_creates_updates_and_tombstones_portable_pdf_annotations() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("pdf-annotations.replica");
    let fixture = temp.path().join("browser-pdf.update");
    fs::write(&fixture, pdf_collaboration_fixture()).unwrap();
    run(&[
        "collab",
        "join",
        replica.to_str().unwrap(),
        "--artifact-id",
        "fixture-pdf",
        "--kind",
        "pdf",
        "--actor-id",
        "coding-agent-pdf-cli",
        "--actor-kind",
        "agent",
        "--mode",
        "edit",
        "--operation-id",
        "join-browser-pdf-annotation-cli",
        "--input",
        fixture.to_str().unwrap(),
        "--client-id",
        "900017",
        "--json",
    ]);

    let annotation = serde_json::json!({
        "id": "annotation-cli-1",
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
        "strokeColor": "#ffd400",
        "color": "#ffd400",
        "opacity": 0.48,
        "contents": "CLI annotation",
        "author": "A3S Agent",
        "created": "2026-08-15T08:00:00.000Z",
    });
    let create = serde_json::json!({
        "type": "pdf-create-annotation",
        "annotationId": "annotation-cli-1",
        "pageIndex": 1,
        "annotation": annotation,
    });
    let created = mutate(&replica, "pdf-create-annotation-cli", &create.to_string());
    assert_eq!(created["data"]["sequence"], 1);
    assert_eq!(created["data"]["mutation"]["type"], "pdf-create-annotation");

    let mut next = annotation.clone();
    next["strokeColor"] = serde_json::json!("#ff0000");
    next["color"] = serde_json::json!("#ff0000");
    next["contents"] = serde_json::json!("Updated from CLI");
    let update = serde_json::json!({
        "type": "pdf-update-annotation",
        "annotationId": "annotation-cli-1",
        "expectedAnnotation": annotation,
        "nextAnnotation": next,
    });
    let updated = mutate(&replica, "pdf-update-annotation-cli", &update.to_string());
    assert_eq!(updated["data"]["sequence"], 2);

    let deleted = mutate(
        &replica,
        "pdf-delete-annotation-cli",
        r#"{"type":"pdf-delete-annotation","annotationId":"annotation-cli-1","expectedSource":"created","expectedPageIndex":1,"expectedType":9}"#,
    );
    assert_eq!(deleted["data"]["sequence"], 3);

    let exported = run(&["collab", "diff", replica.to_str().unwrap(), "--json"]);
    let update = STANDARD
        .decode(exported["data"]["updateBase64"].as_str().unwrap())
        .unwrap();
    let snapshot = pdf_snapshot(&update);
    let record = &snapshot.annotations["annotation-cli-1"];
    assert_eq!(record["source"], "created");
    assert_eq!(record["deleted"], true);
    assert_eq!(record["annotation"]["strokeColor"], "#ff0000");
    assert_eq!(record["annotation"]["color"], "#ff0000");
    assert_eq!(
        record["annotation"]["segmentRects"][0]["size"]["width"],
        300.0
    );
    assert_eq!(record["annotation"]["contents"], "Updated from CLI");
    assert!(snapshot.record_claims.iter().any(|raw| {
        let claim: serde_json::Value = serde_json::from_str(raw).unwrap();
        claim["kind"] == "annotation" && claim["id"] == "annotation-cli-1"
    }));
}

#[test]
fn cli_appends_attributable_pdf_review_records_and_claims() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("pdf-review.replica");
    let fixture = temp.path().join("browser-pdf.update");
    fs::write(&fixture, pdf_collaboration_fixture()).unwrap();
    run(&[
        "collab",
        "join",
        replica.to_str().unwrap(),
        "--artifact-id",
        "fixture-pdf",
        "--kind",
        "pdf",
        "--actor-id",
        "coding-agent-pdf-cli",
        "--actor-kind",
        "agent",
        "--mode",
        "edit",
        "--operation-id",
        "join-browser-pdf-review-cli",
        "--input",
        fixture.to_str().unwrap(),
        "--client-id",
        "900015",
        "--json",
    ]);

    let proposal = r#"{"type":"pdf-propose-redaction","proposalId":"redaction-cli-1","pageIndex":1,"rects":[{"left":10.5,"top":20.25,"right":80.75,"bottom":40.5}],"proposedAt":"2026-08-15T05:00:00.000Z","reason":"Personal data","text":"Account 1234"}"#;
    let proposed = mutate(&replica, "pdf-propose-redaction-cli", proposal);
    assert_eq!(proposed["data"]["sequence"], 1);
    assert_eq!(proposed["data"]["stateChanged"], true);
    assert_eq!(
        proposed["data"]["mutation"]["type"],
        "pdf-propose-redaction"
    );

    let retry = mutate(&replica, "pdf-propose-redaction-cli-retry", proposal);
    assert_eq!(retry["data"]["stateChanged"], false);
    assert!(retry["data"]["sequence"].is_null());

    let decided = mutate(
        &replica,
        "pdf-decide-review-cli",
        r#"{"type":"pdf-decide-review","decisionId":"decision-cli-1","targetKind":"redaction","targetId":"redaction-cli-1","decision":"approve","createdAt":"2026-08-15T05:05:00.000Z"}"#,
    );
    assert_eq!(decided["data"]["sequence"], 2);
    assert_eq!(decided["data"]["mutation"]["type"], "pdf-decide-review");

    let conflict = mutate_failure(
        &replica,
        "pdf-decide-review-cli-conflict",
        r#"{"type":"pdf-decide-review","decisionId":"decision-cli-2","targetKind":"redaction","targetId":"redaction-cli-1","decision":"reject","createdAt":"2026-08-15T05:06:00.000Z"}"#,
    );
    assert_eq!(
        conflict["error"]["code"],
        "office.collaboration.mutation_match_conflict"
    );

    let rotation = mutate(
        &replica,
        "pdf-propose-page-rotation-cli",
        r#"{"type":"pdf-propose-page-rotation","pageOperationId":"page-operation-rotate-cli","pageIndices":[0,2],"degrees":90,"proposedAt":"2026-08-15T05:10:00.000Z"}"#,
    );
    assert_eq!(rotation["data"]["sequence"], 3);
    let deletion = mutate(
        &replica,
        "pdf-propose-page-deletion-cli",
        r#"{"type":"pdf-propose-page-deletion","pageOperationId":"page-operation-delete-cli","pageIndices":[2],"proposedAt":"2026-08-15T05:11:00.000Z"}"#,
    );
    assert_eq!(deletion["data"]["sequence"], 4);
    let reorder = mutate(
        &replica,
        "pdf-propose-page-reorder-cli",
        r#"{"type":"pdf-propose-page-reorder","pageOperationId":"page-operation-reorder-cli","pageOrder":[2,0,1],"proposedAt":"2026-08-15T05:12:00.000Z"}"#,
    );
    assert_eq!(reorder["data"]["sequence"], 5);
    let page_decision = mutate(
        &replica,
        "pdf-decide-page-operation-cli",
        r#"{"type":"pdf-decide-review","decisionId":"decision-page-operation-cli","targetKind":"page-operation","targetId":"page-operation-delete-cli","decision":"reject","createdAt":"2026-08-15T05:15:00.000Z"}"#,
    );
    assert_eq!(page_decision["data"]["sequence"], 6);

    let exported = run(&["collab", "diff", replica.to_str().unwrap(), "--json"]);
    let update = STANDARD
        .decode(exported["data"]["updateBase64"].as_str().unwrap())
        .unwrap();
    let snapshot = pdf_snapshot(&update);
    assert_eq!(
        snapshot.redaction_proposals["redaction-cli-1"],
        serde_json::json!({
            "id": "redaction-cli-1",
            "pageIndex": 1,
            "rects": [{ "left": 10.5, "top": 20.25, "right": 80.75, "bottom": 40.5 }],
            "proposedBy": "coding-agent-pdf-cli",
            "proposedAt": "2026-08-15T05:00:00.000Z",
            "reason": "Personal data",
            "text": "Account 1234"
        })
    );
    assert_eq!(
        snapshot.review_decisions["decision-cli-1"]["actorId"],
        "coding-agent-pdf-cli"
    );
    assert_eq!(
        snapshot.page_operations["page-operation-rotate-cli"],
        serde_json::json!({
            "id": "page-operation-rotate-cli",
            "kind": "rotate",
            "pageIndices": [0, 2],
            "degrees": 90,
            "proposedBy": "coding-agent-pdf-cli",
            "proposedAt": "2026-08-15T05:10:00.000Z"
        })
    );
    assert_eq!(
        snapshot.page_operations["page-operation-delete-cli"]["kind"],
        "delete"
    );
    assert_eq!(
        snapshot.page_operations["page-operation-reorder-cli"]["pageOrder"],
        serde_json::json!([2, 0, 1])
    );
    assert_eq!(
        snapshot.review_decisions["decision-page-operation-cli"]["targetKind"],
        "page-operation"
    );
    assert_eq!(snapshot.record_claims.len(), 6);
    let claims = snapshot
        .record_claims
        .iter()
        .map(|claim| serde_json::from_str::<serde_json::Value>(claim).unwrap())
        .collect::<Vec<_>>();
    assert!(claims
        .iter()
        .any(|claim| { claim["kind"] == "redaction" && claim["id"] == "redaction-cli-1" }));
    assert!(claims
        .iter()
        .any(|claim| { claim["kind"] == "review-decision" && claim["id"] == "decision-cli-1" }));
    assert!(claims.iter().any(|claim| {
        claim["kind"] == "page-operation" && claim["id"] == "page-operation-reorder-cli"
    }));
}

fn mutate(replica: &std::path::Path, operation_id: &str, mutation: &str) -> serde_json::Value {
    run(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        mutation,
        "--actor-id",
        "coding-agent-pdf-cli",
        "--operation-id",
        operation_id,
        "--artifact-id",
        "fixture-pdf",
        "--kind",
        "pdf",
        "--mode",
        "edit",
        "--json",
    ])
}

fn mutate_failure(
    replica: &std::path::Path,
    operation_id: &str,
    mutation: &str,
) -> serde_json::Value {
    run_failure(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        mutation,
        "--actor-id",
        "coding-agent-pdf-cli",
        "--operation-id",
        operation_id,
        "--artifact-id",
        "fixture-pdf",
        "--kind",
        "pdf",
        "--mode",
        "edit",
        "--json",
    ])
}
