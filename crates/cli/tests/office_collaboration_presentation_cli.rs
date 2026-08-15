mod support;

use std::fs;
use std::path::Path;
use std::process::{Command, Output};

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use serde_json::{json, Value as JsonValue};
use support::{
    presentation_collaboration_fixture, presentation_element, presentation_element_order,
    presentation_element_tombstoned, presentation_scene_element, presentation_slide_body_element,
    presentation_slide_title_element,
};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

fn execute(args: &[&str]) -> Output {
    Command::new(binary()).args(args).output().unwrap()
}

fn run(args: &[&str]) -> JsonValue {
    let output = execute(args);
    assert!(output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

fn run_failure(args: &[&str]) -> JsonValue {
    let output = execute(args);
    assert!(!output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

#[test]
fn cli_manages_browser_compatible_presentation_scene_elements() {
    let temp = tempfile::tempdir().unwrap();
    let replica = temp.path().join("presentation.replica");
    let fixture = temp.path().join("browser-presentation.update");
    fs::write(&fixture, presentation_collaboration_fixture()).unwrap();
    run(&[
        "collab",
        "join",
        replica.to_str().unwrap(),
        "--artifact-id",
        "fixture-presentation",
        "--kind",
        "presentation",
        "--actor-id",
        "coding-agent-presentation-cli",
        "--actor-kind",
        "agent",
        "--mode",
        "edit",
        "--operation-id",
        "join-browser-presentation-cli",
        "--input",
        fixture.to_str().unwrap(),
        "--client-id",
        "900016",
        "--json",
    ]);

    let expected = presentation_slide_title_element();
    let mut next = expected.clone();
    next["text"] = json!("CLI shared title");
    next["x"] = json!(18);
    let updated = mutate(
        &replica,
        "presentation-update-element-cli",
        json!({
            "type": "presentation-update-element",
            "containerKind": "slide",
            "containerId": "slide-1",
            "elementId": "element-title",
            "expectedElement": expected,
            "nextElement": next,
        }),
    );
    assert_eq!(updated["data"]["sequence"], 1);
    assert_eq!(
        updated["data"]["mutation"]["type"],
        "presentation-update-element"
    );

    let created_element = presentation_scene_element("element-cli", "CLI object", "shape");
    let created = mutate(
        &replica,
        "presentation-create-element-cli",
        json!({
            "type": "presentation-create-element",
            "containerKind": "slide",
            "containerId": "slide-1",
            "element": created_element,
            "afterElementId": "element-title",
        }),
    );
    assert_eq!(created["data"]["sequence"], 2);

    let moved = mutate(
        &replica,
        "presentation-move-element-cli",
        json!({
            "type": "presentation-move-element",
            "containerKind": "slide",
            "containerId": "slide-1",
            "elementId": "element-cli",
            "expectedAfterElementId": "element-title",
            "afterElementId": null,
        }),
    );
    assert_eq!(moved["data"]["sequence"], 3);

    let deleted = mutate(
        &replica,
        "presentation-delete-element-cli",
        json!({
            "type": "presentation-delete-element",
            "containerKind": "slide",
            "containerId": "slide-2",
            "expectedElement": presentation_slide_body_element(),
        }),
    );
    assert_eq!(deleted["data"]["sequence"], 4);

    let mut stale_next = presentation_slide_title_element();
    stale_next["text"] = json!("Stale CLI title");
    let mutation = json!({
        "type": "presentation-update-element",
        "containerKind": "slide",
        "containerId": "slide-1",
        "elementId": "element-title",
        "expectedElement": presentation_slide_title_element(),
        "nextElement": stale_next,
    })
    .to_string();
    let conflict = run_failure(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        mutation.as_str(),
        "--actor-id",
        "coding-agent-presentation-cli",
        "--operation-id",
        "presentation-conflict-cli",
        "--artifact-id",
        "fixture-presentation",
        "--kind",
        "presentation",
        "--mode",
        "edit",
        "--json",
    ]);
    assert_eq!(
        conflict["error"]["code"],
        "office.collaboration.mutation_match_conflict"
    );

    let exported = run(&["collab", "diff", replica.to_str().unwrap(), "--json"]);
    let update = STANDARD
        .decode(exported["data"]["updateBase64"].as_str().unwrap())
        .unwrap();
    assert_eq!(
        presentation_element(&update, "slides", "slide-1", "element-title").unwrap()["text"],
        "CLI shared title"
    );
    assert_eq!(
        presentation_element(&update, "slides", "slide-1", "element-title").unwrap()["x"],
        18
    );
    assert_eq!(
        presentation_element(&update, "slides", "slide-1", "element-cli").unwrap()["text"],
        "CLI object"
    );
    assert_eq!(
        presentation_element_order(&update, "slides", "slide-1"),
        vec!["element-cli".to_owned(), "element-title".to_owned()]
    );
    assert_eq!(
        presentation_element(&update, "slides", "slide-2", "element-body"),
        None
    );
    assert!(presentation_element_tombstoned(
        &update,
        "slides",
        "slide-2",
        "element-body"
    ));
}

fn mutate(replica: &Path, operation_id: &str, mutation: JsonValue) -> JsonValue {
    let mutation = mutation.to_string();
    run(&[
        "collab",
        "mutate",
        replica.to_str().unwrap(),
        "--mutation",
        mutation.as_str(),
        "--actor-id",
        "coding-agent-presentation-cli",
        "--operation-id",
        operation_id,
        "--artifact-id",
        "fixture-presentation",
        "--kind",
        "presentation",
        "--mode",
        "edit",
        "--json",
    ])
}
