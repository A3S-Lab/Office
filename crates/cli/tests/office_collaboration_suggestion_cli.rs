use std::fs;
use std::path::Path;
use std::process::{Command, Output};

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use yrs::updates::decoder::Decode;
use yrs::{Any, Array, Doc, Map, Out, Transact, Update};

const YJS_DOCUMENT_UPDATE_BASE64: &str = "AQ+y8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtZG9jdW1lbnQoARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhkb2N1bWVudCgBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjQyNDI0Mjpicm93c2VyLWZpeHR1cmUHARthM3Mub2ZmaWNlLmRvY3VtZW50LmNvbnRlbnQDD2RvY3VtZW50U2VjdGlvbgcAsvIZBgMJcGFyYWdyYXBoBwCy8hkHBgQAsvIZCBBIZWxsbyDwn5iAIHdvcmxkKACy8hkHC3BhcmFncmFwaElkAXcIMDAwMDAwMDEoALLyGQcGdGV4dElkAXcIMDAwMDAwMDIoALLyGQYCaWQBdxJkb2N1bWVudC1zZWN0aW9uLTEoARthM3Mub2ZmaWNlLmRvY3VtZW50Lm9wdGlvbnMJcGFnZUNvbG9yAXcHI0Y4RkFGQygBG2Ezcy5vZmZpY2UuZG9jdW1lbnQub3B0aW9ucwx0cmFja0NoYW5nZXMBeAA=";

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

fn join(replica: &Path, input: &Path, actor_id: &str, mode: &str, client_id: &str) {
    let result = run(&[
        "collab",
        "join",
        replica.to_str().unwrap(),
        "--artifact-id",
        "fixture-document",
        "--kind",
        "document",
        "--actor-id",
        actor_id,
        "--actor-kind",
        "agent",
        "--mode",
        mode,
        "--operation-id",
        &format!("join-{actor_id}"),
        "--input",
        input.to_str().unwrap(),
        "--client-id",
        client_id,
        "--json",
    ]);
    assert_eq!(result["data"]["replica"]["manifest"]["mode"], mode);
}

#[test]
fn cli_creates_syncs_and_finalizes_native_document_suggestions() {
    let temp = tempfile::tempdir().unwrap();
    let input = temp.path().join("document.update");
    let proposal = temp.path().join("proposal.update");
    let final_update = temp.path().join("final.update");
    let suggester = temp.path().join("suggester.replica");
    let editor = temp.path().join("editor.replica");
    fs::write(&input, STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap()).unwrap();
    join(&suggester, &input, "agent-suggester", "suggest", "901201");
    join(&editor, &input, "human-editor", "edit", "901202");

    let created = run(&[
        "collab",
        "mutate",
        suggester.to_str().unwrap(),
        "--mutation",
        r#"{"type":"document-suggestion-create","paragraphId":"00000001","expectedTextId":"00000002","startUtf16":6,"endUtf16":8,"expectedText":"😀","replacement":"collaborative","insertionId":"cli-insertion","deletionId":"cli-deletion","author":"A3S Agent","createdAt":"2026-08-17T10:00:00.000Z"}"#,
        "--actor-id",
        "agent-suggester",
        "--operation-id",
        "create-cli-replacement",
        "--artifact-id",
        "fixture-document",
        "--kind",
        "document",
        "--mode",
        "suggest",
        "--json",
    ]);
    assert_eq!(
        created["data"]["mutation"]["type"],
        "document-suggestion-create"
    );
    assert_eq!(created["data"]["sequence"], 1);

    run(&[
        "collab",
        "diff",
        suggester.to_str().unwrap(),
        "--output",
        proposal.to_str().unwrap(),
        "--json",
    ]);
    run(&[
        "collab",
        "apply",
        editor.to_str().unwrap(),
        "--input",
        proposal.to_str().unwrap(),
        "--actor-id",
        "human-editor",
        "--operation-id",
        "sync-cli-replacement",
        "--artifact-id",
        "fixture-document",
        "--kind",
        "document",
        "--mode",
        "edit",
        "--json",
    ]);
    let pending = run(&["collab", "read", editor.to_str().unwrap(), "--json"]);
    assert_eq!(pending["data"]["version"], 3);
    assert_eq!(
        pending["data"]["content"]["suggestions"]
            .as_array()
            .unwrap()
            .len(),
        2
    );
    assert_eq!(
        pending["data"]["content"]["changeDecisions"],
        serde_json::json!([])
    );

    let decided = run(&[
        "collab",
        "mutate",
        editor.to_str().unwrap(),
        "--mutation",
        r#"{"type":"document-suggestion-decide","suggestions":[{"id":"cli-deletion","kind":"deletion","expectedActorId":"agent-suggester","expectedAuthor":"A3S Agent","expectedCreatedAt":"2026-08-17T10:00:00.000Z","expectedText":"😀"},{"id":"cli-insertion","kind":"insertion","expectedActorId":"agent-suggester","expectedAuthor":"A3S Agent","expectedCreatedAt":"2026-08-17T10:00:00.000Z","expectedText":"collaborative"}],"decision":"accept","decidedBy":"Grace Editor","decidedAt":"2026-08-17T10:01:00.000Z"}"#,
        "--actor-id",
        "human-editor",
        "--operation-id",
        "accept-cli-replacement",
        "--artifact-id",
        "fixture-document",
        "--kind",
        "document",
        "--mode",
        "edit",
        "--json",
    ]);
    assert_eq!(
        decided["data"]["mutation"]["type"],
        "document-suggestion-decide"
    );

    let final_projection = run(&["collab", "read", editor.to_str().unwrap(), "--json"]);
    let content = &final_projection["data"]["content"];
    assert_eq!(content["plainText"], "Hello collaborative world");
    assert_eq!(content["paragraphs"][0]["textId"], "00000003");
    assert_eq!(content["suggestions"], serde_json::json!([]));
    assert_eq!(content["changeDecisions"].as_array().unwrap().len(), 2);
    assert!(content["changeDecisions"]
        .as_array()
        .unwrap()
        .iter()
        .all(
            |record| record["decision"] == "accept" && record["decidedByActorId"] == "human-editor"
        ));

    run(&[
        "collab",
        "diff",
        editor.to_str().unwrap(),
        "--output",
        final_update.to_str().unwrap(),
        "--json",
    ]);
    assert_browser_sidecar_roots(&fs::read(final_update).unwrap());
}

fn assert_browser_sidecar_roots(update: &[u8]) {
    let document = Doc::new();
    document.get_or_insert_map("a3s.office.metadata");
    document.get_or_insert_array("a3s.office.bootstrap.initializers");
    document.get_or_insert_xml_fragment("a3s.office.document.content");
    document.get_or_insert_map("a3s.office.document.options");
    document.get_or_insert_map("a3s.office.document.comments");
    document.get_or_insert_array("a3s.office.document.comment-order");
    let decisions = document.get_or_insert_map("a3s.office.document.change-decisions");
    let order = document.get_or_insert_array("a3s.office.document.change-decision-order");
    let claims = document.get_or_insert_array("a3s.office.document.record-claims");
    document
        .transact_mut()
        .apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
    let transaction = document.transact();
    assert_eq!(decisions.len(&transaction), 2);
    assert_eq!(order.len(&transaction), 2);
    assert!(claims.iter(&transaction).any(|claim| {
        matches!(claim, Out::Any(Any::String(value)) if value.contains("change-decision"))
    }));
}
