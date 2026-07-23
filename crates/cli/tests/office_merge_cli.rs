use std::path::Path;
use std::process::{Command, Output};

fn binary() -> &'static str {
    env!("CARGO_BIN_EXE_a3s-office")
}

fn execute(provider: &Path, args: &[&str]) -> Output {
    Command::new(binary())
        .args(args)
        .env("A3S_OFFICECLI_EXECUTABLE", provider)
        .output()
        .unwrap()
}

fn run(provider: &Path, args: &[&str]) -> serde_json::Value {
    let output = execute(provider, args);
    assert!(output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

fn run_failure(provider: &Path, args: &[&str]) -> serde_json::Value {
    let output = execute(provider, args);
    assert!(!output.status.success(), "{output:?}");
    serde_json::from_slice(&output.stdout).unwrap()
}

fn create(provider: &Path, document: &Path) {
    run(provider, &["create", document.to_str().unwrap(), "--json"]);
}

fn text_view(provider: &Path, document: &Path) -> String {
    if document.extension().and_then(|value| value.to_str()) == Some("xlsx") {
        let value = run(
            provider,
            &["get", document.to_str().unwrap(), "/Sheet1/A1", "--json"],
        );
        return value["data"]["node"]["text"].as_str().unwrap().to_string();
    }
    let value = run(
        provider,
        &["view", document.to_str().unwrap(), "text", "--json"],
    );
    value["data"]["result"]["text"]
        .as_str()
        .unwrap()
        .to_string()
}

fn populate_template(provider: &Path, document: &Path) {
    match document.extension().and_then(|value| value.to_str()) {
        Some("docx") => {
            run(
                provider,
                &[
                    "set",
                    document.to_str().unwrap(),
                    "/body/p[1]",
                    "--text",
                    "Hello {{user.name}}",
                    "--json",
                ],
            );
        }
        Some("xlsx") => {
            run(
                provider,
                &[
                    "set",
                    document.to_str().unwrap(),
                    "/Sheet1/A1",
                    "--text",
                    "Hello {{user.name}}",
                    "--json",
                ],
            );
        }
        Some("pptx") => {
            run(
                provider,
                &[
                    "add",
                    document.to_str().unwrap(),
                    "/",
                    "--type",
                    "slide",
                    "--text",
                    "Hello {{user.name}}",
                    "--json",
                ],
            );
        }
        extension => panic!("unexpected extension: {extension:?}"),
    }
}

#[test]
fn native_merge_copies_all_formats_and_accepts_each_data_source_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let explicit_data = temp.path().join("explicit.json");
    let implicit_data = temp.path().join("implicit.json");
    std::fs::write(&explicit_data, br#"{"user":{"name":"Alice"}}"#).unwrap();
    std::fs::write(&implicit_data, br#"{"user":{"name":"Bob"}}"#).unwrap();

    for (extension, data, expected, expected_source) in [
        (
            "docx",
            format!("@{}", explicit_data.display()),
            "Hello Alice",
            explicit_data.display().to_string(),
        ),
        (
            "xlsx",
            implicit_data.display().to_string(),
            "Hello Bob",
            implicit_data.display().to_string(),
        ),
        (
            "pptx",
            r#"{"user":{"name":"Carol"}}"#.to_string(),
            "Hello Carol",
            "inline".to_string(),
        ),
    ] {
        let template = temp.path().join(format!("template.{extension}"));
        let output = temp.path().join(format!("merged.{extension}"));
        create(&provider, &template);
        populate_template(&provider, &template);
        let template_before = std::fs::read(&template).unwrap();

        let merged = run(
            &provider,
            &[
                "merge",
                template.to_str().unwrap(),
                output.to_str().unwrap(),
                "--data",
                &data,
                "--json",
            ],
        );

        assert_eq!(merged["data"]["operation"], "merge");
        assert_eq!(merged["data"]["atomic"], true);
        assert_eq!(merged["data"]["changed"], true);
        assert_eq!(merged["data"]["force"], false);
        assert_eq!(merged["data"]["dataSource"], expected_source);
        assert_eq!(merged["data"]["result"]["replacedCount"], 1);
        assert_eq!(
            merged["data"]["result"]["usedKeys"],
            serde_json::json!(["user.name"])
        );
        assert_eq!(text_view(&provider, &output), expected);
        assert_eq!(text_view(&provider, &template), "Hello {{user.name}}");
        assert_eq!(std::fs::read(&template).unwrap(), template_before);
    }
}

#[test]
fn native_merge_is_no_clobber_by_default_force_is_explicit_and_template_is_immutable() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let template = temp.path().join("template.docx");
    let output = temp.path().join("output.docx");
    create(&provider, &template);
    populate_template(&provider, &template);
    let template_before = std::fs::read(&template).unwrap();

    run(
        &provider,
        &[
            "merge",
            template.to_str().unwrap(),
            output.to_str().unwrap(),
            "--data",
            r#"{"user":{"name":"first"}}"#,
            "--json",
        ],
    );
    let output_before = std::fs::read(&output).unwrap();

    let exists = run_failure(
        &provider,
        &[
            "merge",
            template.to_str().unwrap(),
            output.to_str().unwrap(),
            "--data",
            r#"{"user":{"name":"second"}}"#,
            "--json",
        ],
    );
    assert_eq!(exists["error"]["code"], "use.office.package_exists");
    assert_eq!(std::fs::read(&output).unwrap(), output_before);

    let forced = run(
        &provider,
        &[
            "merge",
            template.to_str().unwrap(),
            output.to_str().unwrap(),
            "--data",
            r#"{"user":{"name":"second"}}"#,
            "--force",
            "--json",
        ],
    );
    assert_eq!(forced["data"]["force"], true);
    assert_eq!(text_view(&provider, &output), "Hello second");
    assert_eq!(std::fs::read(&template).unwrap(), template_before);

    let same = run_failure(
        &provider,
        &[
            "merge",
            template.to_str().unwrap(),
            template.to_str().unwrap(),
            "--data",
            r#"{"user":{"name":"destroy"}}"#,
            "--force",
            "--json",
        ],
    );
    assert_eq!(
        same["error"]["code"],
        "use.office.template_output_same_file"
    );
    assert_eq!(std::fs::read(&template).unwrap(), template_before);

    #[cfg(unix)]
    {
        let hard_link = temp.path().join("same-inode.docx");
        std::fs::hard_link(&template, &hard_link).unwrap();
        let same = run_failure(
            &provider,
            &[
                "merge",
                template.to_str().unwrap(),
                hard_link.to_str().unwrap(),
                "--data",
                r#"{"user":{"name":"destroy"}}"#,
                "--force",
                "--json",
            ],
        );
        assert_eq!(
            same["error"]["code"],
            "use.office.template_output_same_file"
        );
        assert_eq!(std::fs::read(&template).unwrap(), template_before);
    }
}

#[test]
fn native_merge_rejects_invalid_values_before_creating_output() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let template = temp.path().join("template.docx");
    let output = temp.path().join("output.docx");
    let data = temp.path().join("invalid.json");
    create(&provider, &template);
    populate_template(&provider, &template);
    std::fs::write(&data, br#"{"user.name":"bad\u0000value"}"#).unwrap();
    let template_before = std::fs::read(&template).unwrap();

    let invalid = run_failure(
        &provider,
        &[
            "merge",
            template.to_str().unwrap(),
            output.to_str().unwrap(),
            "--data",
            &format!("@{}", data.display()),
            "--json",
        ],
    );

    assert_eq!(
        invalid["error"]["code"],
        "use.office.template_value_invalid"
    );
    assert!(!output.exists());
    assert_eq!(std::fs::read(&template).unwrap(), template_before);
}
