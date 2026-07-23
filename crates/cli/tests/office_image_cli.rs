use std::path::Path;
use std::process::{Command, Output};

const PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
];
const PNG_1X1_BASE64: &str =
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACklEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";

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

#[test]
fn native_cli_adds_reads_and_removes_images_without_officecli() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let image = temp.path().join("logo.png");
    std::fs::write(&image, PNG_1X1).unwrap();

    for (extension, parent) in [
        ("docx", "/body"),
        ("xlsx", "/Sheet1/A1"),
        ("pptx", "/slide[1]"),
    ] {
        let document = temp.path().join(format!("image.{extension}"));
        create(&provider, &document);
        if extension == "pptx" {
            run(
                &provider,
                &[
                    "add",
                    document.to_str().unwrap(),
                    "/",
                    "--type",
                    "slide",
                    "--json",
                ],
            );
        }
        let output = execute(
            &provider,
            &[
                "add",
                document.to_str().unwrap(),
                parent,
                "--type",
                "picture",
                "--input",
                image.to_str().unwrap(),
                "--name",
                "A3S Logo",
                "--alt",
                "Native image test",
                "--width",
                "96",
                "--height",
                "48",
                "--json",
            ],
        );
        assert!(output.status.success(), "{output:?}");
        let stdout = String::from_utf8(output.stdout).unwrap();
        assert!(!stdout.contains(PNG_1X1_BASE64));
        let added: serde_json::Value = serde_json::from_str(&stdout).unwrap();
        assert_eq!(added["data"]["operation"], "add-image");
        assert_eq!(added["data"]["node"]["type"], "picture");
        assert_eq!(added["data"]["node"]["format"]["name"], "A3S Logo");
        assert_eq!(added["data"]["node"]["format"]["alt"], "Native image test");
        assert_eq!(added["data"]["createdImage"]["format"], "png");
        assert_eq!(added["data"]["createdImage"]["widthPx"], 96);
        assert_eq!(added["data"]["createdImage"]["heightPx"], 48);
        assert!(added["data"]["createdImage"].get("data").is_none());
        let path = added["data"]["path"].as_str().unwrap();

        let stats = run(
            &provider,
            &["view", document.to_str().unwrap(), "stats", "--json"],
        );
        assert_eq!(stats["data"]["result"]["pictureCount"], 1);

        run(
            &provider,
            &["remove", document.to_str().unwrap(), path, "--json"],
        );
        let stats = run(
            &provider,
            &["view", document.to_str().unwrap(), "stats", "--json"],
        );
        assert_eq!(stats["data"]["result"]["pictureCount"], 0);
    }
}

#[test]
fn native_cli_validates_image_inputs_and_serializes_batch_receipts() {
    let temp = tempfile::tempdir().unwrap();
    let provider = temp.path().join("must-not-be-invoked");
    let document = temp.path().join("batch.docx");
    create(&provider, &document);

    let directory = temp.path().join("not-an-image-file");
    std::fs::create_dir(&directory).unwrap();
    let failure = run_failure(
        &provider,
        &[
            "add",
            document.to_str().unwrap(),
            "/body",
            "--type",
            "picture",
            "--input",
            directory.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(failure["error"]["code"], "use.office.image_input_invalid");

    let invalid = temp.path().join("invalid.png");
    std::fs::write(&invalid, b"not an image").unwrap();
    let before = std::fs::read(&document).unwrap();
    let failure = run_failure(
        &provider,
        &[
            "add",
            document.to_str().unwrap(),
            "/body",
            "--type",
            "picture",
            "--input",
            invalid.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(failure["error"]["code"], "use.office.image_invalid");
    assert_eq!(std::fs::read(&document).unwrap(), before);

    let batch = temp.path().join("image-batch.json");
    std::fs::write(
        &batch,
        serde_json::to_vec(&serde_json::json!({
            "schemaVersion": 1,
            "mutations": [{
                "operation": "add-image",
                "parent": "/body",
                "image": {
                    "format": "png",
                    "data": PNG_1X1_BASE64,
                    "name": "Batch Logo",
                    "widthPx": 32
                }
            }]
        }))
        .unwrap(),
    )
    .unwrap();
    let result = run(
        &provider,
        &[
            "batch",
            document.to_str().unwrap(),
            "--input",
            batch.to_str().unwrap(),
            "--json",
        ],
    );
    assert_eq!(result["data"]["result"]["applied"], 1);
    assert_eq!(
        result["data"]["result"]["createdImages"][0]["format"],
        "png"
    );
    assert_eq!(result["data"]["result"]["createdImages"][0]["widthPx"], 32);
    assert_eq!(result["data"]["result"]["createdImages"][0]["heightPx"], 32);
    assert!(result["data"]["result"]["createdImages"][0]
        .get("data")
        .is_none());
}
