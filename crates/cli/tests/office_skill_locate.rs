use std::path::Path;

#[test]
fn skill_locate_then_patch_forbids_whole_document_overwrite() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("skills/a3s-office");
    let files = [
        (
            "SKILL.md",
            &[
                "collab read",
                "office_collaboration_read",
                "startUtf16",
                "paragraphId",
                "sheetId",
                "elementId",
                "fieldId",
            ][..],
        ),
        (
            "references/word.md",
            &["paragraphId", "textId", "startUtf16", "endUtf16"][..],
        ),
        (
            "references/spreadsheet.md",
            &["sheetId", "row", "column"][..],
        ),
        (
            "references/presentation.md",
            &["containerKind", "containerId", "elementId"][..],
        ),
        (
            "references/markdown.md",
            &["startUtf16", "endUtf16", "expectedText", "markdown-splice"][..],
        ),
        (
            "references/pdf.md",
            &["pageIndex", "annotationId", "fieldId"][..],
        ),
        (
            "references/mcp.md",
            &[
                "office_collaboration_read",
                "office_collaboration_mutate",
                "startUtf16",
                "paragraphId",
                "sheetId",
                "elementId",
                "fieldId",
            ][..],
        ),
    ];
    for (relative, addresses) in files {
        let path = root.join(relative);
        let text = std::fs::read_to_string(&path)
            .unwrap_or_else(|error| panic!("read {}: {error}", path.display()));
        assert!(
            text.contains("Read the stable edit address"),
            "{} does not tell agents to read a stable edit address",
            path.display()
        );
        assert!(
            text.contains("Do not overwrite the whole document"),
            "{} does not forbid a whole-document overwrite",
            path.display()
        );
        for address in addresses {
            assert!(
                text.contains(address),
                "{} is missing live-edit address `{address}`",
                path.display()
            );
        }
    }
}
