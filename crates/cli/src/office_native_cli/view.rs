use std::path::Path;

use a3s_office::{
    NativeOfficeAnnotatedOptions, NativeOfficeDocument, NativeOfficeIssueFilter,
    NativeOfficeIssueOptions, NativeOfficeIssueReport, NativeOfficeIssueSeverity,
    NativeOfficeRenderedView, DEFAULT_NATIVE_OFFICE_ANNOTATED_LIMIT,
    DEFAULT_NATIVE_OFFICE_ISSUE_LIMIT,
};
use a3s_use_core::UseResult;

use crate::office_artifact::{self, OfficeArtifactKind};

use super::{single_line, usage_error, AllowedOptions, CommandOutput, ParsedArguments};

pub(super) async fn run(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedArguments::parse(args, AllowedOptions::VIEW)?;
    if parsed.positionals.len() != 2 {
        return Err(usage_error(
            "office native view requires <file> and text, annotated, outline, stats, issues, html, svg, or screenshot",
        ));
    }
    let document = NativeOfficeDocument::open(&parsed.positionals[0]).await?;
    match parsed.positionals[1].as_str() {
        "text" | "t" => {
            reject_artifact_options(&parsed, "text")?;
            let view = document.text_view();
            Ok(CommandOutput::success(
                view.text.clone(),
                serde_json::json!({ "view": "text", "result": view }),
            ))
        }
        "annotated" | "a" => {
            reject_output_and_timeout(&parsed, "annotated")?;
            if parsed.node_type.is_some() {
                return Err(usage_error(
                    "--type is available for issues views, not annotated",
                ));
            }
            let view = document.annotated(NativeOfficeAnnotatedOptions {
                limit: parsed
                    .limit
                    .unwrap_or(DEFAULT_NATIVE_OFFICE_ANNOTATED_LIMIT),
            })?;
            Ok(CommandOutput::success(
                view.text.clone(),
                serde_json::json!({ "view": "annotated", "result": view }),
            ))
        }
        "outline" | "o" => {
            reject_artifact_options(&parsed, "outline")?;
            let outline = document.outline();
            let human = outline
                .iter()
                .map(|entry| {
                    format!(
                        "{}{} ({}) {}",
                        "  ".repeat(entry.level),
                        entry.path,
                        entry.node_type.label(),
                        single_line(&entry.text)
                    )
                })
                .collect::<Vec<_>>()
                .join("\n");
            Ok(CommandOutput::success(
                human,
                serde_json::json!({ "view": "outline", "result": outline }),
            ))
        }
        "stats" | "s" => {
            reject_artifact_options(&parsed, "stats")?;
            let statistics = document.statistics();
            let human = format!(
                "{} nodes, {} words, {} characters",
                statistics.node_count, statistics.word_count, statistics.character_count
            );
            Ok(CommandOutput::success(
                human,
                serde_json::json!({ "view": "stats", "result": statistics }),
            ))
        }
        "issues" | "i" => {
            reject_output_and_timeout(&parsed, "issues")?;
            let filter = parsed
                .node_type
                .as_deref()
                .map(NativeOfficeIssueFilter::parse)
                .transpose()?;
            let report = document.issues(NativeOfficeIssueOptions {
                filter,
                limit: parsed.limit.unwrap_or(DEFAULT_NATIVE_OFFICE_ISSUE_LIMIT),
            })?;
            Ok(CommandOutput::success(
                format_issue_report(&report),
                serde_json::json!({ "view": "issues", "result": report }),
            ))
        }
        "html" | "h" => {
            reject_timeout(&parsed, "html")?;
            reject_issue_options(&parsed, "html")?;
            rendered(document.html_view()?, parsed.output.as_deref()).await
        }
        "svg" => {
            reject_timeout(&parsed, "svg")?;
            reject_issue_options(&parsed, "svg")?;
            rendered(document.svg_view()?, parsed.output.as_deref()).await
        }
        "screenshot" | "png" => {
            reject_issue_options(&parsed, "screenshot")?;
            screenshot(document, parsed.output.as_deref(), parsed.timeout_ms).await
        }
        mode => Err(usage_error(format!(
            "native Office view mode '{mode}' is not text, annotated, outline, stats, issues, html, svg, or screenshot"
        ))),
    }
}

fn reject_artifact_options(parsed: &ParsedArguments, view: &str) -> UseResult<()> {
    reject_output_and_timeout(parsed, view)?;
    reject_issue_options(parsed, view)
}

fn reject_output_and_timeout(parsed: &ParsedArguments, view: &str) -> UseResult<()> {
    if parsed.output.is_some() {
        return Err(usage_error(format!(
            "--output is available for html, svg, and screenshot views, not {view}"
        )));
    }
    reject_timeout(parsed, view)
}

fn reject_timeout(parsed: &ParsedArguments, view: &str) -> UseResult<()> {
    if parsed.timeout_ms.is_none() {
        return Ok(());
    }
    Err(usage_error(format!(
        "--timeout-ms is available for screenshot views, not {view}"
    )))
}

fn reject_issue_options(parsed: &ParsedArguments, view: &str) -> UseResult<()> {
    if parsed.node_type.is_none() && parsed.limit.is_none() {
        return Ok(());
    }
    Err(usage_error(format!(
        "--type is available for issues views and --limit is available for annotated or issues views, not {view}"
    )))
}

fn format_issue_report(report: &NativeOfficeIssueReport) -> String {
    let mut lines = vec![format!(
        "Found {} issue(s); returned {}{}.",
        report.count,
        report.returned,
        if report.truncated { " (truncated)" } else { "" }
    )];
    for issue in &report.issues {
        let severity = match issue.severity {
            NativeOfficeIssueSeverity::Error => "ERROR",
            NativeOfficeIssueSeverity::Warning => "WARN",
            NativeOfficeIssueSeverity::Info => "INFO",
        };
        lines.push(format!(
            "[{severity}] {} {}: {}",
            issue.subtype.as_str(),
            issue.path,
            issue.message
        ));
        if let Some(suggestion) = &issue.suggestion {
            lines.push(format!("  Suggestion: {suggestion}"));
        }
    }
    lines.join("\n")
}

async fn rendered(
    view: NativeOfficeRenderedView,
    output: Option<&str>,
) -> UseResult<CommandOutput> {
    if let Some(path) = output.filter(|path| *path != "-") {
        office_artifact::write_new(
            Path::new(path),
            view.content.as_bytes().to_vec(),
            OfficeArtifactKind::SemanticRender,
        )
        .await?;
        return Ok(CommandOutput::success(
            format!(
                "Wrote native Office {:?} semantic preview to '{}'.",
                view.format, path
            ),
            serde_json::json!({
                "view": view.format,
                "result": {
                    "kind": view.kind,
                    "format": view.format,
                    "mediaType": view.media_type,
                    "unitCount": view.unit_count,
                    "byteLength": view.byte_length,
                    "sha256": view.sha256,
                    "outputPath": path
                }
            }),
        ));
    }
    Ok(CommandOutput::success(
        view.content.clone(),
        serde_json::json!({ "view": view.format, "result": view }),
    ))
}

#[cfg(feature = "browser")]
async fn screenshot(
    document: NativeOfficeDocument,
    output: Option<&str>,
    timeout_ms: Option<u64>,
) -> UseResult<CommandOutput> {
    use crate::office_screenshot::{
        capture_native_office_screenshot, NativeOfficeScreenshotRequest,
        DEFAULT_NATIVE_OFFICE_SCREENSHOT_TIMEOUT_MS,
    };

    let output = output
        .filter(|output| *output != "-")
        .ok_or_else(|| usage_error("screenshot view requires --output <file.png>"))?;
    let request = NativeOfficeScreenshotRequest::new(output)
        .with_timeout_ms(timeout_ms.unwrap_or(DEFAULT_NATIVE_OFFICE_SCREENSHOT_TIMEOUT_MS));
    let screenshot = capture_native_office_screenshot(&document, request).await?;
    Ok(CommandOutput::success(
        format!(
            "Wrote native Office screenshot to '{}'.",
            screenshot.output_path.display()
        ),
        serde_json::json!({ "view": "screenshot", "result": screenshot }),
    ))
}

#[cfg(not(feature = "browser"))]
async fn screenshot(
    _document: NativeOfficeDocument,
    _output: Option<&str>,
    _timeout_ms: Option<u64>,
) -> UseResult<CommandOutput> {
    Err(a3s_use_core::UseError::new(
        "use.browser.disabled",
        "Native Office screenshots require the A3S Use Browser feature.",
    )
    .with_suggestion("Use an A3S Use build with Browser support, or request html or svg instead."))
}
