//! Explicit preview MCP surface for the in-process native Office engine.

mod input;
mod session;
mod support;
#[cfg(test)]
mod tests;

use std::path::Path;

use a3s_office::{
    NativeOfficeAnnotatedOptions, NativeOfficeDocument, NativeOfficeIssueOptions,
    NativeOfficeRenderFormat, DEFAULT_NATIVE_OFFICE_ANNOTATED_LIMIT,
    DEFAULT_NATIVE_OFFICE_ISSUE_LIMIT,
};
use a3s_use_core::{UseError, UseResult};
use input::{
    OfficeBatchInput, OfficeCloseInput, OfficeCreateInput, OfficeFileInput, OfficeGetInput,
    OfficeMergeTemplateInput, OfficeOpenInput, OfficeQueryInput, OfficeRawXmlInput,
    OfficeSaveInput, OfficeView, OfficeViewInput,
};
use rmcp::handler::server::{router::tool::ToolRouter, wrapper::Parameters};
use rmcp::model::{CallToolResult, Implementation, ServerCapabilities, ServerInfo};
use rmcp::{tool, tool_handler, tool_router, ServerHandler, ServiceExt};
use session::NativeOfficeSessions;
use support::{
    output_encoding_error, reject_same_file, session_value, tool_result, validate_batch,
    validate_json_bytes, MAX_BATCH_BYTES, MAX_RAW_XML_BYTES,
};

const DEFAULT_QUERY_LIMIT: usize = 200;
const MAX_QUERY_LIMIT: usize = 1_000;

#[cfg(feature = "browser")]
async fn screenshot_view(
    document: &NativeOfficeDocument,
    output: Option<String>,
    timeout_ms: Option<u64>,
) -> UseResult<serde_json::Value> {
    use crate::office_screenshot::{
        capture_native_office_screenshot, NativeOfficeScreenshotRequest,
        DEFAULT_NATIVE_OFFICE_SCREENSHOT_TIMEOUT_MS,
    };

    let output = output.ok_or_else(|| {
        UseError::new(
            "use.office.screenshot_output_invalid",
            "Native Office MCP screenshot view requires output.",
        )
    })?;
    let request = NativeOfficeScreenshotRequest::new(output)
        .with_timeout_ms(timeout_ms.unwrap_or(DEFAULT_NATIVE_OFFICE_SCREENSHOT_TIMEOUT_MS));
    let screenshot = capture_native_office_screenshot(document, request).await?;
    serde_json::to_value(screenshot).map_err(output_encoding_error)
}

#[cfg(not(feature = "browser"))]
async fn screenshot_view(
    _document: &NativeOfficeDocument,
    _output: Option<String>,
    _timeout_ms: Option<u64>,
) -> UseResult<serde_json::Value> {
    Err(UseError::new(
        "use.browser.disabled",
        "Native Office screenshots require the A3S Use Browser feature.",
    )
    .with_suggestion("Use an A3S Use build with Browser support, or request html instead."))
}

#[derive(Clone)]
struct NativeOfficeMcpServer {
    sessions: NativeOfficeSessions,
    tool_router: ToolRouter<Self>,
}

impl NativeOfficeMcpServer {
    fn new() -> Self {
        Self {
            sessions: NativeOfficeSessions::default(),
            tool_router: Self::tool_router(),
        }
    }
}

#[tool_router]
impl NativeOfficeMcpServer {
    #[tool(
        name = "office_install_compat",
        description = "Install or repair the optional pinned OfficeCLI compatibility provider through the bounded A3S component lifecycle",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = true,
            open_world_hint = true
        )
    )]
    async fn office_install_compat(&self) -> Result<CallToolResult, rmcp::ErrorData> {
        let result = async {
            let status = crate::ensure_office_compatibility_ready().await?;
            serde_json::to_value(status).map_err(output_encoding_error)
        }
        .await;
        Ok(tool_result(result))
    }

    #[tool(
        name = "office_validate",
        description = "Validate and identify one local OOXML document without opening a session",
        annotations(
            read_only_hint = true,
            destructive_hint = false,
            idempotent_hint = true,
            open_world_hint = false
        )
    )]
    async fn office_validate(
        &self,
        Parameters(input): Parameters<OfficeFileInput>,
    ) -> Result<CallToolResult, rmcp::ErrorData> {
        let result = async {
            let document = NativeOfficeDocument::open(&input.file).await?;
            Ok(serde_json::json!({
                "valid": true,
                "path": document.package().path(),
                "kind": document.kind(),
                "revision": document.package().source_revision(),
                "contentSha256": document.package().content_sha256()
            }))
        }
        .await;
        Ok(tool_result(result))
    }

    #[tool(
        name = "office_create",
        description = "Create a blank native OOXML document and register a mutable in-memory session",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = false
        )
    )]
    async fn office_create(
        &self,
        Parameters(input): Parameters<OfficeCreateInput>,
    ) -> Result<CallToolResult, rmcp::ErrorData> {
        let result = async {
            let (session, entry) = self.sessions.create(input.session, input.file).await?;
            let state = entry.lock().await;
            Ok(session_value(&session, &state))
        }
        .await;
        Ok(tool_result(result))
    }

    #[tool(
        name = "office_open",
        description = "Open a local OOXML document in a bounded native in-memory session",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = false
        )
    )]
    async fn office_open(
        &self,
        Parameters(input): Parameters<OfficeOpenInput>,
    ) -> Result<CallToolResult, rmcp::ErrorData> {
        let result = async {
            let (session, entry) = self
                .sessions
                .open_existing(input.session, input.file, input.read_only)
                .await?;
            let state = entry.lock().await;
            Ok(session_value(&session, &state))
        }
        .await;
        Ok(tool_result(result))
    }

    #[tool(
        name = "office_list",
        description = "List native Office sessions owned by this MCP server process",
        annotations(
            read_only_hint = true,
            destructive_hint = false,
            idempotent_hint = true,
            open_world_hint = false
        )
    )]
    async fn office_list(&self) -> Result<CallToolResult, rmcp::ErrorData> {
        let mut entries = self.sessions.list().await;
        entries.sort_by(|(left, _), (right, _)| left.as_str().cmp(right.as_str()));
        let mut sessions = Vec::with_capacity(entries.len());
        for (session, entry) in entries {
            let state = entry.lock().await;
            if state.ensure_open(&session).is_ok() {
                sessions.push(session_value(&session, &state));
            }
        }
        Ok(tool_result(Ok(serde_json::json!({
            "count": sessions.len(),
            "sessions": sessions
        }))))
    }

    #[tool(
        name = "office_get",
        description = "Read one stable semantic path from an open native Office session",
        annotations(
            read_only_hint = true,
            destructive_hint = false,
            idempotent_hint = true,
            open_world_hint = false
        )
    )]
    async fn office_get(
        &self,
        Parameters(input): Parameters<OfficeGetInput>,
    ) -> Result<CallToolResult, rmcp::ErrorData> {
        let result = async {
            let depth = input.depth.unwrap_or(1);
            if depth > 64 {
                return Err(UseError::new(
                    "use.office.depth_invalid",
                    "Native Office MCP get depth cannot exceed 64.",
                ));
            }
            let (session, entry) = self.sessions.get(&input.session).await?;
            let state = entry.lock().await;
            state.ensure_open(&session)?;
            let path = input.path.as_deref().unwrap_or("/");
            let node = state.editor.snapshot()?.get(path, depth)?;
            Ok(serde_json::json!({ "session": session, "node": node }))
        }
        .await;
        Ok(tool_result(result))
    }

    #[tool(
        name = "office_query",
        description = "Run a native semantic selector with a bounded result count",
        annotations(
            read_only_hint = true,
            destructive_hint = false,
            idempotent_hint = true,
            open_world_hint = false
        )
    )]
    async fn office_query(
        &self,
        Parameters(input): Parameters<OfficeQueryInput>,
    ) -> Result<CallToolResult, rmcp::ErrorData> {
        let result = async {
            let limit = input.limit.unwrap_or(DEFAULT_QUERY_LIMIT);
            if !(1..=MAX_QUERY_LIMIT).contains(&limit) {
                return Err(UseError::new(
                    "use.office.query_limit_invalid",
                    format!(
                        "Native Office MCP query limit must be between 1 and {MAX_QUERY_LIMIT}."
                    ),
                ));
            }
            let (session, entry) = self.sessions.get(&input.session).await?;
            let state = entry.lock().await;
            state.ensure_open(&session)?;
            let mut results = state.editor.snapshot()?.query(&input.selector)?;
            let matches = results.len();
            results.truncate(limit);
            Ok(serde_json::json!({
                "session": session,
                "matches": matches,
                "returned": results.len(),
                "truncated": matches > results.len(),
                "results": results
            }))
        }
        .await;
        Ok(tool_result(result))
    }

    #[tool(
        name = "office_view",
        description = "Produce a native text, bounded annotated, outline, statistics, bounded issues, standalone all-format HTML or SVG, or Browser-injected PNG screenshot view for an open session",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = false
        )
    )]
    async fn office_view(
        &self,
        Parameters(input): Parameters<OfficeViewInput>,
    ) -> Result<CallToolResult, rmcp::ErrorData> {
        let result = async {
            if input.view != OfficeView::Screenshot
                && (input.output.is_some() || input.timeout_ms.is_some())
            {
                return Err(UseError::new(
                    "use.office.view_options_invalid",
                    "Native Office MCP output and timeoutMs are available only for screenshot views.",
                ));
            }
            if input.view != OfficeView::Issues && input.issue_type.is_some() {
                return Err(UseError::new(
                    "use.office.view_options_invalid",
                    "Native Office MCP issueType is available only for issues views.",
                ));
            }
            if !matches!(input.view, OfficeView::Annotated | OfficeView::Issues)
                && input.limit.is_some()
            {
                return Err(UseError::new(
                    "use.office.view_options_invalid",
                    "Native Office MCP limit is available only for annotated or issues views.",
                ));
            }
            let (session, entry) = self.sessions.get(&input.session).await?;
            let state = entry.lock().await;
            state.ensure_open(&session)?;
            let document = state.editor.snapshot()?;
            drop(state);
            let (view, value) = match input.view {
                OfficeView::Text => ("text", serde_json::to_value(document.text_view())),
                OfficeView::Annotated => (
                    "annotated",
                    serde_json::to_value(document.annotated(NativeOfficeAnnotatedOptions {
                        limit: input
                            .limit
                            .unwrap_or(DEFAULT_NATIVE_OFFICE_ANNOTATED_LIMIT),
                    })?),
                ),
                OfficeView::Outline => ("outline", serde_json::to_value(document.outline())),
                OfficeView::Stats => ("stats", serde_json::to_value(document.statistics())),
                OfficeView::Issues => (
                    "issues",
                    serde_json::to_value(document.issues(NativeOfficeIssueOptions {
                        filter: input.issue_type.map(Into::into),
                        limit: input.limit.unwrap_or(DEFAULT_NATIVE_OFFICE_ISSUE_LIMIT),
                    })?),
                ),
                OfficeView::Html => (
                    "html",
                    serde_json::to_value(document.render(NativeOfficeRenderFormat::Html)?),
                ),
                OfficeView::Svg => (
                    "svg",
                    serde_json::to_value(document.render(NativeOfficeRenderFormat::Svg)?),
                ),
                OfficeView::Screenshot => (
                    "screenshot",
                    Ok(screenshot_view(&document, input.output, input.timeout_ms).await?),
                ),
            };
            let value = value.map_err(output_encoding_error)?;
            Ok(serde_json::json!({ "session": session, "view": view, "result": value }))
        }
        .await;
        Ok(tool_result(result))
    }

    #[tool(
        name = "office_raw_xml",
        description = "Inspect one existing OOXML XML part, limited to 1 MiB of original bytes",
        annotations(
            read_only_hint = true,
            destructive_hint = false,
            idempotent_hint = true,
            open_world_hint = false
        )
    )]
    async fn office_raw_xml(
        &self,
        Parameters(input): Parameters<OfficeRawXmlInput>,
    ) -> Result<CallToolResult, rmcp::ErrorData> {
        let result = async {
            let (session, entry) = self.sessions.get(&input.session).await?;
            let state = entry.lock().await;
            state.ensure_open(&session)?;
            let part = state.editor.raw_xml_part(&input.part)?;
            if part.byte_length > MAX_RAW_XML_BYTES {
                return Err(UseError::new(
                    "use.office.mcp_raw_result_too_large",
                    format!(
                        "OOXML part '{}' is {} bytes; native Office MCP raw output is limited to {MAX_RAW_XML_BYTES} bytes.",
                        part.part, part.byte_length
                    ),
                )
                .with_suggestion("Use 'a3s use office native raw ... --output <file>' for a larger local part."));
            }
            Ok(serde_json::json!({ "session": session, "result": part }))
        }
        .await;
        Ok(tool_result(result))
    }

    #[tool(
        name = "office_apply_batch",
        description = "Apply a bounded typed mutation batch atomically in memory; call office_save to persist it",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = false
        )
    )]
    async fn office_apply_batch(
        &self,
        Parameters(input): Parameters<OfficeBatchInput>,
    ) -> Result<CallToolResult, rmcp::ErrorData> {
        let result = async {
            validate_batch(&input.mutations)?;
            let mutations = input
                .mutations
                .into_iter()
                .map(|mutation| mutation.into_native())
                .collect::<UseResult<Vec<_>>>()?;
            let (session, entry) = self.sessions.get(&input.session).await?;
            let mut state = entry.lock().await;
            state.ensure_mutable(&session)?;
            let result = state.editor.apply_batch(&mutations)?;
            Ok(serde_json::json!({
                "session": session,
                "atomic": true,
                "persisted": false,
                "result": result,
                "document": session_value(&session, &state)
            }))
        }
        .await;
        Ok(tool_result(result))
    }

    #[tool(
        name = "office_merge_template",
        description = "Merge bounded JSON data into a cloned session document and atomically save a distinct output",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = false
        )
    )]
    async fn office_merge_template(
        &self,
        Parameters(input): Parameters<OfficeMergeTemplateInput>,
    ) -> Result<CallToolResult, rmcp::ErrorData> {
        let result = async {
            validate_json_bytes(&input.data, MAX_BATCH_BYTES, "template data")?;
            let (session, entry) = self.sessions.get(&input.session).await?;
            let (mut editor, template_path) = {
                let state = entry.lock().await;
                state.ensure_open(&session)?;
                (
                    state.editor.clone(),
                    state.editor.package().path().to_path_buf(),
                )
            };
            reject_same_file(&template_path, Path::new(&input.output)).await?;
            let merge = editor.merge_template(&input.data)?;
            if input.overwrite {
                editor.save_as(&input.output).await?;
            } else {
                editor.save_as_new(&input.output).await?;
            }
            Ok(serde_json::json!({
                "session": session,
                "templatePath": template_path,
                "outputPath": editor.package().path(),
                "overwrite": input.overwrite,
                "kind": editor.package().kind(),
                "revision": editor.package().source_revision(),
                "result": merge
            }))
        }
        .await;
        Ok(tool_result(result))
    }

    #[tool(
        name = "office_save",
        description = "Atomically persist one mutable native Office session, optionally to a new path",
        annotations(
            read_only_hint = false,
            destructive_hint = true,
            idempotent_hint = false,
            open_world_hint = false
        )
    )]
    async fn office_save(
        &self,
        Parameters(input): Parameters<OfficeSaveInput>,
    ) -> Result<CallToolResult, rmcp::ErrorData> {
        let result = async {
            let (session, entry) = self.sessions.get(&input.session).await?;
            let mut state = entry.lock().await;
            state.ensure_mutable(&session)?;
            match input.output {
                Some(output) if input.overwrite => state.editor.save_as(output).await?,
                Some(output) => state.editor.save_as_new(output).await?,
                None => state.editor.save().await?,
            }
            Ok(serde_json::json!({
                "session": session,
                "saved": true,
                "document": session_value(&session, &state)
            }))
        }
        .await;
        Ok(tool_result(result))
    }

    #[tool(
        name = "office_close",
        description = "Close a native Office session, refusing unsaved changes unless discard is explicit",
        annotations(
            read_only_hint = false,
            destructive_hint = true,
            idempotent_hint = true,
            open_world_hint = false
        )
    )]
    async fn office_close(
        &self,
        Parameters(input): Parameters<OfficeCloseInput>,
    ) -> Result<CallToolResult, rmcp::ErrorData> {
        let result = async {
            let (session, entry) = self.sessions.close(&input.session, input.discard).await?;
            let state = entry.lock().await;
            Ok(serde_json::json!({
                "session": session,
                "closed": true,
                "discarded": input.discard,
                "path": state.editor.package().path()
            }))
        }
        .await;
        Ok(tool_result(result))
    }
}

#[tool_handler]
impl ServerHandler for NativeOfficeMcpServer {
    fn get_info(&self) -> ServerInfo {
        ServerInfo {
            capabilities: ServerCapabilities::builder().enable_tools().build(),
            server_info: Implementation {
                name: "a3s-office-native".to_string(),
                title: Some("A3S Office".to_string()),
                version: env!("CARGO_PKG_VERSION").to_string(),
                icons: None,
                website_url: Some("https://github.com/A3S-Lab/Office".to_string()),
            },
            instructions: Some(
                "Use the built-in native Office tools first; they never require OfficeCLI, Microsoft Office, or LibreOffice. If a requested operation is outside the native surface, request office_install_compat through the host confirmation path and use the separately projected Office compatibility route after it becomes ready. Create or open a native session first. Mutations remain in memory until office_save; office_close refuses unsaved changes unless discard=true."
                    .to_string(),
            ),
            ..Default::default()
        }
    }
}

pub(crate) async fn serve_stdio() -> UseResult<()> {
    let service = NativeOfficeMcpServer::new()
        .serve(rmcp::transport::io::stdio())
        .await
        .map_err(|error| {
            UseError::new(
                "use.mcp.office_native_start_failed",
                format!("Failed to start the native Office MCP server: {error}"),
            )
        })?;
    service.waiting().await.map_err(|error| {
        UseError::new(
            "use.mcp.office_native_failed",
            format!("Native Office MCP server task failed: {error}"),
        )
    })?;
    Ok(())
}
