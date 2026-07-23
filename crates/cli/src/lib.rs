//! Standalone command-line and preview surfaces for A3S Office.

mod office_artifact;
mod office_mcp;
mod office_native_cli;
#[cfg(feature = "browser")]
pub mod office_screenshot;
mod office_skills;
pub mod office_watch;

use a3s_use_core::UseResult;

#[derive(Debug, Clone)]
pub struct CommandOutput {
    pub human: String,
    pub json: serde_json::Value,
    pub exit_code: u8,
    pub should_print: bool,
}

impl CommandOutput {
    pub(crate) fn success(human: impl Into<String>, data: serde_json::Value) -> Self {
        Self {
            human: human.into(),
            json: serde_json::json!({
                "schemaVersion": 1,
                "ok": true,
                "data": data,
            }),
            exit_code: 0,
            should_print: true,
        }
    }

    pub(crate) fn silent() -> Self {
        Self {
            human: String::new(),
            json: serde_json::Value::Null,
            exit_code: 0,
            should_print: false,
        }
    }
}

pub async fn run(args: &[String]) -> UseResult<CommandOutput> {
    match args.first().map(String::as_str) {
        None | Some("help" | "--help" | "-h") => Ok(help()),
        Some("-V" | "--version" | "version") => Ok(version()),
        Some("mcp") => {
            office_mcp::serve_stdio().await?;
            Ok(CommandOutput::silent())
        }
        Some("native") => office_native_cli::run(&args[1..]).await,
        Some("skills") => office_skills::run(&args[1..]).await,
        Some(_) => office_native_cli::run(args).await,
    }
}

pub async fn run_native(args: &[String]) -> UseResult<CommandOutput> {
    office_native_cli::run(args).await
}

pub async fn run_skills(args: &[String]) -> UseResult<CommandOutput> {
    office_skills::run(args).await
}

pub async fn primary_skill_surface() -> Option<(std::path::PathBuf, std::path::PathBuf)> {
    office_skills::primary_skill_surface().await
}

async fn ensure_office_compatibility_ready() -> UseResult<a3s_office::OfficeRuntimeStatus> {
    use a3s_office::OfficeInstallSource;
    use a3s_use_core::{FirstUseInstallPolicy, UseError};

    let status = a3s_office::office_status();
    let explicit_configured =
        std::env::var_os("A3S_OFFICECLI_EXECUTABLE").is_some_and(|value| !value.is_empty());
    if explicit_configured
        && !(status.available && status.source == OfficeInstallSource::Environment)
    {
        return Err(UseError::new(
            "use.office.explicit_provider_invalid",
            "The explicit Office provider is not usable.",
        )
        .with_suggestion("Fix or unset A3S_OFFICECLI_EXECUTABLE before retrying."));
    }
    if status.available {
        return Ok(status);
    }
    if let Some(block) = FirstUseInstallPolicy::from_env()?.blocked_by() {
        return Err(UseError::new(
            "use.office.auto_install_disabled",
            format!(
                "The optional OfficeCLI compatibility provider is not ready and first-use installation is disabled by {}.",
                block.reason()
            ),
        )
        .with_suggestion(
            "Enable first-use installation or prepare OfficeCLI explicitly while online.",
        )
        .with_detail("reason", block.reason()));
    }
    a3s_office::install_office_cli(false).await
}

fn help() -> CommandOutput {
    CommandOutput::success(
        concat!(
            "a3s-office — native Office document tools\n\n",
            "usage:\n",
            "  a3s-office <command> [args]\n",
            "  a3s-office native <command> [args]\n",
            "  a3s-office skills list|get|path [args]\n",
            "  a3s-office mcp\n\n",
            "Run 'a3s-office native --help' for document commands."
        ),
        serde_json::json!({
            "commands": ["native", "skills", "mcp"],
            "formats": ["docx", "xlsx", "pptx"]
        }),
    )
}

fn version() -> CommandOutput {
    CommandOutput::success(
        env!("CARGO_PKG_VERSION"),
        serde_json::json!({ "version": env!("CARGO_PKG_VERSION") }),
    )
}
