use std::path::{Path, PathBuf};

use a3s_use_core::{DomainDiagnostic, Readiness, UseError, UseResult};
use serde::{Deserialize, Serialize};

pub const SUPPORTED_OFFICECLI_VERSION: &str = "1.0.136";
pub(crate) const RECEIPT_FILE: &str = ".a3s-install.json";

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum OfficeInstallSource {
    Environment,
    System,
    Managed,
    Missing,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OfficeRuntimeStatus {
    pub available: bool,
    pub source: OfficeInstallSource,
    pub path: Option<PathBuf>,
    pub version: Option<String>,
    pub managed_root: Option<PathBuf>,
    pub detail: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OfficeInstallReceipt {
    pub schema_version: u32,
    pub provider: String,
    pub version: String,
    pub source_url: String,
    pub artifact_sha256: String,
    pub artifact_bytes: u64,
}

pub fn office_status() -> OfficeRuntimeStatus {
    let managed_root = managed_root().ok();
    if let Some(path) = env_executable("A3S_OFFICECLI_EXECUTABLE") {
        return ready(OfficeInstallSource::Environment, path, None, managed_root);
    }
    if let Some(path) = path_executable() {
        return ready(OfficeInstallSource::System, path, None, managed_root);
    }
    if let Some((path, receipt)) = managed_executable() {
        return ready(
            OfficeInstallSource::Managed,
            path,
            Some(receipt.version),
            managed_root,
        );
    }
    OfficeRuntimeStatus {
        available: false,
        source: OfficeInstallSource::Missing,
        path: None,
        version: None,
        managed_root,
        detail: "The supported OfficeCLI provider is not installed.".to_string(),
    }
}

pub fn doctor() -> DomainDiagnostic {
    let status = office_status();
    if status.available {
        DomainDiagnostic {
            domain: "office".to_string(),
            readiness: Readiness::Ready,
            provider: Some("officecli".to_string()),
            version: status.version,
            path: status.path,
            message: "OfficeCLI is available through its native CLI surface.".to_string(),
            suggestions: Vec::new(),
        }
    } else {
        DomainDiagnostic {
            domain: "office".to_string(),
            readiness: Readiness::Missing,
            provider: None,
            version: None,
            path: None,
            message: status.detail,
            suggestions: vec!["Run 'a3s install use/office'.".to_string()],
        }
    }
}

pub fn discover_office_cli() -> Option<PathBuf> {
    office_status().path
}

pub(crate) fn managed_root() -> UseResult<PathBuf> {
    if let Some(value) = std::env::var_os("A3S_USE_OFFICE_HOME") {
        return absolute(PathBuf::from(value));
    }
    if let Some(value) = std::env::var_os("A3S_DATA_HOME") {
        return Ok(absolute(PathBuf::from(value))?.join("use/office"));
    }
    if let Some(value) = std::env::var_os("XDG_DATA_HOME") {
        return Ok(absolute(PathBuf::from(value))?.join("a3s/use/office"));
    }
    if let Some(home) = std::env::var_os("HOME").map(PathBuf::from) {
        return Ok(absolute(home)?.join(".local/share/a3s/use/office"));
    }
    #[cfg(windows)]
    if let Some(value) = std::env::var_os("LOCALAPPDATA") {
        return Ok(absolute(PathBuf::from(value))?.join("a3s/use/office"));
    }
    Err(office_error(
        "use.office.data_home_missing",
        "Cannot determine the A3S Office data directory.",
    ))
}

pub(crate) fn managed_version_dir() -> UseResult<PathBuf> {
    Ok(managed_root()?.join(SUPPORTED_OFFICECLI_VERSION))
}

pub(crate) fn managed_binary_name() -> &'static str {
    if cfg!(windows) {
        "officecli.exe"
    } else {
        "officecli"
    }
}

pub(crate) fn executable(path: &Path) -> bool {
    if !path.is_file() {
        return false;
    }
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        std::fs::metadata(path)
            .map(|metadata| metadata.permissions().mode() & 0o111 != 0)
            .unwrap_or(false)
    }
    #[cfg(not(unix))]
    {
        true
    }
}

pub(crate) fn office_error(code: &str, message: impl Into<String>) -> UseError {
    UseError::new(code, message)
}

fn env_executable(name: &str) -> Option<PathBuf> {
    std::env::var_os(name)
        .map(PathBuf::from)
        .filter(|path| executable(path))
}

fn path_executable() -> Option<PathBuf> {
    let path = std::env::var_os("PATH")?;
    for directory in std::env::split_paths(&path) {
        for name in if cfg!(windows) {
            ["officecli.exe", "office-cli.exe"]
        } else {
            ["officecli", "office-cli"]
        } {
            let candidate = directory.join(name);
            if executable(&candidate) {
                return Some(candidate);
            }
        }
    }
    None
}

fn managed_executable() -> Option<(PathBuf, OfficeInstallReceipt)> {
    let version_dir = managed_version_dir().ok()?;
    let path = version_dir.join(managed_binary_name());
    if !executable(&path) {
        return None;
    }
    let bytes = std::fs::read(version_dir.join(RECEIPT_FILE)).ok()?;
    let receipt: OfficeInstallReceipt = serde_json::from_slice(&bytes).ok()?;
    if receipt.schema_version != 1
        || receipt.provider != "officecli"
        || receipt.version != SUPPORTED_OFFICECLI_VERSION
    {
        return None;
    }
    Some((path, receipt))
}

fn ready(
    source: OfficeInstallSource,
    path: PathBuf,
    version: Option<String>,
    managed_root: Option<PathBuf>,
) -> OfficeRuntimeStatus {
    OfficeRuntimeStatus {
        available: true,
        source,
        path: Some(path),
        version,
        managed_root,
        detail: "ready".to_string(),
    }
}

fn absolute(path: PathBuf) -> UseResult<PathBuf> {
    if path.is_absolute() {
        return Ok(path);
    }
    std::env::current_dir()
        .map(|current| current.join(path))
        .map_err(|error| {
            office_error(
                "use.office.path_resolution_failed",
                format!("Failed to resolve Office data path: {error}"),
            )
        })
}
