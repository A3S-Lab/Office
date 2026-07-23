use std::collections::HashMap;
use std::ffi::OsString;
use std::path::{Path, PathBuf};
use std::process::Stdio;
use std::sync::Arc;
use std::time::Duration;

use a3s_use_core::{RiskClass, UseResult, UseSessionId};
use async_trait::async_trait;
use sha2::{Digest, Sha256};
use tokio::io::AsyncWriteExt;
use tokio::sync::RwLock;

use crate::discovery::{discover_office_cli, office_error};
use crate::{
    outcome_unknown, BatchRequest, OfficeProvider, OpenDocument, OperationResult, ReadRequest,
};

const DEFAULT_COMMAND_TIMEOUT: Duration = Duration::from_secs(120);

#[derive(Debug, Clone)]
pub struct OfficeCliProvider {
    executable: PathBuf,
    timeout: Duration,
    sessions: Arc<RwLock<HashMap<UseSessionId, OpenDocument>>>,
}

impl OfficeCliProvider {
    pub fn discover() -> UseResult<Self> {
        let executable = discover_office_cli().ok_or_else(|| {
            office_error(
                "use.office.runtime_missing",
                "The supported OfficeCLI provider is not installed.",
            )
            .with_suggestion("Run 'a3s install use/office'.")
        })?;
        Ok(Self::new(executable))
    }

    pub fn new(executable: PathBuf) -> Self {
        Self {
            executable,
            timeout: DEFAULT_COMMAND_TIMEOUT,
            sessions: Arc::new(RwLock::new(HashMap::new())),
        }
    }

    pub fn with_timeout(mut self, timeout: Duration) -> Self {
        self.timeout = timeout;
        self
    }

    pub fn executable(&self) -> &Path {
        &self.executable
    }

    async fn document(&self, session: &UseSessionId) -> UseResult<OpenDocument> {
        self.sessions
            .read()
            .await
            .get(session)
            .cloned()
            .ok_or_else(|| {
                office_error(
                    "use.office.session_missing",
                    format!("Office session '{}' is not open.", session.as_str()),
                )
            })
    }

    async fn run_json(
        &self,
        args: Vec<OsString>,
        stdin: Option<Vec<u8>>,
        risk: RiskClass,
        request_id: Option<&str>,
    ) -> UseResult<serde_json::Value> {
        let output = run_captured(
            &self.executable,
            &args,
            stdin,
            self.timeout,
            risk,
            request_id,
        )
        .await?;
        if output.stdout.iter().all(u8::is_ascii_whitespace) {
            return Ok(serde_json::Value::Null);
        }
        serde_json::from_slice(&output.stdout).map_err(|error| {
            office_error(
                "use.office.output_invalid",
                format!("OfficeCLI returned invalid JSON: {error}"),
            )
            .with_detail(
                "stdout",
                String::from_utf8_lossy(&output.stdout).to_string(),
            )
        })
    }
}

#[async_trait]
impl OfficeProvider for OfficeCliProvider {
    async fn open(&self, request: OpenDocument) -> UseResult<UseSessionId> {
        request.validate()?;
        let path = absolute_document_path(&request.path)?;
        let session = session_id(&path)?;
        let args = vec![
            OsString::from("open"),
            path.as_os_str().to_owned(),
            OsString::from("--json"),
        ];
        self.run_json(args, None, RiskClass::Read, None).await?;
        let mut document = request;
        document.path = path;
        self.sessions
            .write()
            .await
            .insert(session.clone(), document);
        Ok(session)
    }

    async fn read(&self, request: ReadRequest) -> UseResult<serde_json::Value> {
        if request.selector.trim().is_empty() {
            return Err(office_error(
                "use.office.selector_empty",
                "Office selectors cannot be empty.",
            ));
        }
        let document = self.document(&request.session).await?;
        let args = vec![
            OsString::from("get"),
            document.path.as_os_str().to_owned(),
            OsString::from(request.selector),
            OsString::from("--json"),
        ];
        self.run_json(args, None, RiskClass::Read, None).await
    }

    async fn batch(&self, request: BatchRequest) -> UseResult<OperationResult> {
        if request.request_id.trim().is_empty() {
            return Err(office_error(
                "use.office.request_id_empty",
                "Office mutation request IDs cannot be empty.",
            ));
        }
        if request.commands.is_empty() {
            return Err(office_error(
                "use.office.batch_empty",
                "Office mutation batches cannot be empty.",
            ));
        }
        let document = self.document(&request.session).await?;
        if document.read_only {
            return Err(office_error(
                "use.office.read_only",
                format!(
                    "Office session '{}' is read-only.",
                    request.session.as_str()
                ),
            ));
        }
        let input = serde_json::to_vec(&request.commands).map_err(|error| {
            office_error(
                "use.office.batch_invalid",
                format!("Failed to encode OfficeCLI batch commands: {error}"),
            )
        })?;
        let mut args = vec![
            OsString::from("batch"),
            document.path.as_os_str().to_owned(),
            OsString::from("--input"),
            OsString::from("-"),
        ];
        if request.stop_on_error {
            args.push(OsString::from("--stop-on-error"));
        }
        args.push(OsString::from("--json"));
        let output = self
            .run_json(
                args,
                Some(input),
                RiskClass::Mutate,
                Some(&request.request_id),
            )
            .await?;
        Ok(OperationResult {
            request_id: request.request_id,
            output,
        })
    }

    async fn save(&self, session: UseSessionId) -> UseResult<()> {
        let document = self.document(&session).await?;
        let request_id = format!("save-{}", session.as_str());
        let args = vec![
            OsString::from("save"),
            document.path.as_os_str().to_owned(),
            OsString::from("--json"),
        ];
        self.run_json(args, None, RiskClass::Mutate, Some(&request_id))
            .await?;
        Ok(())
    }

    async fn close(&self, session: UseSessionId) -> UseResult<()> {
        let document = self.document(&session).await?;
        let request_id = format!("close-{}", session.as_str());
        let args = vec![
            OsString::from("close"),
            document.path.as_os_str().to_owned(),
            OsString::from("--json"),
        ];
        self.run_json(args, None, RiskClass::Mutate, Some(&request_id))
            .await?;
        self.sessions.write().await.remove(&session);
        Ok(())
    }
}

/// Delegate one native OfficeCLI invocation, preserving argv, stdio, and status.
pub async fn delegate_native(args: &[String]) -> UseResult<u8> {
    let executable = discover_office_cli().ok_or_else(|| {
        office_error(
            "use.office.runtime_missing",
            "The supported OfficeCLI provider is not installed.",
        )
        .with_suggestion("Run 'a3s install use/office'.")
    })?;
    let status = tokio::process::Command::new(&executable)
        .args(args)
        .env("OFFICECLI_SKIP_UPDATE", "1")
        .kill_on_drop(true)
        .status()
        .await
        .map_err(|error| {
            office_error(
                "use.office.launch_failed",
                format!(
                    "Failed to launch OfficeCLI '{}': {error}",
                    executable.display()
                ),
            )
        })?;
    Ok(status
        .code()
        .and_then(|code| u8::try_from(code).ok())
        .unwrap_or(1))
}

struct CapturedOutput {
    stdout: Vec<u8>,
}

async fn run_captured(
    executable: &Path,
    args: &[OsString],
    input: Option<Vec<u8>>,
    timeout: Duration,
    risk: RiskClass,
    request_id: Option<&str>,
) -> UseResult<CapturedOutput> {
    let mut command = tokio::process::Command::new(executable);
    command
        .args(args)
        .env("OFFICECLI_SKIP_UPDATE", "1")
        .stdin(if input.is_some() {
            Stdio::piped()
        } else {
            Stdio::null()
        })
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .kill_on_drop(true);
    let mut child = command.spawn().map_err(|error| {
        office_error(
            "use.office.launch_failed",
            format!(
                "Failed to launch OfficeCLI '{}': {error}",
                executable.display()
            ),
        )
    })?;

    let run = async move {
        if let Some(input) = input {
            let mut stdin = child.stdin.take().ok_or_else(|| {
                std::io::Error::new(
                    std::io::ErrorKind::BrokenPipe,
                    "OfficeCLI stdin unavailable",
                )
            })?;
            stdin.write_all(&input).await?;
            stdin.shutdown().await?;
        }
        child.wait_with_output().await
    };

    let output = match tokio::time::timeout(timeout, run).await {
        Ok(Ok(output)) => output,
        Ok(Err(error)) if risk == RiskClass::Mutate => {
            return Err(outcome_unknown(request_id.unwrap_or("unknown"))
                .with_detail("transportError", error.to_string()))
        }
        Ok(Err(error)) => {
            return Err(office_error(
                "use.office.transport_failed",
                format!("OfficeCLI communication failed: {error}"),
            ))
        }
        Err(_) if risk == RiskClass::Mutate => {
            return Err(outcome_unknown(request_id.unwrap_or("unknown"))
                .with_detail("timeoutMs", duration_millis(timeout)))
        }
        Err(_) => {
            return Err(office_error(
                "use.office.timeout",
                format!("OfficeCLI did not finish within {timeout:?}."),
            ))
        }
    };

    if !output.status.success() {
        return Err(office_error(
            "use.office.command_failed",
            format!(
                "OfficeCLI exited with status {}.",
                output
                    .status
                    .code()
                    .map_or_else(|| "signal".to_string(), |code| code.to_string())
            ),
        )
        .with_detail(
            "exitCode",
            output
                .status
                .code()
                .map_or(serde_json::Value::Null, Into::into),
        )
        .with_detail(
            "stdout",
            String::from_utf8_lossy(&output.stdout).to_string(),
        )
        .with_detail(
            "stderr",
            String::from_utf8_lossy(&output.stderr).to_string(),
        ));
    }
    Ok(CapturedOutput {
        stdout: output.stdout,
    })
}

fn session_id(path: &Path) -> UseResult<UseSessionId> {
    let digest = Sha256::digest(path.as_os_str().as_encoded_bytes());
    let suffix = digest[..8]
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    UseSessionId::parse(format!("office_{suffix}"))
}

fn absolute_document_path(path: &Path) -> UseResult<PathBuf> {
    if path.is_absolute() {
        return Ok(path.to_path_buf());
    }
    std::env::current_dir()
        .map(|current| current.join(path))
        .map_err(|error| {
            office_error(
                "use.office.path_resolution_failed",
                format!("Failed to resolve Office document path: {error}"),
            )
        })
}

fn duration_millis(duration: Duration) -> u64 {
    duration.as_millis().try_into().unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{BatchCommand, DocumentKind};

    #[cfg(unix)]
    fn fixture(script: &str) -> (tempfile::TempDir, PathBuf) {
        use std::os::unix::fs::PermissionsExt;

        let temp = tempfile::tempdir().unwrap();
        let executable = temp.path().join("officecli");
        std::fs::write(&executable, script).unwrap();
        let mut permissions = std::fs::metadata(&executable).unwrap().permissions();
        permissions.set_mode(0o755);
        std::fs::set_permissions(&executable, permissions).unwrap();
        (temp, executable)
    }

    #[cfg(unix)]
    #[tokio::test]
    async fn provider_uses_native_officecli_commands_without_a_pipe_protocol() {
        let (temp, executable) = fixture(
            "#!/bin/sh\ncase \"$1\" in\nopen|close|save) printf '{\"success\":true}' ;;\nget) printf '{\"success\":true,\"data\":{\"text\":\"hello\"}}' ;;\nbatch) input=$(cat); printf '{\"success\":true,\"data\":%s}' \"$input\" ;;\nesac\n",
        );
        let path = temp.path().join("report.docx");
        std::fs::write(&path, b"fixture").unwrap();
        let provider = OfficeCliProvider::new(executable);
        let session = provider
            .open(OpenDocument {
                path,
                kind: DocumentKind::Word,
                read_only: false,
            })
            .await
            .unwrap();

        let read = provider
            .read(ReadRequest {
                session: session.clone(),
                selector: "/body".to_string(),
            })
            .await
            .unwrap();
        assert_eq!(read["data"]["text"], "hello");

        let result = provider
            .batch(BatchRequest {
                session: session.clone(),
                request_id: "mutation-1".to_string(),
                commands: vec![BatchCommand::new("set")],
                stop_on_error: true,
            })
            .await
            .unwrap();
        assert_eq!(result.request_id, "mutation-1");
        assert_eq!(result.output["data"][0]["command"], "set");
        provider.close(session).await.unwrap();
    }

    #[cfg(unix)]
    #[tokio::test]
    async fn timed_out_mutation_is_never_reported_as_retryable_failure() {
        let (temp, executable) = fixture(
            "#!/bin/sh\ncase \"$1\" in\nopen) printf '{\"success\":true}' ;;\nbatch) sleep 2; printf '{\"success\":true}' ;;\nesac\n",
        );
        let path = temp.path().join("report.xlsx");
        std::fs::write(&path, b"fixture").unwrap();
        let provider = OfficeCliProvider::new(executable);
        let session = provider
            .open(OpenDocument {
                path,
                kind: DocumentKind::Spreadsheet,
                read_only: false,
            })
            .await
            .unwrap();
        let provider = provider.with_timeout(Duration::from_millis(500));
        let error = provider
            .batch(BatchRequest {
                session,
                request_id: "mutation-timeout".to_string(),
                commands: vec![BatchCommand::new("set")],
                stop_on_error: true,
            })
            .await
            .unwrap_err();
        assert_eq!(error.code, "use.office.outcome_unknown");
        assert_eq!(error.details["requestId"], "mutation-timeout");
    }

    #[test]
    fn provider_is_send_and_sync() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<OfficeCliProvider>();
    }
}
