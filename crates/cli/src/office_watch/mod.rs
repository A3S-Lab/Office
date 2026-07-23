//! Authenticated loopback live previews for native Office documents.
//!
//! The watch server is a read-only facade over deterministic native HTML. It
//! never invokes OfficeCLI, LibreOffice, or the Browser runtime, and it does not
//! expose a mutation or custom RPC endpoint.

mod page;
mod routes;

use std::future::Future;
use std::net::{Ipv4Addr, SocketAddr};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::{Duration, SystemTime};

use a3s_office::{DocumentKind, NativeOfficeDocument, PackageRevision};
use a3s_use_core::{UseError, UseResult};
use axum::body::Bytes;
use serde::{Deserialize, Serialize};
use tokio::net::TcpListener;
use tokio::sync::{broadcast, RwLock};

pub const DEFAULT_NATIVE_OFFICE_WATCH_POLL_MS: u64 = 250;
pub const MIN_NATIVE_OFFICE_WATCH_POLL_MS: u64 = 50;
pub const MAX_NATIVE_OFFICE_WATCH_POLL_MS: u64 = 10_000;

const EVENT_BUFFER: usize = 32;
const MAX_RELOAD_RETRY: Duration = Duration::from_secs(5);

/// Network and polling settings for one read-only live preview.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeWatchOptions {
    /// Loopback TCP port. Zero asks the operating system for an ephemeral port.
    pub port: u16,
    /// Interval used to detect atomic saves and other on-disk changes.
    pub poll_interval_ms: u64,
}

impl Default for NativeOfficeWatchOptions {
    fn default() -> Self {
        Self {
            port: 0,
            poll_interval_ms: DEFAULT_NATIVE_OFFICE_WATCH_POLL_MS,
        }
    }
}

impl NativeOfficeWatchOptions {
    fn validate(self) -> UseResult<Self> {
        if !(MIN_NATIVE_OFFICE_WATCH_POLL_MS..=MAX_NATIVE_OFFICE_WATCH_POLL_MS)
            .contains(&self.poll_interval_ms)
        {
            return Err(UseError::new(
                "use.office.watch_poll_invalid",
                format!(
                    "Native Office watch polling must be between {MIN_NATIVE_OFFICE_WATCH_POLL_MS} and {MAX_NATIVE_OFFICE_WATCH_POLL_MS} milliseconds."
                ),
            ));
        }
        Ok(self)
    }

    fn poll_interval(self) -> Duration {
        Duration::from_millis(self.poll_interval_ms)
    }
}

/// Startup receipt for an authenticated live preview server.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeWatchReady {
    pub url: String,
    pub address: SocketAddr,
    pub document_name: String,
    pub kind: DocumentKind,
    pub revision: PackageRevision,
    pub render_sha256: String,
    pub version: u64,
    pub poll_interval_ms: u64,
}

/// Last reload failure while the previous valid preview remains available.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeWatchError {
    pub code: String,
    pub message: String,
}

/// Public state returned by the status endpoint and emitted over standard SSE.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeWatchStatus {
    pub healthy: bool,
    pub version: u64,
    pub kind: DocumentKind,
    pub revision: PackageRevision,
    pub render_sha256: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<NativeOfficeWatchError>,
}

#[derive(Debug, Clone)]
struct WatchSnapshot {
    html: Bytes,
    version: u64,
    revision: PackageRevision,
    render_sha256: String,
}

struct WatchState {
    address: SocketAddr,
    token: String,
    path: PathBuf,
    kind: DocumentKind,
    poll_interval: Duration,
    initial_stamp: Option<FileStamp>,
    snapshot: RwLock<WatchSnapshot>,
    error: RwLock<Option<NativeOfficeWatchError>>,
    events: broadcast::Sender<NativeOfficeWatchStatus>,
    shutdown: broadcast::Sender<()>,
}

impl WatchState {
    async fn status(&self) -> NativeOfficeWatchStatus {
        let snapshot = self.snapshot.read().await;
        let error = self.error.read().await.clone();
        NativeOfficeWatchStatus {
            healthy: error.is_none(),
            version: snapshot.version,
            kind: self.kind,
            revision: snapshot.revision.clone(),
            render_sha256: snapshot.render_sha256.clone(),
            error,
        }
    }

    async fn html(&self) -> Bytes {
        self.snapshot.read().await.html.clone()
    }

    async fn replace(&self, document: NativeOfficeDocument) -> UseResult<()> {
        if document.kind() != self.kind {
            return Err(UseError::new(
                "use.office.watch_kind_changed",
                "The watched file changed to a different Office document kind.",
            )
            .with_suggestion("Stop the watch and open the replacement document explicitly."));
        }
        let revision = document.package().source_revision().clone();
        let same_revision = {
            let snapshot = self.snapshot.read().await;
            snapshot.revision == revision
        };
        if same_revision {
            let recovered = self.error.write().await.take().is_some();
            if recovered {
                let _ = self.events.send(self.status().await);
            }
            return Ok(());
        }
        let rendered = document.html_view()?;
        {
            let mut snapshot = self.snapshot.write().await;
            snapshot.html = Bytes::from(rendered.content);
            snapshot.version = snapshot.version.saturating_add(1);
            snapshot.revision = revision;
            snapshot.render_sha256 = rendered.sha256;
        }
        *self.error.write().await = None;
        let _ = self.events.send(self.status().await);
        Ok(())
    }

    async fn report_error(&self, error: &UseError) {
        let next = NativeOfficeWatchError {
            code: error.code.to_string(),
            message: error.message.clone(),
        };
        let mut current = self.error.write().await;
        if current.as_ref() == Some(&next) {
            return;
        }
        *current = Some(next);
        drop(current);
        let _ = self.events.send(self.status().await);
    }
}

/// Bound, authenticated, read-only live preview server.
pub struct NativeOfficeWatchServer {
    listener: TcpListener,
    state: Arc<WatchState>,
    ready: NativeOfficeWatchReady,
}

impl NativeOfficeWatchServer {
    /// Validates and renders the initial document before binding a loopback port.
    pub async fn bind(
        path: impl AsRef<Path>,
        options: NativeOfficeWatchOptions,
    ) -> UseResult<Self> {
        let options = options.validate()?;
        let source = path.as_ref();
        let stamp_before = file_stamp(source).await.ok();
        let document = NativeOfficeDocument::open(source).await?;
        let rendered = document.html_view()?;
        let path = document.package().path().to_path_buf();
        let kind = document.kind();
        let revision = document.package().source_revision().clone();
        let stamp_after = file_stamp(&path).await.ok();
        let initial_stamp = (stamp_before == stamp_after)
            .then_some(stamp_after)
            .flatten();
        let listener = TcpListener::bind((Ipv4Addr::LOCALHOST, options.port))
            .await
            .map_err(|error| {
                UseError::new(
                    "use.office.watch_bind_failed",
                    format!(
                        "Failed to bind the native Office watch server on loopback port {}: {error}",
                        options.port
                    ),
                )
                .with_suggestion("Choose another --port, or use --port 0 for an ephemeral port.")
            })?;
        let address = listener.local_addr().map_err(|error| {
            UseError::new(
                "use.office.watch_bind_failed",
                format!("Failed to inspect the native Office watch listener: {error}"),
            )
        })?;
        let token = random_token()?;
        let url = format!("http://{address}/?token={token}");
        let document_name = path
            .file_name()
            .and_then(|value| value.to_str())
            .unwrap_or("document")
            .to_string();
        let (events, _) = broadcast::channel(EVENT_BUFFER);
        let (shutdown, _) = broadcast::channel(1);
        let ready = NativeOfficeWatchReady {
            url,
            address,
            document_name,
            kind,
            revision: revision.clone(),
            render_sha256: rendered.sha256.clone(),
            version: 1,
            poll_interval_ms: options.poll_interval_ms,
        };
        let state = Arc::new(WatchState {
            address,
            token,
            path,
            kind,
            poll_interval: options.poll_interval(),
            initial_stamp,
            snapshot: RwLock::new(WatchSnapshot {
                html: Bytes::from(rendered.content),
                version: 1,
                revision,
                render_sha256: rendered.sha256,
            }),
            error: RwLock::new(None),
            events,
            shutdown,
        });
        Ok(Self {
            listener,
            state,
            ready,
        })
    }

    pub fn ready(&self) -> &NativeOfficeWatchReady {
        &self.ready
    }

    /// Serves until the supplied shutdown future completes.
    pub async fn serve<F>(self, shutdown: F) -> UseResult<()>
    where
        F: Future<Output = ()> + Send + 'static,
    {
        let monitor_state = Arc::clone(&self.state);
        let monitor_shutdown = self.state.shutdown.subscribe();
        let monitor = tokio::spawn(async move { monitor(monitor_state, monitor_shutdown).await });
        let shutdown_state = Arc::clone(&self.state);
        let shutdown = async move {
            shutdown.await;
            let _ = shutdown_state.shutdown.send(());
        };
        let result = axum::serve(self.listener, routes::router(self.state))
            .with_graceful_shutdown(shutdown)
            .await;
        monitor.abort();
        let _ = monitor.await;
        result.map_err(|error| {
            UseError::new(
                "use.office.watch_server_failed",
                format!("Native Office watch server failed: {error}"),
            )
        })
    }
}

async fn monitor(state: Arc<WatchState>, mut shutdown: broadcast::Receiver<()>) {
    let mut accepted_stamp = state.initial_stamp.clone();
    let mut interval = tokio::time::interval(state.poll_interval);
    let mut retry_at = tokio::time::Instant::now();
    let mut retry_delay = state.poll_interval;
    interval.set_missed_tick_behavior(tokio::time::MissedTickBehavior::Skip);
    loop {
        tokio::select! {
            _ = interval.tick() => {}
            _ = shutdown.recv() => return,
        }
        if tokio::time::Instant::now() < retry_at {
            continue;
        }
        let stamp = match file_stamp(&state.path).await {
            Ok(stamp) => stamp,
            Err(error) => {
                state.report_error(&error).await;
                retry_at = tokio::time::Instant::now() + retry_delay;
                retry_delay = retry_delay.saturating_mul(2).min(MAX_RELOAD_RETRY);
                continue;
            }
        };
        if accepted_stamp.as_ref() == Some(&stamp) {
            continue;
        }
        match NativeOfficeDocument::open(&state.path).await {
            Ok(document) => match state.replace(document).await {
                Ok(()) => {
                    accepted_stamp = Some(stamp);
                    retry_at = tokio::time::Instant::now();
                    retry_delay = state.poll_interval;
                }
                Err(error) => {
                    state.report_error(&error).await;
                    retry_at = tokio::time::Instant::now() + retry_delay;
                    retry_delay = retry_delay.saturating_mul(2).min(MAX_RELOAD_RETRY);
                }
            },
            Err(error) => {
                state.report_error(&error).await;
                retry_at = tokio::time::Instant::now() + retry_delay;
                retry_delay = retry_delay.saturating_mul(2).min(MAX_RELOAD_RETRY);
            }
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct FileStamp {
    length: u64,
    modified: Option<SystemTime>,
    #[cfg(unix)]
    device: u64,
    #[cfg(unix)]
    inode: u64,
    #[cfg(unix)]
    changed_seconds: i64,
    #[cfg(unix)]
    changed_nanoseconds: i64,
}

async fn file_stamp(path: &Path) -> UseResult<FileStamp> {
    let metadata = tokio::fs::metadata(path).await.map_err(|error| {
        UseError::new(
            "use.office.watch_source_unavailable",
            format!(
                "Failed to inspect watched Office file '{}': {error}",
                path.display()
            ),
        )
    })?;
    if !metadata.is_file() {
        return Err(UseError::new(
            "use.office.watch_source_unavailable",
            format!(
                "Watched Office path '{}' is not a regular file.",
                path.display()
            ),
        ));
    }
    #[cfg(unix)]
    {
        use std::os::unix::fs::MetadataExt;
        Ok(FileStamp {
            length: metadata.len(),
            modified: metadata.modified().ok(),
            device: metadata.dev(),
            inode: metadata.ino(),
            changed_seconds: metadata.ctime(),
            changed_nanoseconds: metadata.ctime_nsec(),
        })
    }
    #[cfg(not(unix))]
    {
        Ok(FileStamp {
            length: metadata.len(),
            modified: metadata.modified().ok(),
        })
    }
}

fn random_token() -> UseResult<String> {
    let mut bytes = [0_u8; 32];
    getrandom::fill(&mut bytes).map_err(|error| {
        UseError::new(
            "use.office.watch_token_failed",
            format!("Failed to generate a native Office watch token: {error}"),
        )
    })?;
    Ok(bytes.iter().map(|byte| format!("{byte:02x}")).collect())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn watch_contracts_are_send_and_sync() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<NativeOfficeWatchOptions>();
        assert_send_sync::<NativeOfficeWatchReady>();
        assert_send_sync::<NativeOfficeWatchStatus>();
        assert_send_sync::<NativeOfficeWatchServer>();
    }

    #[tokio::test]
    async fn poll_interval_is_bounded_before_opening_a_document() {
        let error = NativeOfficeWatchServer::bind(
            "missing.docx",
            NativeOfficeWatchOptions {
                poll_interval_ms: MIN_NATIVE_OFFICE_WATCH_POLL_MS - 1,
                ..NativeOfficeWatchOptions::default()
            },
        )
        .await
        .err()
        .expect("invalid polling must fail");
        assert_eq!(error.code, "use.office.watch_poll_invalid");
    }
}
