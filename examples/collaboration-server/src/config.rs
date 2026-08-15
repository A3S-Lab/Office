use std::collections::BTreeSet;
use std::fmt;
use std::net::SocketAddr;
use std::path::{Path, PathBuf};
use std::time::Duration;

use a3s_boot::{parse_validated_acl_config, BootError, Result, Validate};
use a3s_office::MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES;
use serde::Deserialize;

const MIN_SECRET_BYTES: usize = 32;
const MIN_ADMIN_TOKEN_BYTES: usize = 16;

#[derive(Clone, Deserialize)]
pub struct CollaborationConfig {
    pub bind: SocketAddr,
    pub public_ws_url: String,
    pub data_dir: PathBuf,
    pub namespace: String,
    pub allowed_origins: BTreeSet<String>,
    pub ticket_ttl_seconds: u64,
    pub poll_interval_milliseconds: u64,
    pub max_document_payload_bytes: usize,
    pub max_awareness_payload_bytes: usize,
    pub ticket_secret: String,
    pub admin_token: String,
}

impl CollaborationConfig {
    pub fn from_acl_file(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();
        let input = std::fs::read_to_string(path).map_err(|error| {
            BootError::Internal(format!(
                "failed to read collaboration ACL config {}: {error}",
                path.display()
            ))
        })?;
        parse_validated_acl_config(&input)
    }

    pub fn ticket_ttl(&self) -> Duration {
        Duration::from_secs(self.ticket_ttl_seconds)
    }

    pub fn poll_interval(&self) -> Duration {
        Duration::from_millis(self.poll_interval_milliseconds)
    }

    pub fn validate_origin(&self, origin: Option<&str>) -> Result<()> {
        let origin = origin.ok_or_else(|| {
            BootError::Forbidden("the WebSocket Origin header is required".to_string())
        })?;
        if self.allowed_origins.contains(origin) {
            Ok(())
        } else {
            Err(BootError::Forbidden(format!(
                "WebSocket origin '{origin}' is not allowed"
            )))
        }
    }

    #[cfg(test)]
    pub fn for_test(data_dir: PathBuf) -> Self {
        Self {
            bind: ([127, 0, 0, 1], 0).into(),
            public_ws_url: "ws://127.0.0.1:8787/collaboration".to_string(),
            data_dir,
            namespace: "a3s.office".to_string(),
            allowed_origins: BTreeSet::from(["http://localhost:4175".to_string()]),
            ticket_ttl_seconds: 300,
            poll_interval_milliseconds: 25,
            max_document_payload_bytes: MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES,
            max_awareness_payload_bytes: 256 * 1024,
            ticket_secret: "test-ticket-secret-with-at-least-32-bytes".to_string(),
            admin_token: "test-admin-token".to_string(),
        }
    }
}

impl Validate for CollaborationConfig {
    fn validate(&self) -> Result<()> {
        if self.public_ws_url.trim_end_matches('/').is_empty()
            || !(self.public_ws_url.starts_with("ws://")
                || self.public_ws_url.starts_with("wss://"))
        {
            return Err(BootError::BadRequest(
                "public_ws_url must be an absolute ws:// or wss:// URL".to_string(),
            ));
        }
        if self.namespace.trim().is_empty()
            || self.namespace.len() > 256
            || self.namespace.starts_with('.')
            || self.namespace.ends_with('.')
            || !self
                .namespace
                .chars()
                .all(|character| character.is_ascii_alphanumeric() || "._-".contains(character))
        {
            return Err(BootError::BadRequest(
                "namespace must be a bounded dot-separated identifier".to_string(),
            ));
        }
        if self.allowed_origins.is_empty()
            || self
                .allowed_origins
                .iter()
                .any(|origin| !(origin.starts_with("http://") || origin.starts_with("https://")))
        {
            return Err(BootError::BadRequest(
                "allowed_origins must contain explicit HTTP origins".to_string(),
            ));
        }
        if !(30..=3_600).contains(&self.ticket_ttl_seconds) {
            return Err(BootError::BadRequest(
                "ticket_ttl_seconds must be between 30 and 3600".to_string(),
            ));
        }
        if !(10..=60_000).contains(&self.poll_interval_milliseconds) {
            return Err(BootError::BadRequest(
                "poll_interval_milliseconds must be between 10 and 60000".to_string(),
            ));
        }
        if self.max_document_payload_bytes == 0
            || self.max_document_payload_bytes > MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES
        {
            return Err(BootError::BadRequest(format!(
                "max_document_payload_bytes must be between 1 and {MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES}"
            )));
        }
        if !(1..=1024 * 1024).contains(&self.max_awareness_payload_bytes) {
            return Err(BootError::BadRequest(
                "max_awareness_payload_bytes must be between 1 and 1048576".to_string(),
            ));
        }
        if self.ticket_secret.as_bytes().len() < MIN_SECRET_BYTES {
            return Err(BootError::BadRequest(format!(
                "A3S_OFFICE_TICKET_SECRET must contain at least {MIN_SECRET_BYTES} bytes"
            )));
        }
        if self.admin_token.as_bytes().len() < MIN_ADMIN_TOKEN_BYTES {
            return Err(BootError::BadRequest(format!(
                "A3S_OFFICE_ADMIN_TOKEN must contain at least {MIN_ADMIN_TOKEN_BYTES} bytes"
            )));
        }
        Ok(())
    }
}

impl fmt::Debug for CollaborationConfig {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("CollaborationConfig")
            .field("bind", &self.bind)
            .field("public_ws_url", &self.public_ws_url)
            .field("data_dir", &self.data_dir)
            .field("namespace", &self.namespace)
            .field("allowed_origins", &self.allowed_origins)
            .field("ticket_ttl_seconds", &self.ticket_ttl_seconds)
            .field(
                "poll_interval_milliseconds",
                &self.poll_interval_milliseconds,
            )
            .field(
                "max_document_payload_bytes",
                &self.max_document_payload_bytes,
            )
            .field(
                "max_awareness_payload_bytes",
                &self.max_awareness_payload_bytes,
            )
            .field("ticket_secret", &"[redacted]")
            .field("admin_token", &"[redacted]")
            .finish()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn checked_in_acl_config_parses_and_validates() {
        let input = include_str!("../collaboration-server.acl")
            .replace(
                "env(\"A3S_OFFICE_PUBLIC_HOST\", \"127.0.0.1:8787\")",
                "\"127.0.0.1:8787\"",
            )
            .replace(
                "env(\"A3S_OFFICE_DATA_DIR\", \"./data/collaboration\")",
                "\"./data/collaboration\"",
            )
            .replace(
                "env(\"A3S_OFFICE_TICKET_SECRET\")",
                "\"test-ticket-secret-with-at-least-32-bytes\"",
            )
            .replace(
                "env(\"A3S_OFFICE_ADMIN_TOKEN\")",
                "\"test-admin-token-with-at-least-16-bytes\"",
            );

        let config: CollaborationConfig = parse_validated_acl_config(&input).unwrap();
        assert_eq!(config.bind, "127.0.0.1:8787".parse().unwrap());
        assert_eq!(config.public_ws_url, "ws://127.0.0.1:8787/collaboration");
        assert_eq!(config.namespace, "a3s.office");
    }
}
