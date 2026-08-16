use std::time::{SystemTime, UNIX_EPOCH};

use a3s_boot::{BootError, Result};
use a3s_office::{
    NATIVE_OFFICE_COLLABORATION_PROTOCOL, NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
};
use base64::engine::general_purpose::URL_SAFE_NO_PAD;
use base64::Engine as _;
use hmac::{Hmac, Mac};
use sha2::Sha256;

use crate::config::CollaborationConfig;
use crate::protocol::{TicketClaims, TicketRequest, TicketResponse};

type HmacSha256 = Hmac<Sha256>;

const TICKET_VERSION: u32 = 2;
const CLOCK_SKEW_SECONDS: u64 = 5;

#[derive(Clone)]
pub struct TicketService {
    secret: Vec<u8>,
    admin_token: String,
    namespace: String,
    public_ws_url: String,
    ttl_seconds: u64,
}

impl TicketService {
    pub fn new(config: &CollaborationConfig) -> Self {
        Self {
            secret: config.ticket_secret.as_bytes().to_vec(),
            admin_token: config.admin_token.clone(),
            namespace: config.namespace.clone(),
            public_ws_url: config.public_ws_url.trim_end_matches('/').to_string(),
            ttl_seconds: config.ticket_ttl().as_secs(),
        }
    }

    pub fn issue(
        &self,
        authorization: Option<&str>,
        request: TicketRequest,
    ) -> Result<TicketResponse> {
        self.authorize_admin(authorization)?;
        request.validate()?;
        let issued_at = unix_timestamp()?;
        let expires_at = issued_at
            .checked_add(self.ttl_seconds)
            .ok_or_else(|| BootError::Internal("ticket expiration overflowed".to_string()))?;
        let claims = TicketClaims {
            ticket_version: TICKET_VERSION,
            artifact_id: request.artifact_id,
            artifact_kind: request.artifact_kind,
            namespace: self.namespace.clone(),
            actor_id: request.actor_id,
            actor_name: request.actor_name,
            actor_kind: request.actor_kind,
            mode: request.mode,
            issued_at,
            expires_at,
        };
        let ticket = self.sign(&claims)?;
        let web_socket_url = format!(
            "{}/{}/{}?ticket={}",
            self.public_ws_url,
            claims.artifact_kind.as_str(),
            claims.artifact_id,
            ticket
        );
        Ok(TicketResponse {
            ticket,
            web_socket_url,
            expires_at,
            protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL,
            version: NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
        })
    }

    pub fn verify(&self, ticket: &str) -> Result<TicketClaims> {
        let (payload, signature) = ticket.split_once('.').ok_or_else(|| {
            BootError::Unauthorized("the collaboration ticket is malformed".to_string())
        })?;
        if payload.is_empty() || signature.is_empty() || signature.contains('.') {
            return Err(BootError::Unauthorized(
                "the collaboration ticket is malformed".to_string(),
            ));
        }
        let signature = URL_SAFE_NO_PAD.decode(signature).map_err(|_| {
            BootError::Unauthorized("the collaboration ticket signature is malformed".to_string())
        })?;
        let mut mac = HmacSha256::new_from_slice(&self.secret)
            .map_err(|error| BootError::Internal(format!("invalid HMAC key: {error}")))?;
        mac.update(payload.as_bytes());
        mac.verify_slice(&signature).map_err(|_| {
            BootError::Unauthorized("the collaboration ticket signature is invalid".to_string())
        })?;
        let payload = URL_SAFE_NO_PAD.decode(payload).map_err(|_| {
            BootError::Unauthorized("the collaboration ticket payload is malformed".to_string())
        })?;
        let claims: TicketClaims = serde_json::from_slice(&payload).map_err(|_| {
            BootError::Unauthorized("the collaboration ticket payload is invalid".to_string())
        })?;
        self.validate_claims(&claims)?;
        Ok(claims)
    }

    fn sign(&self, claims: &TicketClaims) -> Result<String> {
        let payload = serde_json::to_vec(claims)
            .map_err(|error| BootError::Internal(format!("failed to encode ticket: {error}")))?;
        let payload = URL_SAFE_NO_PAD.encode(payload);
        let mut mac = HmacSha256::new_from_slice(&self.secret)
            .map_err(|error| BootError::Internal(format!("invalid HMAC key: {error}")))?;
        mac.update(payload.as_bytes());
        let signature = URL_SAFE_NO_PAD.encode(mac.finalize().into_bytes());
        Ok(format!("{payload}.{signature}"))
    }

    fn authorize_admin(&self, authorization: Option<&str>) -> Result<()> {
        let token = authorization
            .and_then(|value| value.strip_prefix("Bearer "))
            .filter(|value| !value.is_empty())
            .ok_or_else(|| {
                BootError::Unauthorized(
                    "a Bearer token is required to issue collaboration tickets".to_string(),
                )
            })?;
        if constant_time_token_match(&self.admin_token, token)? {
            Ok(())
        } else {
            Err(BootError::Unauthorized(
                "the ticket issuer Bearer token is invalid".to_string(),
            ))
        }
    }

    fn validate_claims(&self, claims: &TicketClaims) -> Result<()> {
        if claims.ticket_version != TICKET_VERSION || claims.namespace != self.namespace {
            return Err(BootError::Unauthorized(
                "the collaboration ticket version or namespace is unsupported".to_string(),
            ));
        }
        claims.room_identity().map_err(|_| {
            BootError::Unauthorized("the collaboration ticket room identity is invalid".to_string())
        })?;
        if claims.actor_id.trim().is_empty() || claims.actor_id.as_bytes().len() > 256 {
            return Err(BootError::Unauthorized(
                "the collaboration ticket actor is invalid".to_string(),
            ));
        }
        if claims.actor_name.is_empty()
            || claims.actor_name != claims.actor_name.trim()
            || claims.actor_name.chars().count() > 256
        {
            return Err(BootError::Unauthorized(
                "the collaboration ticket actor name is invalid".to_string(),
            ));
        }
        let now = unix_timestamp()?;
        if claims.issued_at > now.saturating_add(CLOCK_SKEW_SECONDS)
            || claims.expires_at <= now
            || claims.expires_at.saturating_sub(claims.issued_at) > self.ttl_seconds
        {
            return Err(BootError::Unauthorized(
                "the collaboration ticket is expired or has an invalid lifetime".to_string(),
            ));
        }
        Ok(())
    }
}

fn constant_time_token_match(expected: &str, received: &str) -> Result<bool> {
    const DOMAIN: &[u8] = b"a3s.office.collaboration.admin-token.v1";
    let mut expected_mac = HmacSha256::new_from_slice(expected.as_bytes())
        .map_err(|error| BootError::Internal(format!("invalid admin token key: {error}")))?;
    expected_mac.update(DOMAIN);
    let expected_tag = expected_mac.finalize().into_bytes();

    let mut received_mac = HmacSha256::new_from_slice(received.as_bytes())
        .map_err(|error| BootError::Internal(format!("invalid admin token key: {error}")))?;
    received_mac.update(DOMAIN);
    Ok(received_mac.verify_slice(&expected_tag).is_ok())
}

pub fn unix_timestamp() -> Result<u64> {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_secs())
        .map_err(|error| BootError::Internal(format!("system clock is before Unix epoch: {error}")))
}

#[cfg(test)]
mod tests {
    use super::*;
    use a3s_office::{
        NativeOfficeCollaborationActorKind, NativeOfficeCollaborationArtifactKind,
        NativeOfficeCollaborationMode,
    };

    #[test]
    fn tickets_are_signed_scoped_and_tamper_evident() {
        let temp = tempfile::tempdir().unwrap();
        let config = CollaborationConfig::for_test(temp.path().to_path_buf());
        let service = TicketService::new(&config);
        let response = service
            .issue(
                Some("Bearer test-admin-token"),
                TicketRequest {
                    artifact_id: "quarterly-plan".to_string(),
                    artifact_kind: NativeOfficeCollaborationArtifactKind::Document,
                    actor_id: "user-42".to_string(),
                    actor_name: "Ada Reviewer".to_string(),
                    actor_kind: NativeOfficeCollaborationActorKind::Human,
                    mode: NativeOfficeCollaborationMode::Edit,
                },
            )
            .unwrap();
        let claims = service.verify(&response.ticket).unwrap();
        assert_eq!(claims.artifact_id, "quarterly-plan");
        assert_eq!(claims.actor_id, "user-42");
        assert_eq!(claims.actor_name, "Ada Reviewer");
        assert_eq!(claims.mode, NativeOfficeCollaborationMode::Edit);

        let mut tampered = response.ticket.into_bytes();
        let last = tampered.last_mut().unwrap();
        *last = if *last == b'A' { b'B' } else { b'A' };
        assert!(service
            .verify(std::str::from_utf8(&tampered).unwrap())
            .is_err());
    }

    #[test]
    fn ticket_issuer_requires_the_admin_bearer_token() {
        let temp = tempfile::tempdir().unwrap();
        let config = CollaborationConfig::for_test(temp.path().to_path_buf());
        let service = TicketService::new(&config);
        let request = TicketRequest {
            artifact_id: "quarterly-plan".to_string(),
            artifact_kind: NativeOfficeCollaborationArtifactKind::Document,
            actor_id: "user-42".to_string(),
            actor_name: "Ada Reviewer".to_string(),
            actor_kind: NativeOfficeCollaborationActorKind::Human,
            mode: NativeOfficeCollaborationMode::View,
        };
        assert!(matches!(
            service.issue(Some("Bearer wrong-token-value"), request),
            Err(BootError::Unauthorized(_))
        ));
    }
}
