use std::path::PathBuf;
use std::sync::atomic::AtomicBool;
use std::sync::{Arc, Mutex};

use a3s_boot::{BootError, Result};
use a3s_office::{
    NativeOfficeCollaborationActorKind, NativeOfficeCollaborationCreateRequest,
    NativeOfficeCollaborationMode, NativeOfficeCollaborationStore,
    NativeOfficeCollaborationTransportMessage, NativeOfficeCollaborationTransportPollResult,
    NativeOfficeCollaborationTransportReceiveRequest,
    NativeOfficeCollaborationTransportReceiveResult, NativeOfficeCollaborationTransportSession,
    MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH,
};
use sha2::{Digest, Sha256};

use crate::protocol::RoomIdentity;

#[derive(Debug)]
pub struct CollaborationRoom {
    pub identity: RoomIdentity,
    pub room_name: String,
    pub poller_started: AtomicBool,
    transport: Arc<Mutex<NativeOfficeCollaborationTransportSession>>,
    server_client_id: u64,
}

impl CollaborationRoom {
    pub async fn open(data_dir: PathBuf, identity: RoomIdentity) -> Result<Arc<Self>> {
        let room_hash = sha256_hex(identity.key_material().as_bytes());
        let store_path = data_dir.join(&room_hash);
        let create_identity = identity.clone();
        let create_hash = room_hash.clone();
        let transport = tokio::task::spawn_blocking(move || {
            std::fs::create_dir_all(&data_dir).map_err(BootError::from)?;
            let store = if store_path.exists() {
                NativeOfficeCollaborationStore::open(&store_path).map_err(map_office_error)?
            } else {
                NativeOfficeCollaborationStore::create(NativeOfficeCollaborationCreateRequest {
                    store: store_path,
                    artifact_id: create_identity.artifact_id,
                    kind: create_identity.artifact_kind,
                    actor_id: format!("server-{}", &create_hash[..32]),
                    actor_kind: NativeOfficeCollaborationActorKind::System,
                    mode: NativeOfficeCollaborationMode::Edit,
                    operation_id: format!("create-room-{create_hash}"),
                    namespace: Some(create_identity.namespace),
                    client_id: None,
                    initial_update: None,
                })
                .map_err(map_office_error)?
            };
            NativeOfficeCollaborationTransportSession::attach(store).map_err(map_office_error)
        })
        .await
        .map_err(map_join_error)??;
        let server_client_id = transport.manifest().client_id;
        Ok(Arc::new(Self {
            identity,
            room_name: format!("office:{room_hash}"),
            poller_started: AtomicBool::new(false),
            transport: Arc::new(Mutex::new(transport)),
            server_client_id,
        }))
    }

    pub fn server_client_id(&self) -> u64 {
        self.server_client_id
    }

    pub async fn synchronize(&self) -> Result<NativeOfficeCollaborationTransportMessage> {
        let transport = Arc::clone(&self.transport);
        tokio::task::spawn_blocking(move || {
            let transport = transport
                .lock()
                .map_err(|_| BootError::Internal("room transport lock is poisoned".to_string()))?;
            transport.synchronize().map_err(map_office_error)
        })
        .await
        .map_err(map_join_error)?
    }

    pub async fn receive(
        &self,
        request: NativeOfficeCollaborationTransportReceiveRequest,
    ) -> Result<NativeOfficeCollaborationTransportReceiveResult> {
        let transport = Arc::clone(&self.transport);
        tokio::task::spawn_blocking(move || {
            let mut transport = transport
                .lock()
                .map_err(|_| BootError::Internal("room transport lock is poisoned".to_string()))?;
            transport.receive(request).map_err(map_office_error)
        })
        .await
        .map_err(map_join_error)?
    }

    pub async fn poll(&self) -> Result<NativeOfficeCollaborationTransportPollResult> {
        let transport = Arc::clone(&self.transport);
        tokio::task::spawn_blocking(move || {
            let mut transport = transport
                .lock()
                .map_err(|_| BootError::Internal("room transport lock is poisoned".to_string()))?;
            transport
                .poll(MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH)
                .map_err(map_office_error)
        })
        .await
        .map_err(map_join_error)?
    }
}

pub fn delivery_operation_id(
    room: &RoomIdentity,
    sender_client_id: u64,
    message_type: a3s_office::NativeOfficeCollaborationTransportMessageType,
    payload: &[u8],
) -> String {
    let mut digest = Sha256::new();
    digest.update(b"a3s.office.collaboration.delivery.v1\0");
    digest.update(room.key_material().as_bytes());
    digest.update(b"\0");
    digest.update(sender_client_id.to_be_bytes());
    digest.update(format!("{message_type:?}").as_bytes());
    digest.update(payload);
    format!("delivery-{}", hex_digest(digest.finalize().as_slice()))
}

fn sha256_hex(value: &[u8]) -> String {
    hex_digest(Sha256::digest(value).as_slice())
}

fn hex_digest(value: &[u8]) -> String {
    let mut output = String::with_capacity(value.len() * 2);
    for byte in value {
        use std::fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn map_office_error(error: a3s_use_core::UseError) -> BootError {
    let message = format!("{}: {}", error.code, error.message);
    if error.code.contains("conflict") {
        BootError::Conflict(message)
    } else if error.code.contains("too_large") || error.code.contains("size") {
        BootError::PayloadTooLarge(message)
    } else if error.code.contains("store_corrupt") || error.code.contains("io") {
        BootError::InternalServerError(message)
    } else {
        BootError::BadRequest(message)
    }
}

fn map_join_error(error: tokio::task::JoinError) -> BootError {
    BootError::Internal(format!("blocking collaboration task failed: {error}"))
}
