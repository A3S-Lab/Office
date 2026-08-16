use std::collections::HashMap;
use std::fmt;
use std::sync::atomic::Ordering;
use std::sync::Arc;

use a3s_boot::{
    BootError, Result, WebSocketGatewayConnection, WebSocketGatewayServer, WebSocketMessage,
};
use a3s_office::{
    NativeOfficeCollaborationTransportAuthorization, NativeOfficeCollaborationTransportMessageType,
    NativeOfficeCollaborationTransportReceiveRequest, NATIVE_OFFICE_COLLABORATION_PROTOCOL,
    NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
};
use serde_json::{json, Value};
use tokio::sync::{Mutex, RwLock};
use yrs::sync::AwarenessUpdate;
use yrs::updates::decoder::Decode;

use crate::config::CollaborationConfig;
use crate::protocol::{
    parse_artifact_kind, trusted_origin, websocket_document_message, AwarenessWireMessage,
    DocumentWireMessage, HelloMessage, RoomIdentity, TicketClaims, TicketRequest, TicketResponse,
    ACK_EVENT, AWARENESS_EVENT, PEER_LEFT_EVENT, READY_EVENT,
};
use crate::room::{delivery_operation_id, CollaborationRoom};
use crate::ticket::TicketService;

#[derive(Clone)]
struct ConnectionState {
    claims: TicketClaims,
    room: Arc<CollaborationRoom>,
    client_id: Option<u64>,
    awareness: Option<AwarenessWireMessage>,
}

pub struct CollaborationService {
    config: CollaborationConfig,
    tickets: TicketService,
    rooms: Mutex<HashMap<String, Arc<CollaborationRoom>>>,
    connections: RwLock<HashMap<u64, ConnectionState>>,
}

impl CollaborationService {
    pub fn new(config: CollaborationConfig) -> Self {
        Self {
            tickets: TicketService::new(&config),
            config,
            rooms: Mutex::new(HashMap::new()),
            connections: RwLock::new(HashMap::new()),
        }
    }

    pub fn issue_ticket(
        &self,
        authorization: Option<&str>,
        request: TicketRequest,
    ) -> Result<TicketResponse> {
        self.tickets.issue(authorization, request)
    }

    pub async fn authenticate_connection(
        self: &Arc<Self>,
        connection: WebSocketGatewayConnection,
    ) -> Result<()> {
        self.config
            .validate_origin(connection.request().header("origin"))?;
        let ticket = connection
            .request()
            .query_param("ticket")
            .ok_or_else(|| BootError::Unauthorized("missing collaboration ticket".to_string()))?;
        let claims = self.tickets.verify(ticket)?;
        let path_kind = connection.request().param("artifactKind").ok_or_else(|| {
            BootError::BadRequest("missing artifact kind path parameter".to_string())
        })?;
        let path_id = connection.request().param("artifactId").ok_or_else(|| {
            BootError::BadRequest("missing artifact ID path parameter".to_string())
        })?;
        let path_kind = parse_artifact_kind(path_kind)?;
        if path_kind != claims.artifact_kind || path_id != claims.artifact_id {
            return Err(BootError::Forbidden(
                "the ticket does not authorize this collaboration room".to_string(),
            ));
        }
        let identity = claims.room_identity()?;
        let room = self.room(identity).await?;
        connection.join(room.room_name.clone())?;
        let previous = self.connections.write().await.insert(
            connection.id(),
            ConnectionState {
                claims,
                room: Arc::clone(&room),
                client_id: None,
                awareness: None,
            },
        );
        if previous.is_some() {
            return Err(BootError::Conflict(format!(
                "WebSocket connection {} is already authenticated",
                connection.id()
            )));
        }
        self.start_room_poller(room, connection.server());
        Ok(())
    }

    pub async fn disconnect(&self, connection: WebSocketGatewayConnection) -> Result<()> {
        let Some(state) = self.connections.write().await.remove(&connection.id()) else {
            return Ok(());
        };
        if let Some(client_id) = state.client_id {
            connection
                .broadcast_to_room(
                    state.room.room_name.clone(),
                    WebSocketMessage::new(
                        PEER_LEFT_EVENT,
                        json!({
                            "protocol": NATIVE_OFFICE_COLLABORATION_PROTOCOL,
                            "version": NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
                            "senderClientId": client_id,
                        }),
                    ),
                )
                .await?;
        }
        Ok(())
    }

    pub async fn hello(
        &self,
        connection: WebSocketGatewayConnection,
        message: WebSocketMessage,
    ) -> Result<()> {
        let hello = HelloMessage::from_message(message)?;
        let state = {
            let mut connections = self.connections.write().await;
            let room_name = connections
                .get(&connection.id())
                .ok_or_else(|| {
                    BootError::Unauthorized("connection is not authenticated".to_string())
                })?
                .room
                .room_name
                .clone();
            if connections.iter().any(|(id, member)| {
                *id != connection.id()
                    && member.room.room_name == room_name
                    && member.client_id == Some(hello.sender_client_id)
                    && member.client_id.is_some()
            }) {
                return Err(BootError::Conflict(format!(
                    "Yjs client ID {} is already active in this room",
                    hello.sender_client_id
                )));
            }
            let state = connections.get_mut(&connection.id()).ok_or_else(|| {
                BootError::Unauthorized("connection is not authenticated".to_string())
            })?;
            if hello.sender_client_id == state.room.server_client_id() {
                return Err(BootError::Conflict(
                    "the browser Yjs client ID collides with the durable server replica"
                        .to_string(),
                ));
            }
            match state.client_id {
                Some(existing) if existing != hello.sender_client_id => {
                    return Err(BootError::Conflict(
                        "a WebSocket connection cannot change its Yjs client ID".to_string(),
                    ));
                }
                _ => state.client_id = Some(hello.sender_client_id),
            }
            state.clone()
        };

        connection
            .emit(WebSocketMessage::new(
                READY_EVENT,
                json!({
                    "protocol": NATIVE_OFFICE_COLLABORATION_PROTOCOL,
                    "version": NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
                    "artifactId": state.claims.artifact_id,
                    "artifactKind": state.claims.artifact_kind,
                    "namespace": state.claims.namespace,
                    "actorId": state.claims.actor_id,
                    "actorName": state.claims.actor_name,
                    "actorKind": state.claims.actor_kind,
                    "mode": state.claims.mode,
                    "senderClientId": state.room.server_client_id(),
                }),
            ))
            .await?;

        if state.claims.can_publish_document_updates() {
            let handshake = state.room.synchronize().await?;
            connection
                .emit(websocket_document_message(&handshake)?)
                .await?;
        }
        for awareness in self.awareness_snapshots(connection.id(), &state.room).await {
            connection
                .emit(WebSocketMessage::json(AWARENESS_EVENT, &awareness)?)
                .await?;
        }
        Ok(())
    }

    pub async fn receive_document(
        &self,
        connection: WebSocketGatewayConnection,
        message: WebSocketMessage,
    ) -> Result<()> {
        let state = self.registered_connection(connection.id()).await?;
        let wire = DocumentWireMessage::from_message(message)?;
        let mut native =
            wire.decode(&state.room.identity, self.config.max_document_payload_bytes)?;
        self.require_sender(&state, native.sender_client_id)?;
        let mutating =
            native.message_type != NativeOfficeCollaborationTransportMessageType::SyncStep1;
        if mutating && !state.claims.can_publish_document_updates() {
            return Err(BootError::Forbidden(
                "view and suggest tickets cannot publish Yjs document updates".to_string(),
            ));
        }
        let operation_id = mutating.then(|| {
            delivery_operation_id(
                &state.room.identity,
                native.sender_client_id,
                native.message_type,
                &native.payload,
            )
        });
        if native.message_type == NativeOfficeCollaborationTransportMessageType::Update {
            native.origin = operation_id
                .as_ref()
                .map(|operation_id| trusted_origin(&state.claims, operation_id.clone()));
        }
        let broadcast = native.clone();
        let result = state
            .room
            .receive(
                NativeOfficeCollaborationTransportReceiveRequest {
                    message: native,
                    operation_id: operation_id.clone(),
                    if_state_vector: None,
                },
                NativeOfficeCollaborationTransportAuthorization {
                    actor_id: state.claims.actor_id.clone(),
                    actor_kind: state.claims.actor_kind,
                    actor_name: state.claims.actor_name.clone(),
                    mode: state.claims.mode,
                },
            )
            .await?;

        if let Some(response) = result.response {
            connection
                .emit(websocket_document_message(&response)?)
                .await?;
        }
        if result
            .apply
            .as_ref()
            .is_some_and(|apply| apply.state_changed)
        {
            connection
                .broadcast_to_room(
                    state.room.room_name.clone(),
                    websocket_document_message(&broadcast)?,
                )
                .await?;
        }
        if let (Some(operation_id), Some(apply)) = (operation_id, result.apply) {
            connection
                .emit(WebSocketMessage::new(
                    ACK_EVENT,
                    json!({
                        "operationId": operation_id,
                        "sequence": apply.sequence,
                        "duplicate": apply.duplicate,
                        "stateChanged": apply.state_changed,
                    }),
                ))
                .await?;
        }
        Ok(())
    }

    pub async fn receive_awareness(
        &self,
        connection: WebSocketGatewayConnection,
        message: WebSocketMessage,
    ) -> Result<()> {
        let state = self.registered_connection(connection.id()).await?;
        let wire = AwarenessWireMessage::from_message(message)?;
        self.require_sender(&state, wire.sender_client_id)?;
        let payload = wire.decode(
            &state.room.identity,
            self.config.max_awareness_payload_bytes,
        )?;
        validate_awareness_payload(&payload, wire.sender_client_id, &state)?;
        self.connections
            .write()
            .await
            .get_mut(&connection.id())
            .ok_or_else(|| BootError::Unauthorized("connection is not authenticated".to_string()))?
            .awareness = Some(wire.clone());
        connection
            .broadcast_to_room(
                state.room.room_name.clone(),
                WebSocketMessage::json(AWARENESS_EVENT, &wire)?,
            )
            .await?;
        Ok(())
    }

    async fn room(&self, identity: RoomIdentity) -> Result<Arc<CollaborationRoom>> {
        let key = identity.key_material();
        let mut rooms = self.rooms.lock().await;
        if let Some(room) = rooms.get(&key) {
            return Ok(Arc::clone(room));
        }
        let room = CollaborationRoom::open(self.config.data_dir.clone(), identity).await?;
        rooms.insert(key, Arc::clone(&room));
        Ok(room)
    }

    async fn registered_connection(&self, connection_id: u64) -> Result<ConnectionState> {
        let state = self
            .connections
            .read()
            .await
            .get(&connection_id)
            .cloned()
            .ok_or_else(|| {
                BootError::Unauthorized("connection is not authenticated".to_string())
            })?;
        if state.client_id.is_none() {
            return Err(BootError::BadRequest(
                "collaboration.hello must be sent before room messages".to_string(),
            ));
        }
        Ok(state)
    }

    fn require_sender(&self, state: &ConnectionState, sender_client_id: u64) -> Result<()> {
        if state.client_id == Some(sender_client_id) {
            Ok(())
        } else {
            Err(BootError::Forbidden(
                "the message sender client ID does not match this WebSocket connection".to_string(),
            ))
        }
    }

    async fn awareness_snapshots(
        &self,
        connection_id: u64,
        room: &CollaborationRoom,
    ) -> Vec<AwarenessWireMessage> {
        self.connections
            .read()
            .await
            .iter()
            .filter(|(id, state)| **id != connection_id && state.room.room_name == room.room_name)
            .filter_map(|(_, state)| state.awareness.clone())
            .collect()
    }

    fn start_room_poller(
        self: &Arc<Self>,
        room: Arc<CollaborationRoom>,
        server: WebSocketGatewayServer,
    ) {
        if room
            .poller_started
            .compare_exchange(false, true, Ordering::AcqRel, Ordering::Acquire)
            .is_err()
        {
            return;
        }
        let interval = self.config.poll_interval();
        tokio::spawn(async move {
            let mut ticker = tokio::time::interval(interval);
            ticker.set_missed_tick_behavior(tokio::time::MissedTickBehavior::Skip);
            loop {
                ticker.tick().await;
                match room.poll().await {
                    Ok(batch) => {
                        for message in batch.messages {
                            let outbound = match websocket_document_message(&message) {
                                Ok(outbound) => outbound,
                                Err(error) => {
                                    eprintln!("failed to encode collaboration update: {error}");
                                    continue;
                                }
                            };
                            if let Err(error) = server
                                .broadcast_to_room(room.room_name.clone(), outbound)
                                .await
                            {
                                eprintln!("failed to broadcast collaboration update: {error}");
                            }
                        }
                    }
                    Err(error) => eprintln!("failed to poll collaboration room: {error}"),
                }
            }
        });
    }
}

impl fmt::Debug for CollaborationService {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("CollaborationService")
            .field("config", &self.config)
            .finish_non_exhaustive()
    }
}

fn validate_awareness_payload(
    payload: &[u8],
    sender_client_id: u64,
    state: &ConnectionState,
) -> Result<()> {
    let update = AwarenessUpdate::decode_v1(payload)
        .map_err(|error| BootError::BadRequest(format!("invalid Yjs Awareness update: {error}")))?;
    if update.clients.len() != 1 {
        return Err(BootError::BadRequest(
            "an awareness message must update exactly one authenticated client".to_string(),
        ));
    }
    let Some((client_id, entry)) = update.clients.iter().next() else {
        return Err(BootError::BadRequest(
            "an awareness message must update exactly one authenticated client".to_string(),
        ));
    };
    if client_id.get() != sender_client_id {
        return Err(BootError::Forbidden(
            "an awareness update cannot impersonate another Yjs client".to_string(),
        ));
    }
    let value: Value = serde_json::from_str(entry.json.as_ref()).map_err(|error| {
        BootError::BadRequest(format!("awareness state is not valid JSON: {error}"))
    })?;
    if value.is_null() {
        return Ok(());
    }
    let office = value
        .get("a3sOffice")
        .and_then(Value::as_object)
        .ok_or_else(|| {
            BootError::BadRequest(
                "awareness state must contain the typed a3sOffice presence field".to_string(),
            )
        })?;
    let actor = office.get("actor").and_then(Value::as_object);
    let valid = office.get("protocol").and_then(Value::as_str)
        == Some(NATIVE_OFFICE_COLLABORATION_PROTOCOL)
        && office.get("version").and_then(Value::as_u64)
            == Some(NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION as u64)
        && office.get("artifactId").and_then(Value::as_str)
            == Some(state.room.identity.artifact_id.as_str())
        && office.get("artifactKind").and_then(Value::as_str)
            == Some(state.room.identity.artifact_kind.as_str())
        && office.get("namespace").and_then(Value::as_str)
            == Some(state.room.identity.namespace.as_str())
        && office.get("mode").and_then(Value::as_str) == Some(state.claims.mode.as_str())
        && actor
            .and_then(|actor| actor.get("id"))
            .and_then(Value::as_str)
            == Some(state.claims.actor_id.as_str())
        && actor
            .and_then(|actor| actor.get("kind"))
            .and_then(Value::as_str)
            == Some(state.claims.actor_kind.as_str())
        && actor
            .and_then(|actor| actor.get("name"))
            .and_then(Value::as_str)
            == Some(state.claims.actor_name.as_str());
    if valid {
        Ok(())
    } else {
        Err(BootError::Forbidden(
            "the Office presence identity does not match the signed collaboration ticket"
                .to_string(),
        ))
    }
}
