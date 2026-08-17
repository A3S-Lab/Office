use std::path::PathBuf;
use std::sync::{Arc, Mutex};
use std::time::Duration;

use a3s_office::{
    NativeOfficeCollaborationAwarenessMessage, NativeOfficeCollaborationOrigin,
    NativeOfficeCollaborationPresenceActivity, NativeOfficeCollaborationPresenceLocation,
    NativeOfficeCollaborationPresenceProfile, NativeOfficeCollaborationPresenceSession,
    NativeOfficeCollaborationPresenceSnapshot, NativeOfficeCollaborationPresenceUpdate,
    NativeOfficeCollaborationStore, NativeOfficeCollaborationTransportMessage,
    NativeOfficeCollaborationTransportMessageType, NativeOfficeCollaborationTransportPollResult,
    NativeOfficeCollaborationTransportReceiveRequest,
    NativeOfficeCollaborationTransportReceiveResult, NativeOfficeCollaborationTransportSession,
    MAX_NATIVE_OFFICE_COLLABORATION_AWARENESS_BYTES, MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH,
    MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES,
    MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES,
};
use a3s_use_core::{UseError, UseResult};
use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use serde::Deserialize;
use serde_json::{json, Value};
use tokio::io::{AsyncBufRead, AsyncBufReadExt, AsyncWriteExt, BufReader, Stdout};

use crate::CommandOutput;

use super::arguments::{collaboration_cli_error, usage_error};
use super::binary_io::spawn_blocking;

const DEFAULT_POLL_MS: u64 = 100;
const MIN_POLL_MS: u64 = 50;
const MAX_POLL_MS: u64 = 10_000;
const MAX_TIMEOUT_MS: u64 = 24 * 60 * 60 * 1_000;
const MAX_YJS_CLIENT_ID: u64 = (1_u64 << 53) - 1;
const MAX_SESSION_JSONL_BYTES: usize =
    MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES.div_ceil(3) * 4 + 64 * 1024;

type SharedSession = Arc<Mutex<NativeOfficeCollaborationTransportSession>>;

pub(super) async fn run(args: &[String]) -> UseResult<CommandOutput> {
    let arguments = SessionArguments::parse(args)?;
    let store = arguments.store.clone();
    let presence_enabled = arguments.presence_profile.is_some();
    let session = spawn_blocking(move || {
        let store = NativeOfficeCollaborationStore::open(store)?;
        if presence_enabled {
            let replica_client_id = store.inspect()?.manifest.client_id;
            let sender_client_id = random_client_id(replica_client_id)?;
            NativeOfficeCollaborationTransportSession::attach_with_sender_client_id(
                store,
                sender_client_id,
            )
        } else {
            NativeOfficeCollaborationTransportSession::attach(store)
        }
    })
    .await?;
    let manifest = session.manifest().clone();
    let starting_sequence = session.cursor_sequence();
    let sender_client_id = session.sender_client_id();
    let session = Arc::new(Mutex::new(session));
    let mut presence = arguments
        .presence_profile
        .clone()
        .map(|profile| {
            NativeOfficeCollaborationPresenceSession::new_with_sender_client_id(
                manifest.clone(),
                sender_client_id,
                profile,
            )
        })
        .transpose()?;
    let mut output = SessionOutput::new();

    output
        .write(json!({
            "schemaVersion": 1,
            "type": "ready",
            "store": arguments.store,
            "protocol": manifest.protocol,
            "version": manifest.protocol_version,
            "artifactId": manifest.artifact_id,
            "artifactKind": manifest.kind,
            "namespace": manifest.namespace,
            "actorId": manifest.actor_id,
            "actorKind": manifest.actor_kind,
            "mode": manifest.mode,
            "clientId": sender_client_id,
            "replicaClientId": manifest.client_id,
            "presenceEnabled": presence.is_some(),
            "actorName": arguments.presence_profile.as_ref().map(|profile| &profile.name),
            "startingSequence": starting_sequence,
            "cursorSequence": starting_sequence,
            "pollMs": arguments.poll_ms,
            "timeoutMs": arguments.timeout_ms,
        }))
        .await?;

    let result = run_loop(&session, &arguments, &mut presence, &mut output).await;
    match result {
        Ok(()) => Ok(CommandOutput::silent()),
        Err(error) => {
            output.error(&error).await?;
            Ok(CommandOutput::silent_failure())
        }
    }
}

fn random_client_id(replica_client_id: u64) -> UseResult<u64> {
    loop {
        let mut bytes = [0_u8; 8];
        getrandom::fill(&mut bytes).map_err(|error| {
            collaboration_cli_error(
                "office.collaboration.session_failed",
                format!("Failed to create an ephemeral Presence client ID: {error}"),
            )
        })?;
        let client_id = u64::from_le_bytes(bytes) & MAX_YJS_CLIENT_ID;
        if client_id != 0 && client_id != replica_client_id {
            return Ok(client_id);
        }
    }
}

async fn run_loop(
    session: &SharedSession,
    arguments: &SessionArguments,
    presence: &mut Option<NativeOfficeCollaborationPresenceSession>,
    output: &mut SessionOutput,
) -> UseResult<()> {
    let initial = session_call(session, |session| session.synchronize()).await?;
    output.outbound("initial-connect", &initial).await?;
    if let Some(presence) = presence.as_ref() {
        output
            .outbound_awareness("initial-connect", &presence.local_message()?)
            .await?;
    }

    let mut stdin = BufReader::new(tokio::io::stdin());
    let mut input = Vec::new();
    let mut interval = tokio::time::interval(Duration::from_millis(arguments.poll_ms));
    interval.set_missed_tick_behavior(tokio::time::MissedTickBehavior::Skip);
    let timeout = timeout_signal(arguments.timeout_ms);
    tokio::pin!(timeout);
    let termination = termination_signal();
    tokio::pin!(termination);

    loop {
        tokio::select! {
            line = read_bounded_jsonl(&mut stdin, &mut input) => {
                let Some(line) = line? else {
                    complete_session("stdin-eof", session, presence, output).await?;
                    return Ok(());
                };
                match parse_input(&line)? {
                    SessionInput::Reconnect => {
                        let message = session_call(session, |session| session.synchronize()).await?;
                        output.outbound("reconnect", &message).await?;
                        if let Some(presence) = presence.as_mut() {
                            let cleared = presence.clear_remote()?;
                            if cleared.changed {
                                output.presence("reconnect-clear", cleared.changed, &cleared.snapshot).await?;
                            }
                            output
                                .outbound_awareness("reconnect", &presence.local_message()?)
                                .await?;
                        }
                    }
                    SessionInput::Receive {
                        message,
                        operation_id,
                        if_state_vector_base64,
                    } => {
                        let request = receive_request(
                            *message,
                            operation_id.clone(),
                            if_state_vector_base64,
                        )?;
                        let received = session_call(session, move |session| session.receive(request)).await?;
                        output.received(operation_id.as_deref(), &received).await?;
                        if let Some(response) = &received.response {
                            output.outbound("peer-sync-step1", response).await?;
                        }
                    }
                    SessionInput::SetPresence { activity, location } => {
                        let presence = required_presence(presence)?;
                        let message = presence.update(NativeOfficeCollaborationPresenceUpdate {
                            activity,
                            location,
                        })?;
                        output.outbound_awareness("local-presence", &message).await?;
                        output
                            .presence("local-presence", true, &presence.snapshot()?)
                            .await?;
                    }
                    SessionInput::ReceiveAwareness { message } => {
                        let presence = required_presence(presence)?;
                        let message = awareness_message(*message)?;
                        let received = presence.receive(message)?;
                        output
                            .presence("remote-awareness", received.changed, &received.snapshot)
                            .await?;
                    }
                    SessionInput::PeerLeft { sender_client_id } => {
                        let presence = required_presence(presence)?;
                        let received = presence.peer_left(sender_client_id)?;
                        output
                            .presence("peer-left", received.changed, &received.snapshot)
                            .await?;
                    }
                    SessionInput::Close => {
                        drain_backlog(session, output).await?;
                        complete_session("host-close", session, presence, output).await?;
                        return Ok(());
                    }
                }
            }
            _ = interval.tick() => {
                let batch = session_call(session, |session| {
                    session.poll(MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH)
                }).await?;
                output.poll(&batch).await?;
                if batch.has_more {
                    drain_backlog(session, output).await?;
                }
            }
            _ = &mut timeout => {
                complete_session("timeout", session, presence, output).await?;
                return Ok(());
            }
            _ = &mut termination => {
                complete_session("signal", session, presence, output).await?;
                return Ok(());
            }
        }
    }
}

async fn complete_session(
    reason: &str,
    session: &SharedSession,
    presence: &mut Option<NativeOfficeCollaborationPresenceSession>,
    output: &mut SessionOutput,
) -> UseResult<()> {
    if let Some(presence) = presence.as_mut() {
        let removal = presence.disconnect()?;
        output.outbound_awareness("disconnect", &removal).await?;
    }
    output
        .complete(reason, session_cursor(session).await?)
        .await
}

fn required_presence(
    presence: &mut Option<NativeOfficeCollaborationPresenceSession>,
) -> UseResult<&mut NativeOfficeCollaborationPresenceSession> {
    presence.as_mut().ok_or_else(|| {
        collaboration_cli_error(
            "office.collaboration.presence_unavailable",
            "Native Presence requires collab session --actor-name <display-name>.",
        )
    })
}

async fn drain_backlog(session: &SharedSession, output: &mut SessionOutput) -> UseResult<()> {
    loop {
        let batch = session_call(session, |session| {
            session.poll(MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH)
        })
        .await?;
        output.poll(&batch).await?;
        if !batch.has_more {
            return Ok(());
        }
    }
}

async fn session_cursor(session: &SharedSession) -> UseResult<u64> {
    session_call(session, |session| Ok(session.cursor_sequence())).await
}

async fn session_call<T, F>(session: &SharedSession, operation: F) -> UseResult<T>
where
    T: Send + 'static,
    F: FnOnce(&mut NativeOfficeCollaborationTransportSession) -> UseResult<T> + Send + 'static,
{
    let session = Arc::clone(session);
    spawn_blocking(move || {
        let mut session = session.lock().map_err(|_| {
            collaboration_cli_error(
                "office.collaboration.session_failed",
                "The live collaboration session lock is unavailable.",
            )
        })?;
        operation(&mut session)
    })
    .await
}

#[derive(Debug, Deserialize)]
#[serde(tag = "type", rename_all = "kebab-case", deny_unknown_fields)]
enum SessionInput {
    Reconnect,
    Receive {
        message: Box<SessionMessageInput>,
        #[serde(default, rename = "operationId")]
        operation_id: Option<String>,
        #[serde(default, rename = "ifStateVectorBase64")]
        if_state_vector_base64: Option<String>,
    },
    SetPresence {
        activity: NativeOfficeCollaborationPresenceActivity,
        location: Option<NativeOfficeCollaborationPresenceLocation>,
    },
    ReceiveAwareness {
        message: Box<SessionAwarenessMessageInput>,
    },
    PeerLeft {
        #[serde(rename = "senderClientId")]
        sender_client_id: u64,
    },
    Close,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct SessionMessageInput {
    protocol: String,
    version: u32,
    artifact_id: String,
    artifact_kind: a3s_office::NativeOfficeCollaborationArtifactKind,
    namespace: String,
    sender_client_id: u64,
    #[serde(rename = "type")]
    message_type: NativeOfficeCollaborationTransportMessageType,
    payload_base64: String,
    #[serde(default)]
    origin: Option<NativeOfficeCollaborationOrigin>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct SessionAwarenessMessageInput {
    protocol: String,
    version: u32,
    artifact_id: String,
    artifact_kind: a3s_office::NativeOfficeCollaborationArtifactKind,
    namespace: String,
    sender_client_id: u64,
    payload_base64: String,
}

fn parse_input(line: &[u8]) -> UseResult<SessionInput> {
    serde_json::from_slice(line).map_err(|error| {
        collaboration_cli_error(
            "office.collaboration.session_input_invalid",
            format!("The collaboration session input is not valid JSONL: {error}"),
        )
    })
}

fn receive_request(
    input: SessionMessageInput,
    operation_id: Option<String>,
    if_state_vector_base64: Option<String>,
) -> UseResult<NativeOfficeCollaborationTransportReceiveRequest> {
    let maximum = match input.message_type {
        NativeOfficeCollaborationTransportMessageType::SyncStep1 => {
            MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES
        }
        NativeOfficeCollaborationTransportMessageType::SyncStep2
        | NativeOfficeCollaborationTransportMessageType::Update => {
            MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES
        }
    };
    let payload = decode_bounded_base64(&input.payload_base64, maximum, "message payload")?;
    let if_state_vector = if_state_vector_base64
        .map(|value| {
            decode_bounded_base64(
                &value,
                MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES,
                "state-vector precondition",
            )
        })
        .transpose()?;
    Ok(NativeOfficeCollaborationTransportReceiveRequest {
        message: NativeOfficeCollaborationTransportMessage {
            protocol: input.protocol,
            version: input.version,
            artifact_id: input.artifact_id,
            artifact_kind: input.artifact_kind,
            namespace: input.namespace,
            sender_client_id: input.sender_client_id,
            message_type: input.message_type,
            payload,
            origin: input.origin,
        },
        operation_id,
        if_state_vector,
    })
}

fn awareness_message(
    input: SessionAwarenessMessageInput,
) -> UseResult<NativeOfficeCollaborationAwarenessMessage> {
    Ok(NativeOfficeCollaborationAwarenessMessage {
        protocol: input.protocol,
        version: input.version,
        artifact_id: input.artifact_id,
        artifact_kind: input.artifact_kind,
        namespace: input.namespace,
        sender_client_id: input.sender_client_id,
        payload: decode_bounded_base64(
            &input.payload_base64,
            MAX_NATIVE_OFFICE_COLLABORATION_AWARENESS_BYTES,
            "Awareness payload",
        )?,
    })
}

fn decode_bounded_base64(value: &str, maximum: usize, label: &str) -> UseResult<Vec<u8>> {
    let maximum_encoded = maximum.div_ceil(3) * 4;
    if value.len() > maximum_encoded {
        return Err(collaboration_cli_error(
            "office.collaboration.input_too_large",
            format!("The collaboration session {label} exceeds its encoded size limit."),
        ));
    }
    let decoded = STANDARD.decode(value).map_err(|error| {
        collaboration_cli_error(
            "office.collaboration.session_input_invalid",
            format!("The collaboration session {label} is not valid base64: {error}"),
        )
    })?;
    if decoded.len() > maximum {
        return Err(collaboration_cli_error(
            "office.collaboration.input_too_large",
            format!(
                "The collaboration session {label} is {} bytes; the limit is {maximum} bytes.",
                decoded.len()
            ),
        ));
    }
    Ok(decoded)
}

async fn read_bounded_jsonl<R>(reader: &mut R, output: &mut Vec<u8>) -> UseResult<Option<Vec<u8>>>
where
    R: AsyncBufRead + Unpin,
{
    output.clear();
    loop {
        let available = reader.fill_buf().await.map_err(|error| {
            collaboration_cli_error(
                "office.collaboration.session_input_failed",
                format!("Failed to read collaboration session input: {error}"),
            )
        })?;
        if available.is_empty() {
            if output.is_empty() {
                return Ok(None);
            }
            break;
        }
        let newline = available.iter().position(|byte| *byte == b'\n');
        let consumed = newline.map_or(available.len(), |index| index + 1);
        let content = newline.map_or(available, |index| &available[..index]);
        if output.len().saturating_add(content.len()) > MAX_SESSION_JSONL_BYTES {
            return Err(collaboration_cli_error(
                "office.collaboration.input_too_large",
                format!(
                    "A collaboration session JSONL record exceeds the {MAX_SESSION_JSONL_BYTES}-byte limit."
                ),
            ));
        }
        output.extend_from_slice(content);
        reader.consume(consumed);
        if newline.is_some() {
            break;
        }
    }
    if output.last() == Some(&b'\r') {
        output.pop();
    }
    if output.is_empty() {
        return Err(collaboration_cli_error(
            "office.collaboration.session_input_invalid",
            "Collaboration session JSONL records must not be empty.",
        ));
    }
    Ok(Some(output.clone()))
}

struct SessionOutput {
    stdout: Stdout,
}

impl SessionOutput {
    fn new() -> Self {
        Self {
            stdout: tokio::io::stdout(),
        }
    }

    async fn outbound(
        &mut self,
        reason: &str,
        message: &NativeOfficeCollaborationTransportMessage,
    ) -> UseResult<()> {
        self.write(json!({
            "schemaVersion": 1,
            "type": "outbound",
            "reason": reason,
            "message": message_json(message),
        }))
        .await
    }

    async fn outbound_awareness(
        &mut self,
        reason: &str,
        message: &NativeOfficeCollaborationAwarenessMessage,
    ) -> UseResult<()> {
        self.write(json!({
            "schemaVersion": 1,
            "type": "outbound-awareness",
            "reason": reason,
            "message": awareness_message_json(message),
        }))
        .await
    }

    async fn presence(
        &mut self,
        reason: &str,
        changed: bool,
        snapshot: &NativeOfficeCollaborationPresenceSnapshot,
    ) -> UseResult<()> {
        self.write(json!({
            "schemaVersion": 1,
            "type": "presence",
            "reason": reason,
            "changed": changed,
            "snapshot": snapshot,
        }))
        .await
    }

    async fn received(
        &mut self,
        operation_id: Option<&str>,
        result: &NativeOfficeCollaborationTransportReceiveResult,
    ) -> UseResult<()> {
        self.write(json!({
            "schemaVersion": 1,
            "type": "received",
            "messageType": result.kind,
            "operationId": operation_id,
            "ignored": result.ignored,
            "duplicate": result.apply.as_ref().map(|apply| apply.duplicate),
            "stateChanged": result.apply.as_ref().map(|apply| apply.state_changed),
            "sequence": result.apply.as_ref().and_then(|apply| apply.sequence),
            "stateVectorBase64": result.apply.as_ref().map(|apply| STANDARD.encode(&apply.state_vector)),
        }))
        .await
    }

    async fn poll(
        &mut self,
        batch: &NativeOfficeCollaborationTransportPollResult,
    ) -> UseResult<()> {
        let reason = if batch.resynchronized {
            "history-compacted"
        } else {
            "durable-update"
        };
        for message in &batch.messages {
            self.outbound(reason, message).await?;
        }
        Ok(())
    }

    async fn complete(&mut self, reason: &str, cursor_sequence: u64) -> UseResult<()> {
        self.write(json!({
            "schemaVersion": 1,
            "type": "complete",
            "reason": reason,
            "cursorSequence": cursor_sequence,
        }))
        .await
    }

    async fn error(&mut self, error: &UseError) -> UseResult<()> {
        self.write(json!({
            "schemaVersion": 1,
            "type": "error",
            "error": error,
        }))
        .await
    }

    async fn write(&mut self, value: Value) -> UseResult<()> {
        let text = serde_json::to_vec(&value).map_err(|error| {
            output_error(format!(
                "Failed to serialize collaboration session output: {error}"
            ))
        })?;
        self.stdout.write_all(&text).await.map_err(|error| {
            output_error(format!(
                "Failed to write collaboration session output: {error}"
            ))
        })?;
        self.stdout.write_all(b"\n").await.map_err(|error| {
            output_error(format!(
                "Failed to terminate collaboration session output: {error}"
            ))
        })?;
        self.stdout.flush().await.map_err(|error| {
            output_error(format!(
                "Failed to flush collaboration session output: {error}"
            ))
        })
    }
}

fn message_json(message: &NativeOfficeCollaborationTransportMessage) -> Value {
    let mut value = json!({
        "protocol": message.protocol,
        "version": message.version,
        "artifactId": message.artifact_id,
        "artifactKind": message.artifact_kind,
        "namespace": message.namespace,
        "senderClientId": message.sender_client_id,
        "type": message.message_type,
        "payloadBase64": STANDARD.encode(&message.payload),
    });
    if let Some(origin) = &message.origin {
        value["origin"] = json!(origin);
    }
    value
}

fn awareness_message_json(message: &NativeOfficeCollaborationAwarenessMessage) -> Value {
    json!({
        "protocol": message.protocol,
        "version": message.version,
        "artifactId": message.artifact_id,
        "artifactKind": message.artifact_kind,
        "namespace": message.namespace,
        "senderClientId": message.sender_client_id,
        "payloadBase64": STANDARD.encode(&message.payload),
    })
}

fn output_error(message: impl Into<String>) -> UseError {
    collaboration_cli_error("office.collaboration.session_output_failed", message)
}

async fn timeout_signal(timeout_ms: Option<u64>) {
    match timeout_ms {
        Some(timeout_ms) => tokio::time::sleep(Duration::from_millis(timeout_ms)).await,
        None => std::future::pending().await,
    }
}

async fn termination_signal() {
    #[cfg(unix)]
    {
        if let Ok(mut terminate) =
            tokio::signal::unix::signal(tokio::signal::unix::SignalKind::terminate())
        {
            tokio::select! {
                _ = tokio::signal::ctrl_c() => {}
                _ = terminate.recv() => {}
            }
            return;
        }
    }
    let _ = tokio::signal::ctrl_c().await;
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct SessionArguments {
    store: PathBuf,
    poll_ms: u64,
    timeout_ms: Option<u64>,
    presence_profile: Option<NativeOfficeCollaborationPresenceProfile>,
}

impl SessionArguments {
    fn parse(args: &[String]) -> UseResult<Self> {
        let mut positionals = Vec::new();
        let mut poll_ms = None;
        let mut timeout_ms = None;
        let mut actor_name = None;
        let mut actor_color = None;
        let mut actor_avatar_url = None;
        let mut json = false;
        let mut index = 1;
        while index < args.len() {
            match args[index].as_str() {
                "--poll-ms" => {
                    set_option(&mut poll_ms, args, index, "--poll-ms")?;
                    index += 2;
                }
                "--timeout-ms" => {
                    set_option(&mut timeout_ms, args, index, "--timeout-ms")?;
                    index += 2;
                }
                "--actor-name" => {
                    set_option(&mut actor_name, args, index, "--actor-name")?;
                    index += 2;
                }
                "--actor-color" => {
                    set_option(&mut actor_color, args, index, "--actor-color")?;
                    index += 2;
                }
                "--actor-avatar-url" => {
                    set_option(&mut actor_avatar_url, args, index, "--actor-avatar-url")?;
                    index += 2;
                }
                "--json" => {
                    if json {
                        return Err(usage_error("--json may be specified only once"));
                    }
                    json = true;
                    index += 1;
                }
                "--" => {
                    positionals.extend(args[index + 1..].iter().cloned());
                    break;
                }
                option if option.starts_with('-') => {
                    return Err(usage_error(format!(
                        "Unknown collaboration session option '{option}'."
                    )));
                }
                value => {
                    positionals.push(value.to_owned());
                    index += 1;
                }
            }
        }
        if !json {
            return Err(usage_error(
                "collab session requires --json because stdin and stdout use JSONL",
            ));
        }
        if positionals.len() != 1 {
            return Err(usage_error(
                "collab session requires exactly one collaboration replica path",
            ));
        }
        let poll_ms = parse_optional_u64("--poll-ms", poll_ms)?.unwrap_or(DEFAULT_POLL_MS);
        if !(MIN_POLL_MS..=MAX_POLL_MS).contains(&poll_ms) {
            return Err(usage_error(format!(
                "--poll-ms must be between {MIN_POLL_MS} and {MAX_POLL_MS}"
            )));
        }
        let timeout_ms = parse_optional_u64("--timeout-ms", timeout_ms)?;
        if timeout_ms.is_some_and(|value| value == 0 || value > MAX_TIMEOUT_MS) {
            return Err(usage_error(format!(
                "--timeout-ms must be between 1 and {MAX_TIMEOUT_MS}"
            )));
        }
        if actor_name.is_none() && (actor_color.is_some() || actor_avatar_url.is_some()) {
            return Err(usage_error(
                "--actor-color and --actor-avatar-url require --actor-name",
            ));
        }
        let presence_profile = actor_name.map(|name| NativeOfficeCollaborationPresenceProfile {
            name,
            color: actor_color,
            avatar_url: actor_avatar_url,
        });
        Ok(Self {
            store: PathBuf::from(positionals.remove(0)),
            poll_ms,
            timeout_ms,
            presence_profile,
        })
    }
}

fn set_option(
    target: &mut Option<String>,
    args: &[String],
    index: usize,
    name: &str,
) -> UseResult<()> {
    if target.is_some() {
        return Err(usage_error(format!("{name} may be specified only once")));
    }
    let value = args
        .get(index + 1)
        .filter(|value| !value.starts_with('-'))
        .ok_or_else(|| usage_error(format!("{name} requires a value")))?;
    *target = Some(value.clone());
    Ok(())
}

fn parse_optional_u64(name: &str, value: Option<String>) -> UseResult<Option<u64>> {
    value
        .map(|value| {
            value.parse().map_err(|_| {
                usage_error(format!(
                    "{name} requires a non-negative integer, received '{value}'"
                ))
            })
        })
        .transpose()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn values(values: &[&str]) -> Vec<String> {
        values.iter().map(|value| (*value).to_owned()).collect()
    }

    #[test]
    fn session_arguments_require_jsonl_and_bounded_polling() {
        let parsed = SessionArguments::parse(&values(&[
            "session",
            "agent.replica",
            "--poll-ms",
            "125",
            "--timeout-ms",
            "5000",
            "--json",
        ]))
        .unwrap();
        assert_eq!(parsed.store, PathBuf::from("agent.replica"));
        assert_eq!(parsed.poll_ms, 125);
        assert_eq!(parsed.timeout_ms, Some(5_000));
        assert_eq!(parsed.presence_profile, None);

        let parsed = SessionArguments::parse(&values(&[
            "session",
            "agent.replica",
            "--actor-name",
            "A3S Agent",
            "--actor-color",
            "#2563eb",
            "--json",
        ]))
        .unwrap();
        assert_eq!(
            parsed.presence_profile,
            Some(NativeOfficeCollaborationPresenceProfile {
                name: "A3S Agent".to_owned(),
                color: Some("#2563eb".to_owned()),
                avatar_url: None,
            })
        );

        for invalid in [
            values(&["session", "agent.replica"]),
            values(&["session", "one", "two", "--json"]),
            values(&["session", "one", "--poll-ms", "49", "--json"]),
            values(&["session", "one", "--timeout-ms", "0", "--json"]),
            values(&["session", "one", "--unknown", "--json"]),
            values(&["session", "one", "--actor-color", "#fff", "--json"]),
        ] {
            assert!(SessionArguments::parse(&invalid).is_err(), "{invalid:?}");
        }
    }

    #[tokio::test]
    async fn bounded_jsonl_reader_rejects_empty_records() {
        let mut reader = BufReader::new(&b"\n"[..]);
        let mut buffer = Vec::new();
        let error = read_bounded_jsonl(&mut reader, &mut buffer)
            .await
            .unwrap_err();
        assert_eq!(error.code, "office.collaboration.session_input_invalid");
    }
}
