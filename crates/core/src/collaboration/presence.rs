use std::sync::atomic::{AtomicU64, Ordering};

use a3s_use_core::UseResult;
use serde::{Deserialize, Serialize};
use yrs::sync::{Awareness, AwarenessUpdate};
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{ClientID, Doc};

use super::{
    collaboration_error, validate_client_id, NativeOfficeCollaborationAwarenessMessage,
    NativeOfficeCollaborationManifest, NativeOfficeCollaborationParticipant,
    NativeOfficeCollaborationPresenceProfile, NativeOfficeCollaborationPresenceReceiveResult,
    NativeOfficeCollaborationPresenceSnapshot, NativeOfficeCollaborationPresenceState,
    NativeOfficeCollaborationPresenceUpdate, MAX_NATIVE_OFFICE_COLLABORATION_AWARENESS_BYTES,
    NATIVE_OFFICE_COLLABORATION_PROTOCOL, NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
};

mod validation;

use validation::{validate_location, validate_manifest, validate_presence_state, validate_profile};

static NEXT_PRESENCE_SEQUENCE: AtomicU64 = AtomicU64::new(0);

/// An ephemeral Yjs Awareness peer for one native Office replica.
///
/// This state machine deliberately owns a separate in-memory Yrs document. It
/// never reads or writes the durable collaboration store, so presence,
/// selection, and activity cannot leak into canonical Office content or
/// operation history.
#[derive(Debug)]
pub struct NativeOfficeCollaborationPresenceSession {
    awareness: Awareness,
    manifest: NativeOfficeCollaborationManifest,
    client_id: u64,
    local_state: NativeOfficeCollaborationPresenceState,
    disconnected: bool,
}

impl NativeOfficeCollaborationPresenceSession {
    pub fn new(
        manifest: NativeOfficeCollaborationManifest,
        profile: NativeOfficeCollaborationPresenceProfile,
    ) -> UseResult<Self> {
        let client_id = manifest.client_id;
        Self::new_with_sender_client_id(manifest, client_id, profile)
    }

    /// Create Presence for an ephemeral host connection identity while the
    /// durable replica retains its stable Yrs client ID.
    pub fn new_with_sender_client_id(
        manifest: NativeOfficeCollaborationManifest,
        client_id: u64,
        profile: NativeOfficeCollaborationPresenceProfile,
    ) -> UseResult<Self> {
        validate_manifest(&manifest)?;
        validate_client_id(client_id)?;
        validate_profile(&profile)?;
        let local_state = NativeOfficeCollaborationPresenceState {
            protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
            version: NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
            artifact_id: manifest.artifact_id.clone(),
            artifact_kind: manifest.kind,
            namespace: manifest.namespace.clone(),
            presence_id: next_presence_id(client_id),
            actor: super::NativeOfficeCollaborationPresenceActor {
                id: manifest.actor_id.clone(),
                name: profile.name,
                color: profile.color,
                avatar_url: profile.avatar_url,
                kind: manifest.actor_kind,
            },
            mode: manifest.mode,
            activity: super::NativeOfficeCollaborationPresenceActivity::Active,
            location: None,
        };
        validate_presence_state(&local_state, &manifest)?;
        let mut awareness = Awareness::new(Doc::with_client_id(client_id));
        set_local_state(&mut awareness, &local_state)?;
        Ok(Self {
            awareness,
            manifest,
            client_id,
            local_state,
            disconnected: false,
        })
    }

    pub fn local_state(&self) -> UseResult<&NativeOfficeCollaborationPresenceState> {
        self.ensure_connected()?;
        Ok(&self.local_state)
    }

    pub fn local_message(&self) -> UseResult<NativeOfficeCollaborationAwarenessMessage> {
        self.ensure_connected()?;
        self.encoded_local_message()
    }

    pub fn update(
        &mut self,
        update: NativeOfficeCollaborationPresenceUpdate,
    ) -> UseResult<NativeOfficeCollaborationAwarenessMessage> {
        self.ensure_connected()?;
        if let Some(location) = &update.location {
            validate_location(location, self.manifest.kind)?;
        }
        let mut next = self.local_state.clone();
        next.activity = update.activity;
        next.location = update.location;
        set_local_state(&mut self.awareness, &next)?;
        self.local_state = next;
        self.encoded_local_message()
    }

    pub fn receive(
        &mut self,
        message: NativeOfficeCollaborationAwarenessMessage,
    ) -> UseResult<NativeOfficeCollaborationPresenceReceiveResult> {
        self.ensure_connected()?;
        self.validate_message_identity(&message)?;
        if message.sender_client_id == self.client_id {
            return Ok(NativeOfficeCollaborationPresenceReceiveResult {
                ignored: true,
                changed: false,
                snapshot: self.snapshot()?,
            });
        }
        if message.payload.len() > MAX_NATIVE_OFFICE_COLLABORATION_AWARENESS_BYTES {
            return Err(collaboration_error(
                "office.collaboration.presence_message_invalid",
                format!(
                    "The Awareness payload is {} bytes; the limit is {} bytes.",
                    message.payload.len(),
                    MAX_NATIVE_OFFICE_COLLABORATION_AWARENESS_BYTES
                ),
            ));
        }
        let update = AwarenessUpdate::decode_v1(&message.payload).map_err(|error| {
            collaboration_error(
                "office.collaboration.presence_message_invalid",
                format!("The Awareness payload is not a valid Yjs v1 update: {error}"),
            )
        })?;
        if update.clients.len() != 1 {
            return Err(collaboration_error(
                "office.collaboration.presence_message_invalid",
                "An Awareness message must update exactly one sender client.",
            ));
        }
        let Some((client_id, entry)) = update.clients.iter().next() else {
            return Err(collaboration_error(
                "office.collaboration.presence_message_invalid",
                "An Awareness message must update exactly one sender client.",
            ));
        };
        if client_id.get() != message.sender_client_id {
            return Err(collaboration_error(
                "office.collaboration.presence_message_invalid",
                "The Awareness payload client ID does not match its sender envelope.",
            ));
        }
        if entry.json.as_ref() != "null" {
            let envelope: OfficeAwarenessState = serde_json::from_str(entry.json.as_ref())
                .map_err(|error| {
                    collaboration_error(
                        "office.collaboration.presence_message_invalid",
                        format!("The Awareness state is not valid Office presence JSON: {error}"),
                    )
                })?;
            validate_presence_state(&envelope.a3s_office, &self.manifest)?;
        }
        let changed = self
            .awareness
            .apply_update_summary(update)
            .map_err(|error| {
                collaboration_error(
                    "office.collaboration.presence_message_invalid",
                    format!("The Awareness update could not be applied: {error}"),
                )
            })?
            .is_some();
        Ok(NativeOfficeCollaborationPresenceReceiveResult {
            ignored: false,
            changed,
            snapshot: self.snapshot()?,
        })
    }

    pub fn peer_left(
        &mut self,
        sender_client_id: u64,
    ) -> UseResult<NativeOfficeCollaborationPresenceReceiveResult> {
        self.ensure_connected()?;
        validate_client_id(sender_client_id)?;
        if sender_client_id == self.client_id {
            return Err(collaboration_error(
                "office.collaboration.presence_message_invalid",
                "A remote peer-left event cannot remove the local Presence participant.",
            ));
        }
        let client_id = ClientID::new(sender_client_id);
        let changed = self
            .awareness
            .state::<serde_json::Value>(client_id)
            .is_some();
        if changed {
            self.awareness.remove_state(client_id);
        }
        Ok(NativeOfficeCollaborationPresenceReceiveResult {
            ignored: false,
            changed,
            snapshot: self.snapshot()?,
        })
    }

    pub fn clear_remote(&mut self) -> UseResult<NativeOfficeCollaborationPresenceReceiveResult> {
        self.ensure_connected()?;
        let remote = self
            .awareness
            .iter()
            .filter_map(|(client_id, state)| {
                (client_id.get() != self.client_id && state.data.is_some()).then_some(client_id)
            })
            .collect::<Vec<_>>();
        for client_id in &remote {
            self.awareness.remove_state(*client_id);
        }
        Ok(NativeOfficeCollaborationPresenceReceiveResult {
            ignored: false,
            changed: !remote.is_empty(),
            snapshot: self.snapshot()?,
        })
    }

    pub fn snapshot(&self) -> UseResult<NativeOfficeCollaborationPresenceSnapshot> {
        self.ensure_connected()?;
        let mut participants = Vec::new();
        for (client_id, client_state) in self.awareness.iter() {
            let Some(json) = client_state.data else {
                continue;
            };
            let envelope: OfficeAwarenessState =
                serde_json::from_str(json.as_ref()).map_err(|error| {
                    collaboration_error(
                        "office.collaboration.presence_message_invalid",
                        format!("An applied Awareness state is not valid Office presence: {error}"),
                    )
                })?;
            validate_presence_state(&envelope.a3s_office, &self.manifest)?;
            participants.push(NativeOfficeCollaborationParticipant {
                client_id: client_id.get(),
                local: client_id.get() == self.client_id,
                state: envelope.a3s_office,
            });
        }
        participants.sort_by_key(|participant| participant.client_id);
        Ok(NativeOfficeCollaborationPresenceSnapshot {
            local_client_id: self.client_id,
            participants,
        })
    }

    /// Clear and encode the local Awareness state for an orderly disconnect.
    /// The returned tombstone is ephemeral and must be forwarded before the
    /// host closes its room transport.
    pub fn disconnect(&mut self) -> UseResult<NativeOfficeCollaborationAwarenessMessage> {
        self.ensure_connected()?;
        self.awareness.clean_local_state();
        let message = self.encoded_local_message()?;
        self.disconnected = true;
        Ok(message)
    }

    fn encoded_local_message(&self) -> UseResult<NativeOfficeCollaborationAwarenessMessage> {
        let update = self
            .awareness
            .update_with_clients([ClientID::new(self.client_id)])
            .map_err(|error| {
                collaboration_error(
                    "office.collaboration.presence_message_invalid",
                    format!("The local Awareness update could not be encoded: {error}"),
                )
            })?;
        Ok(NativeOfficeCollaborationAwarenessMessage {
            protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
            version: NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
            artifact_id: self.manifest.artifact_id.clone(),
            artifact_kind: self.manifest.kind,
            namespace: self.manifest.namespace.clone(),
            sender_client_id: self.client_id,
            payload: update.encode_v1(),
        })
    }

    fn validate_message_identity(
        &self,
        message: &NativeOfficeCollaborationAwarenessMessage,
    ) -> UseResult<()> {
        if message.protocol != NATIVE_OFFICE_COLLABORATION_PROTOCOL
            || message.version != NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION
        {
            return Err(collaboration_error(
                "office.collaboration.presence_message_invalid",
                "The Office Awareness protocol or version is unsupported.",
            ));
        }
        if message.artifact_id != self.manifest.artifact_id
            || message.artifact_kind != self.manifest.kind
            || message.namespace != self.manifest.namespace
        {
            return Err(collaboration_error(
                "office.collaboration.presence_identity_mismatch",
                "The Office Awareness message belongs to another artifact, kind, or namespace.",
            ));
        }
        validate_client_id(message.sender_client_id)
    }

    fn ensure_connected(&self) -> UseResult<()> {
        if !self.disconnected {
            return Ok(());
        }
        Err(collaboration_error(
            "office.collaboration.presence_disconnected",
            "The native Office Presence session has already disconnected.",
        ))
    }
}

#[derive(Debug, Serialize, Deserialize)]
struct OfficeAwarenessState {
    #[serde(rename = "a3sOffice")]
    a3s_office: NativeOfficeCollaborationPresenceState,
}

fn set_local_state(
    awareness: &mut Awareness,
    state: &NativeOfficeCollaborationPresenceState,
) -> UseResult<()> {
    awareness
        .set_local_state(OfficeAwarenessState {
            a3s_office: state.clone(),
        })
        .map_err(|error| {
            collaboration_error(
                "office.collaboration.presence_invalid",
                format!("The local Office Presence state could not be encoded: {error}"),
            )
        })
}

fn next_presence_id(client_id: u64) -> String {
    let sequence = NEXT_PRESENCE_SEQUENCE.fetch_add(1, Ordering::Relaxed) + 1;
    format!("{client_id}:native:{}:{sequence:x}", std::process::id())
}
