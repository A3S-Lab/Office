use std::path::PathBuf;
use std::time::Duration;

use a3s_office::{
    NativeOfficeCollaborationEventBatch, NativeOfficeCollaborationEventsRequest,
    NativeOfficeCollaborationInspection, NativeOfficeCollaborationResetEvent,
    NativeOfficeCollaborationStore, NativeOfficeCollaborationUpdateEvent,
    MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH,
};
use a3s_use_core::{UseError, UseResult};
use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use serde_json::{json, Value};
use tokio::io::{AsyncWriteExt, Stdout};

use crate::CommandOutput;

use super::arguments::usage_error;
use super::binary_io::spawn_blocking;

const DEFAULT_POLL_MS: u64 = 250;
const MIN_POLL_MS: u64 = 50;
const MAX_POLL_MS: u64 = 10_000;
const MAX_TIMEOUT_MS: u64 = 24 * 60 * 60 * 1_000;
const MAX_EVENTS: u64 = 100_000;

pub(super) async fn run(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = WatchArguments::parse(args)?;
    let store_path = parsed.store.clone();
    let initial_limit = batch_limit(parsed.max_events, 0);
    let initial_after_sequence = parsed.after_sequence;
    let (store, mut batch, inspection) = spawn_blocking(move || {
        let store = NativeOfficeCollaborationStore::open(store_path)?;
        let batch = store.events(NativeOfficeCollaborationEventsRequest {
            after_sequence: initial_after_sequence,
            limit: initial_limit,
        })?;
        let inspection = store.inspect()?;
        Ok((store, batch, inspection))
    })
    .await?;

    let mut output = WatchOutput::new(parsed.json, parsed.include_updates);
    output.ready(&inspection, &batch, &parsed).await?;

    let timeout = timeout_signal(parsed.timeout_ms);
    tokio::pin!(timeout);
    let termination = termination_signal();
    tokio::pin!(termination);

    let mut cursor_sequence = batch.starting_sequence;
    let mut event_count = 0_u64;
    let (completion_reason, current_sequence) = loop {
        let current_sequence = batch.current_sequence;
        if let Some(reset) = batch.reset.take() {
            output
                .reset(&reset, batch.starting_sequence, batch.checkpoint_sequence)
                .await?;
            cursor_sequence = reset.sequence;
            event_count += 1;
        }
        for event in batch.updates.drain(..) {
            output.update(&event).await?;
            cursor_sequence = event.sequence;
            event_count += 1;
        }

        if parsed
            .max_events
            .is_some_and(|maximum| event_count >= maximum)
        {
            break ("max-events", current_sequence);
        }

        let limit = batch_limit(parsed.max_events, event_count);
        let wait_before_fetch = !batch.has_more;
        let next_store = store.clone();
        let next = async move {
            if wait_before_fetch {
                tokio::time::sleep(Duration::from_millis(parsed.poll_ms)).await;
            }
            spawn_blocking(move || {
                next_store.events(NativeOfficeCollaborationEventsRequest {
                    after_sequence: Some(cursor_sequence),
                    limit,
                })
            })
            .await
        };
        tokio::pin!(next);
        let next_batch = tokio::select! {
            _ = &mut timeout => break ("timeout", current_sequence),
            _ = &mut termination => break ("signal", current_sequence),
            result = &mut next => result?,
        };
        batch = next_batch;
    };

    output
        .complete(
            completion_reason,
            cursor_sequence,
            current_sequence,
            event_count,
        )
        .await?;
    Ok(CommandOutput::silent())
}

fn batch_limit(max_events: Option<u64>, event_count: u64) -> usize {
    max_events
        .map(|maximum| maximum.saturating_sub(event_count))
        .unwrap_or(MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH as u64)
        .min(MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH as u64)
        .max(1) as usize
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

struct WatchOutput {
    stdout: Stdout,
    json: bool,
    include_updates: bool,
}

impl WatchOutput {
    fn new(json: bool, include_updates: bool) -> Self {
        Self {
            stdout: tokio::io::stdout(),
            json,
            include_updates,
        }
    }

    async fn ready(
        &mut self,
        inspection: &NativeOfficeCollaborationInspection,
        batch: &NativeOfficeCollaborationEventBatch,
        arguments: &WatchArguments,
    ) -> UseResult<()> {
        if self.json {
            self.write_json(json!({
                "schemaVersion": 1,
                "type": "ready",
                "store": inspection.store,
                "artifactId": inspection.manifest.artifact_id,
                "artifactKind": inspection.manifest.kind,
                "actorId": inspection.manifest.actor_id,
                "actorKind": inspection.manifest.actor_kind,
                "mode": inspection.manifest.mode,
                "startingSequence": batch.starting_sequence,
                "cursorSequence": batch.starting_sequence,
                "checkpointSequence": batch.checkpoint_sequence,
                "currentSequence": batch.current_sequence,
                "stateVectorBase64": STANDARD.encode(&batch.current_state_vector),
                "stateVectorSha256": batch.current_state_vector_sha256,
                "pollMs": arguments.poll_ms,
                "timeoutMs": arguments.timeout_ms,
                "maxEvents": arguments.max_events,
                "includeUpdates": arguments.include_updates,
            }))
            .await
        } else {
            self.write_text(format!(
                "Watching collaboration replica '{}' for {} '{}' from sequence {}.",
                inspection.store.display(),
                inspection.manifest.kind.as_str(),
                inspection.manifest.artifact_id,
                batch.starting_sequence
            ))
            .await
        }
    }

    async fn reset(
        &mut self,
        event: &NativeOfficeCollaborationResetEvent,
        requested_after_sequence: u64,
        checkpoint_sequence: u64,
    ) -> UseResult<()> {
        if self.json {
            let mut value = json!({
                "schemaVersion": 1,
                "type": "reset",
                "reason": "history-compacted",
                "requestedAfterSequence": requested_after_sequence,
                "checkpointSequence": checkpoint_sequence,
                "sequence": event.sequence,
                "cursorSequence": event.sequence,
                "updateBytes": event.update_bytes,
                "updateSha256": event.update_sha256,
            });
            if self.include_updates {
                value["updateBase64"] = Value::String(STANDARD.encode(&event.update));
            }
            self.write_json(value).await
        } else {
            self.write_text(format!(
                "reset sequence={} checkpoint={} bytes={} reason=history-compacted",
                event.sequence, checkpoint_sequence, event.update_bytes
            ))
            .await
        }
    }

    async fn update(&mut self, event: &NativeOfficeCollaborationUpdateEvent) -> UseResult<()> {
        if self.json {
            let mut value = json!({
                "schemaVersion": 1,
                "type": "update",
                "sequence": event.sequence,
                "cursorSequence": event.sequence,
                "operationId": event.operation_id,
                "operationKind": event.operation_kind,
                "actorId": event.actor_id,
                "actorKind": event.actor_kind,
                "mode": event.mode,
                "artifactId": event.artifact_id,
                "artifactKind": event.artifact_kind,
                "payloadSha256": event.payload_sha256,
                "updateBytes": event.update_bytes,
                "updateSha256": event.update_sha256,
                "beforeStateVectorSha256": event.before_state_vector_sha256,
                "afterStateVectorSha256": event.after_state_vector_sha256,
            });
            if let Some(origin) = &event.origin {
                value["origin"] = json!(origin);
            }
            if self.include_updates {
                value["updateBase64"] = Value::String(STANDARD.encode(&event.update));
            }
            self.write_json(value).await
        } else {
            self.write_text(format!(
                "update sequence={} operation={} actor={} bytes={}",
                event.sequence, event.operation_id, event.actor_id, event.update_bytes
            ))
            .await
        }
    }

    async fn complete(
        &mut self,
        reason: &str,
        cursor_sequence: u64,
        current_sequence: u64,
        event_count: u64,
    ) -> UseResult<()> {
        if self.json {
            self.write_json(json!({
                "schemaVersion": 1,
                "type": "complete",
                "reason": reason,
                "cursorSequence": cursor_sequence,
                "currentSequence": current_sequence,
                "eventCount": event_count,
            }))
            .await
        } else {
            self.write_text(format!(
                "Stopped collaboration watch: reason={reason} cursor={cursor_sequence} events={event_count}."
            ))
            .await
        }
    }

    async fn write_json(&mut self, value: Value) -> UseResult<()> {
        let text = serde_json::to_string(&value).map_err(|error| {
            output_error(format!(
                "Failed to serialize collaboration watch output: {error}"
            ))
        })?;
        self.write_text(text).await
    }

    async fn write_text(&mut self, text: String) -> UseResult<()> {
        self.stdout
            .write_all(text.as_bytes())
            .await
            .map_err(|error| {
                output_error(format!(
                    "Failed to write collaboration watch output: {error}"
                ))
            })?;
        self.stdout.write_all(b"\n").await.map_err(|error| {
            output_error(format!(
                "Failed to write collaboration watch output: {error}"
            ))
        })?;
        self.stdout.flush().await.map_err(|error| {
            output_error(format!(
                "Failed to flush collaboration watch output: {error}"
            ))
        })
    }
}

fn output_error(message: impl Into<String>) -> UseError {
    UseError::new("office.collaboration.watch_output_failed", message)
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct WatchArguments {
    store: PathBuf,
    after_sequence: Option<u64>,
    poll_ms: u64,
    timeout_ms: Option<u64>,
    max_events: Option<u64>,
    include_updates: bool,
    json: bool,
}

impl WatchArguments {
    fn parse(args: &[String]) -> UseResult<Self> {
        let mut positionals = Vec::new();
        let mut after_sequence = None;
        let mut poll_ms = None;
        let mut timeout_ms = None;
        let mut max_events = None;
        let mut include_updates = false;
        let mut json = false;
        let mut index = 1;
        while index < args.len() {
            match args[index].as_str() {
                "--after-sequence" => {
                    set_option(&mut after_sequence, args, index, "--after-sequence")?;
                    index += 2;
                }
                "--poll-ms" => {
                    set_option(&mut poll_ms, args, index, "--poll-ms")?;
                    index += 2;
                }
                "--timeout-ms" => {
                    set_option(&mut timeout_ms, args, index, "--timeout-ms")?;
                    index += 2;
                }
                "--max-events" => {
                    set_option(&mut max_events, args, index, "--max-events")?;
                    index += 2;
                }
                "--include-updates" => {
                    set_flag(&mut include_updates, "--include-updates")?;
                    index += 1;
                }
                "--json" => {
                    set_flag(&mut json, "--json")?;
                    index += 1;
                }
                "--" => {
                    positionals.extend(args[index + 1..].iter().cloned());
                    break;
                }
                option if option.starts_with('-') => {
                    return Err(usage_error(format!(
                        "Unknown collaboration watch option '{option}'."
                    )));
                }
                value => {
                    positionals.push(value.to_owned());
                    index += 1;
                }
            }
        }
        if positionals.len() != 1 {
            return Err(usage_error(
                "collab watch requires exactly one collaboration replica path",
            ));
        }

        let after_sequence = parse_optional_u64("--after-sequence", after_sequence)?;
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
        let max_events = parse_optional_u64("--max-events", max_events)?;
        if max_events.is_some_and(|value| value == 0 || value > MAX_EVENTS) {
            return Err(usage_error(format!(
                "--max-events must be between 1 and {MAX_EVENTS}"
            )));
        }

        Ok(Self {
            store: PathBuf::from(positionals.remove(0)),
            after_sequence,
            poll_ms,
            timeout_ms,
            max_events,
            include_updates,
            json,
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

fn set_flag(target: &mut bool, name: &str) -> UseResult<()> {
    if *target {
        return Err(usage_error(format!("{name} may be specified only once")));
    }
    *target = true;
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
    fn watch_arguments_are_explicit_and_bounded() {
        let parsed = WatchArguments::parse(&values(&[
            "watch",
            "agent.replica",
            "--after-sequence",
            "41",
            "--poll-ms",
            "125",
            "--timeout-ms",
            "5000",
            "--max-events",
            "7",
            "--include-updates",
            "--json",
        ]))
        .unwrap();
        assert_eq!(parsed.store, PathBuf::from("agent.replica"));
        assert_eq!(parsed.after_sequence, Some(41));
        assert_eq!(parsed.poll_ms, 125);
        assert_eq!(parsed.timeout_ms, Some(5_000));
        assert_eq!(parsed.max_events, Some(7));
        assert!(parsed.include_updates);
        assert!(parsed.json);

        for invalid in [
            values(&["watch"]),
            values(&["watch", "one", "two"]),
            values(&["watch", "one", "--poll-ms", "49"]),
            values(&["watch", "one", "--timeout-ms", "0"]),
            values(&["watch", "one", "--max-events", "100001"]),
            values(&["watch", "one", "--include-updates", "--include-updates"]),
            values(&["watch", "one", "--unknown"]),
        ] {
            assert!(WatchArguments::parse(&invalid).is_err(), "{invalid:?}");
        }
    }
}
