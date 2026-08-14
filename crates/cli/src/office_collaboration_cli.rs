mod arguments;
mod binary_io;
mod watch;

use std::path::PathBuf;

use a3s_office::{
    NativeOfficeCollaborationActorKind, NativeOfficeCollaborationApplyRequest,
    NativeOfficeCollaborationArtifactKind, NativeOfficeCollaborationCheckpointRequest,
    NativeOfficeCollaborationCreateRequest, NativeOfficeCollaborationMode,
    NativeOfficeCollaborationStore, MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES,
    MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES,
};
use a3s_use_core::UseResult;
use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use serde::Serialize;
use serde_json::{json, Value};

use crate::CommandOutput;
use arguments::{collaboration_cli_error, parse_value, usage_error, ParsedOptions};
use binary_io::{read_binary_input, read_optional_encoded_input, spawn_blocking, write_new_binary};

const HELP: &str = concat!(
    "a3s-office collab - native Yjs/Yrs collaboration replica\n\n",
    /*
    "a3s-office collab — native Yjs/Yrs collaboration replica\n\n",
    */
    "usage:\n",
    "  a3s-office collab create <store> --artifact-id <id> --kind <kind> --actor-id <id> --operation-id <id> [--actor-kind human|agent|system] [--mode view|comment|suggest|edit] [--namespace <name>] [--client-id <u53>] [--json]\n",
    "  a3s-office collab join <store> --artifact-id <id> --kind <kind> --actor-id <id> --operation-id <id> --input <update.bin> [create options] [--json]\n",
    "  a3s-office collab inspect <store> [--json]\n",
    "  a3s-office collab diff <store> [--state-vector <base64>|--state-vector-input <file>] [--output <update.bin>] [--json]\n",
    "  a3s-office collab sync-step1 <store> [--output <message.bin>] [--json]\n",
    "  a3s-office collab encode-update --input <update.bin> [--output <message.bin>] [--json]\n",
    "  a3s-office collab handle-message <store> --input <message.bin> [--output <response.bin>|mutation identity options] [--json]\n",
    "  a3s-office collab apply <store> --input <update.bin> --actor-id <id> --operation-id <id> --artifact-id <id> --kind <kind> --mode <mode> [--if-state-vector <base64>|--if-state-vector-input <file>] [--json]\n",
    "  a3s-office collab watch <store> [--after-sequence <u64>] [--poll-ms <50..10000>] [--timeout-ms <u64>] [--max-events <u64>] [--include-updates] [--json]\n",
    "  a3s-office collab checkpoint <store> --actor-id <id> --operation-id <id> --artifact-id <id> --kind <kind> --mode <mode> [--if-state-vector <base64>|--if-state-vector-input <file>] [--json]\n",
    "  a3s-office collab leave <store> --actor-id <id> --operation-id <id> --artifact-id <id> --kind <kind> --mode <mode> [--if-state-vector <base64>|--if-state-vector-input <file>] [--json]\n\n",
    "Kinds: document, markdown, spreadsheet, presentation, pdf.\n",
    "Binary updates and state vectors use the standard Yjs v1 encoding. Output paths are no-clobber."
);

pub(crate) async fn run(args: &[String]) -> UseResult<CommandOutput> {
    if matches!(args.first().map(String::as_str), Some("watch")) {
        return watch::run(args).await;
    }
    let filtered = args
        .iter()
        .filter(|argument| argument.as_str() != "--json")
        .cloned()
        .collect::<Vec<_>>();
    match filtered.first().map(String::as_str) {
        None | Some("help" | "--help" | "-h") => Ok(help()),
        Some("create") => create(&filtered, false).await,
        Some("join") => create(&filtered, true).await,
        Some("inspect") => inspect(&filtered).await,
        Some("diff" | "synchronize" | "sync") => diff(&filtered).await,
        Some("sync-step1") => sync_step1(&filtered).await,
        Some("encode-update") => encode_update(&filtered).await,
        Some("handle-message") => handle_message(&filtered).await,
        Some("apply") => apply(&filtered).await,
        Some("checkpoint") => checkpoint(&filtered, false).await,
        Some("leave") => checkpoint(&filtered, true).await,
        Some(command) => Err(usage_error(format!(
            "Unknown collaboration command '{command}'."
        ))),
    }
}

async fn create(args: &[String], join: bool) -> UseResult<CommandOutput> {
    let parsed = ParsedOptions::parse(args)?;
    parsed.reject_unknown(&[
        "actor-id",
        "actor-kind",
        "artifact-id",
        "client-id",
        "input",
        "kind",
        "mode",
        "namespace",
        "operation-id",
    ])?;
    if join && parsed.value("input")?.is_none() {
        return Err(usage_error("collab join requires --input <update.bin>"));
    }
    if !join && parsed.value("input")?.is_some() {
        return Err(usage_error(
            "collab create does not accept --input; use collab join",
        ));
    }
    let store_path = parsed.one_positional("collaboration replica path")?;
    let artifact_id = parsed.required("artifact-id")?.to_owned();
    let kind =
        parse_value::<NativeOfficeCollaborationArtifactKind>(parsed.required("kind")?, "--kind")?;
    let actor_id = parsed.required("actor-id")?.to_owned();
    let actor_kind = parse_value::<NativeOfficeCollaborationActorKind>(
        parsed.value("actor-kind")?.unwrap_or("agent"),
        "--actor-kind",
    )?;
    let mode = parse_value::<NativeOfficeCollaborationMode>(
        parsed.value("mode")?.unwrap_or("edit"),
        "--mode",
    )?;
    let operation_id = parsed.required("operation-id")?.to_owned();
    let namespace = parsed.value("namespace")?.map(str::to_owned);
    let client_id = parsed
        .value("client-id")?
        .map(|value| {
            value.parse::<u64>().map_err(|_| {
                usage_error(format!(
                    "--client-id requires a non-zero 53-bit integer, received '{value}'"
                ))
            })
        })
        .transpose()?;
    let initial_update = match parsed.value("input")? {
        Some(path) => Some(
            read_binary_input(
                &PathBuf::from(path),
                MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES,
                "update",
            )
            .await?,
        ),
        None => None,
    };
    let store = spawn_blocking(move || {
        NativeOfficeCollaborationStore::create(NativeOfficeCollaborationCreateRequest {
            store: PathBuf::from(store_path),
            artifact_id,
            kind,
            actor_id,
            actor_kind,
            mode,
            operation_id,
            namespace,
            client_id,
            initial_update,
        })
    })
    .await?;
    let inspection = spawn_blocking(move || store.inspect()).await?;
    let data = inspection_json(&inspection)?;
    Ok(CommandOutput::success(
        format!(
            "{} collaboration replica '{}' for {} artifact '{}'.",
            if join { "Joined" } else { "Created" },
            inspection.store.display(),
            inspection.manifest.kind.as_str(),
            inspection.manifest.artifact_id
        ),
        json!({
            "action": if join { "joined" } else { "created" },
            "replica": data,
        }),
    ))
}

async fn inspect(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedOptions::parse(args)?;
    parsed.reject_unknown(&[])?;
    let store_path = parsed.one_positional("collaboration replica path")?;
    let inspection =
        spawn_blocking(move || NativeOfficeCollaborationStore::open(store_path)?.inspect()).await?;
    let data = inspection_json(&inspection)?;
    Ok(CommandOutput::success(
        format!(
            "Collaboration replica '{}' is at sequence {} with {} pending update(s).",
            inspection.store.display(),
            inspection.current_sequence,
            inspection.update_count
        ),
        data,
    ))
}

async fn diff(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedOptions::parse(args)?;
    parsed.reject_unknown(&["output", "state-vector", "state-vector-input"])?;
    let store_path = parsed.one_positional("collaboration replica path")?;
    let remote_state_vector = read_optional_encoded_input(
        parsed.value("state-vector")?,
        parsed.value("state-vector-input")?,
        MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES,
        "state vector",
    )
    .await?;
    let result = spawn_blocking(move || {
        NativeOfficeCollaborationStore::open(store_path)?
            .synchronize(remote_state_vector.as_deref())
    })
    .await?;
    let output = parsed.value("output")?.map(PathBuf::from);
    if let Some(path) = &output {
        write_new_binary(path.clone(), result.update.clone(), "collaboration update").await?;
    }
    let update_base64 = output.is_none().then(|| STANDARD.encode(&result.update));
    Ok(CommandOutput::success(
        match &output {
            Some(path) => format!(
                "Wrote {}-byte collaboration diff to '{}'.",
                result.update.len(),
                path.display()
            ),
            None => format!("Encoded a {}-byte collaboration diff.", result.update.len()),
        },
        json!({
            "action": "diff",
            "output": output,
            "updateBase64": update_base64,
            "updateBytes": result.update.len(),
            "updateSha256": result.update_sha256,
            "stateVectorBase64": STANDARD.encode(&result.state_vector),
            "stateVectorSha256": result.state_vector_sha256,
        }),
    ))
}

async fn apply(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedOptions::parse(args)?;
    parsed.reject_unknown(&[
        "actor-id",
        "artifact-id",
        "if-state-vector",
        "if-state-vector-input",
        "input",
        "kind",
        "mode",
        "operation-id",
    ])?;
    let store_path = parsed.one_positional("collaboration replica path")?;
    let input = PathBuf::from(parsed.required("input")?);
    let update = read_binary_input(
        &input,
        MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES,
        "update",
    )
    .await?;
    let request = mutation_request(&parsed, update).await?;
    let result =
        spawn_blocking(move || NativeOfficeCollaborationStore::open(store_path)?.apply(request))
            .await?;
    let value = value_with_base64(&result, &result.state_vector)?;
    Ok(CommandOutput::success(
        format!(
            "{} operation '{}' at sequence {}{}.",
            if result.duplicate {
                "Replayed"
            } else {
                "Applied"
            },
            result.operation_id,
            result
                .sequence
                .map(|value| value.to_string())
                .unwrap_or_else(|| "unchanged".to_owned()),
            if result.checkpointed {
                " and checkpointed"
            } else {
                ""
            }
        ),
        value,
    ))
}

async fn sync_step1(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedOptions::parse(args)?;
    parsed.reject_unknown(&["output"])?;
    let store_path = parsed.one_positional("collaboration replica path")?;
    let result =
        spawn_blocking(move || NativeOfficeCollaborationStore::open(store_path)?.sync_step1())
            .await?;
    let output = parsed.value("output")?.map(PathBuf::from);
    if let Some(path) = &output {
        write_new_binary(path.clone(), result.message.clone(), "y-sync message").await?;
    }
    Ok(CommandOutput::success(
        match &output {
            Some(path) => format!(
                "Wrote {}-byte y-sync SyncStep1 message to '{}'.",
                result.message.len(),
                path.display()
            ),
            None => format!(
                "Encoded a {}-byte y-sync SyncStep1 message.",
                result.message.len()
            ),
        },
        json!({
            "action": "sync-step1",
            "messageBase64": output.is_none().then(|| STANDARD.encode(&result.message)),
            "messageBytes": result.message.len(),
            "messageSha256": result.message_sha256,
            "output": output,
            "stateVectorBase64": STANDARD.encode(&result.state_vector),
            "stateVectorSha256": result.state_vector_sha256,
        }),
    ))
}

async fn encode_update(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedOptions::parse(args)?;
    parsed.reject_unknown(&["input", "output"])?;
    if !parsed.positionals.is_empty() {
        return Err(usage_error(
            "collab encode-update does not accept a replica path",
        ));
    }
    let input = PathBuf::from(parsed.required("input")?);
    let update = read_binary_input(
        &input,
        MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES,
        "update",
    )
    .await?;
    let message =
        spawn_blocking(move || NativeOfficeCollaborationStore::encode_sync_update(&update)).await?;
    let output = parsed.value("output")?.map(PathBuf::from);
    if let Some(path) = &output {
        write_new_binary(path.clone(), message.clone(), "y-sync message").await?;
    }
    Ok(CommandOutput::success(
        match &output {
            Some(path) => format!(
                "Wrote {}-byte y-sync Update message to '{}'.",
                message.len(),
                path.display()
            ),
            None => format!("Encoded a {}-byte y-sync Update message.", message.len()),
        },
        json!({
            "action": "encode-update",
            "messageBase64": output.is_none().then(|| STANDARD.encode(&message)),
            "messageBytes": message.len(),
            "output": output,
        }),
    ))
}

async fn checkpoint(args: &[String], leave: bool) -> UseResult<CommandOutput> {
    let parsed = ParsedOptions::parse(args)?;
    parsed.reject_unknown(&[
        "actor-id",
        "artifact-id",
        "if-state-vector",
        "if-state-vector-input",
        "kind",
        "mode",
        "operation-id",
    ])?;
    let store_path = parsed.one_positional("collaboration replica path")?;
    let request = checkpoint_request(&parsed).await?;
    let result = spawn_blocking(move || {
        let store = NativeOfficeCollaborationStore::open(store_path)?;
        if leave {
            store.leave(request)
        } else {
            store.checkpoint(request)
        }
    })
    .await?;
    let mut value = value_with_base64(&result, &result.state_vector)?;
    value["action"] = Value::String(if leave { "left" } else { "checkpointed" }.to_owned());
    Ok(CommandOutput::success(
        format!(
            "{} collaboration replica at sequence {}; compacted {} update(s).",
            if leave { "Left" } else { "Checkpointed" },
            result.sequence,
            result.compacted_updates
        ),
        value,
    ))
}

async fn handle_message(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = ParsedOptions::parse(args)?;
    parsed.reject_unknown(&[
        "actor-id",
        "artifact-id",
        "if-state-vector",
        "if-state-vector-input",
        "input",
        "kind",
        "mode",
        "operation-id",
        "output",
    ])?;
    let store_path = parsed.one_positional("collaboration replica path")?;
    let message = read_binary_input(
        &PathBuf::from(parsed.required("input")?),
        MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES + 32,
        "y-sync message",
    )
    .await?;
    let mutation = if parsed.value("operation-id")?.is_some() {
        Some(mutation_request(&parsed, Vec::new()).await?)
    } else {
        for option in [
            "actor-id",
            "artifact-id",
            "if-state-vector",
            "if-state-vector-input",
            "kind",
            "mode",
        ] {
            if parsed.value(option)?.is_some() {
                return Err(usage_error(format!(
                    "--{option} requires --operation-id for a mutating y-sync message"
                )));
            }
        }
        None
    };
    let result = spawn_blocking(move || {
        NativeOfficeCollaborationStore::open(store_path)?.handle_sync_message(&message, mutation)
    })
    .await?;
    let output = parsed.value("output")?.map(PathBuf::from);
    if let Some(path) = &output {
        let response = result.response.clone().ok_or_else(|| {
            usage_error("--output requires a y-sync message that produces a response")
        })?;
        write_new_binary(path.clone(), response, "y-sync response").await?;
    }
    let mut value = serde_json::to_value(&result).map_err(|error| {
        collaboration_cli_error(
            "office.collaboration.output_failed",
            format!("Failed to encode y-sync output: {error}"),
        )
    })?;
    value["action"] = Value::String("handle-message".to_owned());
    value["output"] = serde_json::to_value(&output).unwrap_or(Value::Null);
    value["responseBase64"] = result
        .response
        .as_ref()
        .filter(|_| output.is_none())
        .map(|response| Value::String(STANDARD.encode(response)))
        .unwrap_or(Value::Null);
    value["stateVector"] = Value::String(STANDARD.encode(&result.state_vector));
    if let Some(apply) = &result.apply {
        value["apply"]["stateVector"] = Value::String(STANDARD.encode(&apply.state_vector));
    }
    Ok(CommandOutput::success(
        match &output {
            Some(path) => format!(
                "Handled y-sync {:?} and wrote its response to '{}'.",
                result.kind,
                path.display()
            ),
            None => format!("Handled y-sync {:?}.", result.kind),
        },
        value,
    ))
}

async fn mutation_request(
    parsed: &ParsedOptions,
    update: Vec<u8>,
) -> UseResult<NativeOfficeCollaborationApplyRequest> {
    Ok(NativeOfficeCollaborationApplyRequest {
        operation_id: parsed.required("operation-id")?.to_owned(),
        actor_id: parsed.required("actor-id")?.to_owned(),
        mode: parse_value(parsed.required("mode")?, "--mode")?,
        expected_artifact_id: parsed.required("artifact-id")?.to_owned(),
        expected_kind: parse_value(parsed.required("kind")?, "--kind")?,
        update,
        if_state_vector: read_optional_encoded_input(
            parsed.value("if-state-vector")?,
            parsed.value("if-state-vector-input")?,
            MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES,
            "state-vector precondition",
        )
        .await?,
    })
}

async fn checkpoint_request(
    parsed: &ParsedOptions,
) -> UseResult<NativeOfficeCollaborationCheckpointRequest> {
    Ok(NativeOfficeCollaborationCheckpointRequest {
        operation_id: parsed.required("operation-id")?.to_owned(),
        actor_id: parsed.required("actor-id")?.to_owned(),
        mode: parse_value(parsed.required("mode")?, "--mode")?,
        expected_artifact_id: parsed.required("artifact-id")?.to_owned(),
        expected_kind: parse_value(parsed.required("kind")?, "--kind")?,
        if_state_vector: read_optional_encoded_input(
            parsed.value("if-state-vector")?,
            parsed.value("if-state-vector-input")?,
            MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES,
            "state-vector precondition",
        )
        .await?,
    })
}

fn inspection_json(
    inspection: &a3s_office::NativeOfficeCollaborationInspection,
) -> UseResult<Value> {
    value_with_base64(inspection, &inspection.state_vector)
}

fn value_with_base64(value: &impl Serialize, state_vector: &[u8]) -> UseResult<Value> {
    let mut value = serde_json::to_value(value).map_err(|error| {
        collaboration_cli_error(
            "office.collaboration.output_failed",
            format!("Failed to encode collaboration output: {error}"),
        )
    })?;
    value["stateVector"] = Value::String(STANDARD.encode(state_vector));
    Ok(value)
}

fn help() -> CommandOutput {
    CommandOutput::success(
        HELP,
        json!({
            "commands": ["create", "join", "inspect", "diff", "sync-step1", "encode-update", "handle-message", "apply", "watch", "checkpoint", "leave"],
            "encoding": "yjs-v1",
            "syncProtocol": "y-sync-v1",
            "formats": ["document", "markdown", "spreadsheet", "presentation", "pdf"],
            "modes": ["view", "comment", "suggest", "edit"],
        }),
    )
}
