use std::io::Write;
use std::time::Duration;

use a3s_use_core::{UseError, UseResult};

use crate::office_watch::{NativeOfficeWatchOptions, NativeOfficeWatchServer};
use crate::CommandOutput;

use super::usage_error;

const MAX_TIMEOUT_MS: u64 = 24 * 60 * 60 * 1_000;

pub(super) async fn run(args: &[String]) -> UseResult<CommandOutput> {
    let parsed = WatchArguments::parse(args)?;
    let server = NativeOfficeWatchServer::bind(
        &parsed.file,
        NativeOfficeWatchOptions {
            port: parsed.port,
            poll_interval_ms: parsed.poll_ms,
        },
    )
    .await?;
    let ready = server.ready().clone();
    write_ready(&ready, parsed.timeout_ms, parsed.json)?;

    let shutdown = async move {
        if let Some(timeout_ms) = parsed.timeout_ms {
            tokio::select! {
                _ = tokio::time::sleep(Duration::from_millis(timeout_ms)) => {}
                _ = termination_signal() => {}
            }
        } else {
            termination_signal().await;
        }
    };
    server.serve(shutdown).await?;
    Ok(CommandOutput::silent())
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

fn write_ready(
    ready: &crate::office_watch::NativeOfficeWatchReady,
    timeout_ms: Option<u64>,
    json: bool,
) -> UseResult<()> {
    let value = CommandOutput::success(
        format!(
            "Watch: {}\nDocument: {} ({:?})\nRefresh interval: {} ms\nPress Ctrl+C to stop.",
            ready.url, ready.document_name, ready.kind, ready.poll_interval_ms
        ),
        serde_json::json!({
            "operation": "watch",
            "ready": true,
            "server": ready,
            "timeoutMs": timeout_ms
        }),
    );
    let text = if json {
        serde_json::to_string(&value.json).map_err(|error| {
            UseError::new(
                "use.office.watch_receipt_failed",
                format!("Failed to serialize the native Office watch receipt: {error}"),
            )
        })?
    } else {
        value.human
    };
    let mut stdout = std::io::stdout().lock();
    stdout.write_all(text.as_bytes()).map_err(output_error)?;
    stdout.write_all(b"\n").map_err(output_error)?;
    stdout.flush().map_err(output_error)
}

fn output_error(error: std::io::Error) -> UseError {
    UseError::new(
        "use.office.watch_output_failed",
        format!("Failed to write the native Office watch startup receipt: {error}"),
    )
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct WatchArguments {
    file: String,
    port: u16,
    poll_ms: u64,
    timeout_ms: Option<u64>,
    json: bool,
}

impl WatchArguments {
    fn parse(args: &[String]) -> UseResult<Self> {
        let mut positionals = Vec::new();
        let mut port = None;
        let mut poll_ms = None;
        let mut timeout_ms = None;
        let mut json = false;
        let mut index = 1;
        while index < args.len() {
            match args[index].as_str() {
                "--json" => {
                    if json {
                        return Err(usage_error("--json may be specified only once"));
                    }
                    json = true;
                    index += 1;
                }
                "--port" => {
                    set_option(&mut port, args, index, "--port")?;
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
                "--" => {
                    positionals.extend_from_slice(&args[index + 1..]);
                    break;
                }
                option if option.starts_with('-') => {
                    return Err(usage_error(format!(
                        "unknown native Office watch option '{option}'"
                    )));
                }
                value => {
                    positionals.push(value.to_string());
                    index += 1;
                }
            }
        }
        if positionals.len() != 1 {
            return Err(usage_error(
                "office native watch requires exactly one <file>",
            ));
        }
        let timeout_ms = timeout_ms
            .map(|value| parse_u64("--timeout-ms", &value))
            .transpose()?;
        if timeout_ms.is_some_and(|value| value == 0 || value > MAX_TIMEOUT_MS) {
            return Err(usage_error(format!(
                "--timeout-ms must be between 1 and {MAX_TIMEOUT_MS}"
            )));
        }
        let poll_ms = poll_ms
            .map(|value| parse_u64("--poll-ms", &value))
            .transpose()?
            .unwrap_or(crate::office_watch::DEFAULT_NATIVE_OFFICE_WATCH_POLL_MS);
        if !(crate::office_watch::MIN_NATIVE_OFFICE_WATCH_POLL_MS
            ..=crate::office_watch::MAX_NATIVE_OFFICE_WATCH_POLL_MS)
            .contains(&poll_ms)
        {
            return Err(usage_error(format!(
                "--poll-ms must be between {} and {}",
                crate::office_watch::MIN_NATIVE_OFFICE_WATCH_POLL_MS,
                crate::office_watch::MAX_NATIVE_OFFICE_WATCH_POLL_MS
            )));
        }
        Ok(Self {
            file: positionals.remove(0),
            port: port
                .map(|value| parse_u16("--port", &value))
                .transpose()?
                .unwrap_or(0),
            poll_ms,
            timeout_ms,
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

fn parse_u64(name: &str, value: &str) -> UseResult<u64> {
    value.parse().map_err(|_| {
        usage_error(format!(
            "{name} requires a non-negative integer, received '{value}'"
        ))
    })
}

fn parse_u16(name: &str, value: &str) -> UseResult<u16> {
    value.parse().map_err(|_| {
        usage_error(format!(
            "{name} requires an integer from 0 through 65535, received '{value}'"
        ))
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn values(values: &[&str]) -> Vec<String> {
        values.iter().map(|value| (*value).to_string()).collect()
    }

    #[test]
    fn watch_arguments_are_explicit_and_bounded() {
        let parsed = WatchArguments::parse(&values(&[
            "watch",
            "deck.pptx",
            "--port",
            "0",
            "--poll-ms",
            "125",
            "--timeout-ms",
            "5000",
            "--json",
        ]))
        .unwrap();
        assert_eq!(parsed.file, "deck.pptx");
        assert_eq!(parsed.port, 0);
        assert_eq!(parsed.poll_ms, 125);
        assert_eq!(parsed.timeout_ms, Some(5_000));
        assert!(parsed.json);

        for invalid in [
            values(&["watch"]),
            values(&["watch", "one.docx", "two.docx"]),
            values(&["watch", "one.docx", "--port", "65536"]),
            values(&["watch", "one.docx", "--poll-ms", "49"]),
            values(&["watch", "one.docx", "--timeout-ms", "0"]),
            values(&["watch", "one.docx", "--unknown"]),
        ] {
            assert!(WatchArguments::parse(&invalid).is_err(), "{invalid:?}");
        }
    }
}
