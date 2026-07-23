use std::process::ExitCode;

#[tokio::main]
async fn main() -> ExitCode {
    let args = std::env::args().skip(1).collect::<Vec<_>>();
    let json = args.iter().any(|argument| argument == "--json");
    match a3s_office_cli::run(&args).await {
        Ok(output) => {
            if output.should_print && json {
                println!(
                    "{}",
                    serde_json::to_string_pretty(&output.json).unwrap_or_default()
                );
            } else if output.should_print && !output.human.is_empty() {
                println!("{}", output.human);
            }
            ExitCode::from(output.exit_code)
        }
        Err(error) => {
            if json {
                let output = serde_json::json!({
                    "schemaVersion": 1,
                    "ok": false,
                    "error": error,
                });
                println!(
                    "{}",
                    serde_json::to_string_pretty(&output).unwrap_or_default()
                );
            } else {
                eprintln!("a3s-office: {error}");
                if let Some(suggestion) = &error.suggestion {
                    eprintln!("suggestion: {suggestion}");
                }
            }
            ExitCode::from(1)
        }
    }
}
