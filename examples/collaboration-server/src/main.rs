use std::path::PathBuf;

use a3s_boot::Result;
use a3s_office_collaboration_server::{serve, CollaborationConfig};

const DEFAULT_CONFIG: &str = "examples/collaboration-server/collaboration-server.acl";

#[tokio::main]
async fn main() -> Result<()> {
    let config_path = std::env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from(DEFAULT_CONFIG));
    let config = CollaborationConfig::from_acl_file(&config_path)?;
    let bind = config.bind;
    println!(
        "A3S Office collaboration server listening on {bind} using {}",
        config_path.display()
    );
    serve(config).await
}
