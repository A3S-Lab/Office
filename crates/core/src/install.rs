use std::fs::OpenOptions;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};

use a3s_use_core::UseResult;
use fs2::FileExt;
use sha2::{Digest, Sha256};
use tokio::io::AsyncWriteExt;

use crate::discovery::{
    executable, managed_binary_name, managed_root, managed_version_dir, office_error,
    office_status, OfficeInstallReceipt, OfficeInstallSource, OfficeRuntimeStatus, RECEIPT_FILE,
    SUPPORTED_OFFICECLI_VERSION,
};

const MAX_ARTIFACT_BYTES: u64 = 64 * 1024 * 1024;
const LOCK_FILE: &str = ".install.lock";
const STAGE_PREFIX: &str = ".stage-";
const BACKUP_PREFIX: &str = ".backup-";
const ALLOWED_DOWNLOAD_HOSTS: &[&str] = &["github.com", "release-assets.githubusercontent.com"];

#[derive(Debug, Clone, Copy)]
struct PinnedAsset {
    name: &'static str,
    bytes: u64,
    sha256: &'static str,
}

pub async fn install_office_cli(force: bool) -> UseResult<OfficeRuntimeStatus> {
    let current = office_status();
    if !force && current.available {
        return Ok(current);
    }

    let root = managed_root()?;
    let _lock = acquire_lock(&root).await?;
    cleanup_stale_paths(&root).await?;

    let current = office_status();
    if !force && current.available {
        return Ok(current);
    }

    let asset = pinned_asset(std::env::consts::OS, std::env::consts::ARCH)?;
    let source_url = format!(
        "https://github.com/iOfficeAI/OfficeCLI/releases/download/v{SUPPORTED_OFFICECLI_VERSION}/{}",
        asset.name
    );
    let stage = create_stage(&root).await?;
    let staged_executable = stage.join(managed_binary_name());
    let install_result = async {
        let observed = download(&source_url, &staged_executable).await?;
        if observed.bytes != asset.bytes || observed.sha256 != asset.sha256 {
            return Err(office_error(
                "use.office.integrity_mismatch",
                format!(
                    "OfficeCLI artifact integrity mismatch: expected {} bytes and {}, got {} bytes and {}.",
                    asset.bytes, asset.sha256, observed.bytes, observed.sha256
                ),
            ));
        }
        make_executable(&staged_executable).await?;
        write_receipt(
            &stage,
            &OfficeInstallReceipt {
                schema_version: 1,
                provider: "officecli".to_string(),
                version: SUPPORTED_OFFICECLI_VERSION.to_string(),
                source_url,
                artifact_sha256: observed.sha256,
                artifact_bytes: observed.bytes,
            },
        )
        .await?;
        activate(&stage, &managed_version_dir()?).await
    }
    .await;
    if install_result.is_err() {
        let _ = tokio::fs::remove_dir_all(&stage).await;
    }
    install_result?;

    let path = managed_version_dir()?.join(managed_binary_name());
    if executable(&path) {
        Ok(OfficeRuntimeStatus {
            available: true,
            source: OfficeInstallSource::Managed,
            path: Some(path),
            version: Some(SUPPORTED_OFFICECLI_VERSION.to_string()),
            managed_root: Some(root),
            detail: "ready".to_string(),
        })
    } else {
        Err(office_error(
            "use.office.install_failed",
            "OfficeCLI installation completed without a usable managed executable.",
        ))
    }
}

pub async fn repair_office_cli() -> UseResult<OfficeRuntimeStatus> {
    let status = office_status();
    if status.available {
        Ok(status)
    } else {
        install_office_cli(true).await
    }
}

/// Remove only files owned by the A3S-managed OfficeCLI installation.
pub async fn uninstall_managed_office_cli() -> UseResult<bool> {
    let root = managed_root()?;
    let _lock = acquire_lock(&root).await?;
    uninstall_at(&root).await
}

async fn uninstall_at(root: &Path) -> UseResult<bool> {
    let version_dir = root.join(SUPPORTED_OFFICECLI_VERSION);
    let mut changed = if owned_install_directory(&version_dir).await? {
        remove_if_exists(&version_dir).await?
    } else {
        false
    };
    let mut entries = match tokio::fs::read_dir(root).await {
        Ok(entries) => entries,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(changed),
        Err(error) => {
            return Err(office_error(
                "use.office.uninstall_failed",
                format!(
                    "Failed to inspect Office data root '{}': {error}",
                    root.display()
                ),
            ))
        }
    };
    while let Some(entry) = entries.next_entry().await.map_err(|error| {
        office_error(
            "use.office.uninstall_failed",
            format!(
                "Failed to inspect Office data root '{}': {error}",
                root.display()
            ),
        )
    })? {
        let name = entry.file_name();
        let name = name.to_string_lossy();
        if (name.starts_with(STAGE_PREFIX) || name.starts_with(BACKUP_PREFIX))
            && owned_install_directory(&entry.path()).await?
        {
            changed |= remove_entry(&entry.path()).await?;
        }
    }
    Ok(changed)
}

fn pinned_asset(os: &str, arch: &str) -> UseResult<PinnedAsset> {
    let asset = match (os, arch) {
        ("macos", "aarch64") => PinnedAsset {
            name: "officecli-mac-arm64",
            bytes: 33_539_136,
            sha256: "b8582853cc464fa0bdb2fabc2803821472c9449c38b365a7be79fcb53d6356e7",
        },
        ("macos", "x86_64") => PinnedAsset {
            name: "officecli-mac-x64",
            bytes: 34_477_296,
            sha256: "f0073b16a5181837d0b0df3e264a338066b02f4ac16f4758538873fbc32bf9b2",
        },
        ("linux", "aarch64") => PinnedAsset {
            name: "officecli-linux-arm64",
            bytes: 34_508_250,
            sha256: "25bda0d225932159b14ea6ade532c8dbecd2136f1b5c672c003b45bd75afbbb2",
        },
        ("linux", "x86_64") => PinnedAsset {
            name: "officecli-linux-x64",
            bytes: 35_088_440,
            sha256: "2ca3d81be529fd103a7af95a2039b051a08af9d1b5c2c96e85e88731008c402c",
        },
        ("windows", "aarch64") => PinnedAsset {
            name: "officecli-win-arm64.exe",
            bytes: 33_591_172,
            sha256: "94e5ebb1900974681430197035476e2e6c6e905e2e45c2c71ed7de0a3ba87131",
        },
        ("windows", "x86_64") => PinnedAsset {
            name: "officecli-win-x64.exe",
            bytes: 33_144_696,
            sha256: "0ba8550bb236a2a23982311a747b22f318d7bd18c1c06a402d96f7642c85fb6a",
        },
        _ => {
            return Err(office_error(
                "use.office.platform_unsupported",
                format!("OfficeCLI {SUPPORTED_OFFICECLI_VERSION} does not support {os}/{arch}."),
            ))
        }
    };
    Ok(asset)
}

struct DownloadedArtifact {
    bytes: u64,
    sha256: String,
}

async fn download(url: &str, destination: &Path) -> UseResult<DownloadedArtifact> {
    let parsed = reqwest::Url::parse(url).map_err(|error| {
        office_error(
            "use.office.download_source_invalid",
            format!("Invalid OfficeCLI download URL: {error}"),
        )
    })?;
    if parsed.scheme() != "https"
        || !parsed
            .host_str()
            .is_some_and(|host| ALLOWED_DOWNLOAD_HOSTS.contains(&host))
    {
        return Err(office_error(
            "use.office.download_source_invalid",
            "OfficeCLI download source is not an approved HTTPS host.",
        ));
    }
    let redirect = reqwest::redirect::Policy::custom(|attempt| {
        let approved = attempt.previous().len() < 10
            && attempt.url().scheme() == "https"
            && attempt
                .url()
                .host_str()
                .is_some_and(|host| ALLOWED_DOWNLOAD_HOSTS.contains(&host));
        if approved {
            attempt.follow()
        } else {
            attempt.error("OfficeCLI download redirected to an unapproved host")
        }
    });
    let client = reqwest::Client::builder()
        .user_agent("a3s-use/0.1")
        .redirect(redirect)
        .timeout(std::time::Duration::from_secs(300))
        .build()
        .map_err(|error| {
            office_error(
                "use.office.download_failed",
                format!("Failed to build OfficeCLI download client: {error}"),
            )
        })?;
    let mut response = client
        .get(parsed)
        .send()
        .await
        .map_err(|error| {
            office_error(
                "use.office.download_failed",
                format!("Failed to download OfficeCLI: {error}"),
            )
        })?
        .error_for_status()
        .map_err(|error| {
            office_error(
                "use.office.download_failed",
                format!("OfficeCLI download failed: {error}"),
            )
        })?;
    if response
        .content_length()
        .is_some_and(|length| length > MAX_ARTIFACT_BYTES)
    {
        return Err(office_error(
            "use.office.download_too_large",
            format!("OfficeCLI download exceeds the {MAX_ARTIFACT_BYTES}-byte limit."),
        ));
    }
    let mut file = tokio::fs::OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(destination)
        .await
        .map_err(|error| {
            office_error(
                "use.office.install_failed",
                format!("Failed to create '{}': {error}", destination.display()),
            )
        })?;
    let mut hasher = Sha256::new();
    let mut bytes = 0_u64;
    while let Some(chunk) = response.chunk().await.map_err(|error| {
        office_error(
            "use.office.download_failed",
            format!("Failed to read OfficeCLI download: {error}"),
        )
    })? {
        bytes = bytes.checked_add(chunk.len() as u64).ok_or_else(|| {
            office_error(
                "use.office.download_too_large",
                "OfficeCLI download size overflowed.",
            )
        })?;
        if bytes > MAX_ARTIFACT_BYTES {
            return Err(office_error(
                "use.office.download_too_large",
                format!("OfficeCLI download exceeds the {MAX_ARTIFACT_BYTES}-byte limit."),
            ));
        }
        hasher.update(&chunk);
        file.write_all(&chunk).await.map_err(|error| {
            office_error(
                "use.office.install_failed",
                format!("Failed to write '{}': {error}", destination.display()),
            )
        })?;
    }
    file.sync_all().await.map_err(|error| {
        office_error(
            "use.office.install_failed",
            format!("Failed to sync '{}': {error}", destination.display()),
        )
    })?;
    Ok(DownloadedArtifact {
        bytes,
        sha256: format!("{:x}", hasher.finalize()),
    })
}

async fn acquire_lock(root: &Path) -> UseResult<InstallLock> {
    tokio::fs::create_dir_all(root).await.map_err(|error| {
        office_error(
            "use.office.install_failed",
            format!(
                "Failed to create Office data root '{}': {error}",
                root.display()
            ),
        )
    })?;
    let path = root.join(LOCK_FILE);
    tokio::task::spawn_blocking(move || {
        let file = OpenOptions::new()
            .create(true)
            .read(true)
            .write(true)
            .truncate(false)
            .open(&path)
            .map_err(|error| {
                office_error(
                    "use.office.install_failed",
                    format!(
                        "Failed to open Office install lock '{}': {error}",
                        path.display()
                    ),
                )
            })?;
        file.lock_exclusive().map_err(|error| {
            office_error(
                "use.office.install_failed",
                format!(
                    "Failed to acquire Office install lock '{}': {error}",
                    path.display()
                ),
            )
        })?;
        Ok(InstallLock { _file: file })
    })
    .await
    .map_err(|error| {
        office_error(
            "use.office.install_failed",
            format!("Office install lock task failed: {error}"),
        )
    })?
}

struct InstallLock {
    _file: std::fs::File,
}

async fn create_stage(root: &Path) -> UseResult<PathBuf> {
    static NEXT_ID: AtomicU64 = AtomicU64::new(1);
    for _ in 0..32 {
        let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
        let path = root.join(format!("{STAGE_PREFIX}{}-{id}", std::process::id()));
        match tokio::fs::create_dir(&path).await {
            Ok(()) => return Ok(path),
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(office_error(
                    "use.office.install_failed",
                    format!(
                        "Failed to create Office staging directory '{}': {error}",
                        path.display()
                    ),
                ))
            }
        }
    }
    Err(office_error(
        "use.office.install_failed",
        "Failed to allocate a unique Office staging directory.",
    ))
}

async fn cleanup_stale_paths(root: &Path) -> UseResult<()> {
    let mut entries = tokio::fs::read_dir(root).await.map_err(|error| {
        office_error(
            "use.office.install_failed",
            format!(
                "Failed to inspect Office data root '{}': {error}",
                root.display()
            ),
        )
    })?;
    while let Some(entry) = entries.next_entry().await.map_err(|error| {
        office_error(
            "use.office.install_failed",
            format!(
                "Failed to inspect Office data root '{}': {error}",
                root.display()
            ),
        )
    })? {
        let name = entry.file_name();
        let name = name.to_string_lossy();
        if name.starts_with(STAGE_PREFIX) && owned_install_directory(&entry.path()).await? {
            remove_entry(&entry.path()).await?;
        } else if name.starts_with(BACKUP_PREFIX) && owned_install_directory(&entry.path()).await? {
            let target = root.join(SUPPORTED_OFFICECLI_VERSION);
            let target_exists = tokio::fs::try_exists(&target).await.map_err(|error| {
                office_error(
                    "use.office.install_failed",
                    format!("Failed to inspect '{}': {error}", target.display()),
                )
            })?;
            if target_exists && owned_install_directory(&target).await? {
                remove_entry(&entry.path()).await?;
            } else if !target_exists {
                tokio::fs::rename(entry.path(), &target)
                    .await
                    .map_err(|error| {
                        office_error(
                            "use.office.install_failed",
                            format!(
                                "Failed to recover Office installation '{}': {error}",
                                target.display()
                            ),
                        )
                    })?;
            }
        }
    }
    Ok(())
}

async fn activate(stage: &Path, target: &Path) -> UseResult<()> {
    if !owned_install_directory(stage).await? {
        return Err(office_error(
            "use.office.install_stage_unowned",
            format!(
                "OfficeCLI staging directory '{}' has no valid A3S install receipt.",
                stage.display()
            ),
        ));
    }
    let had_target = tokio::fs::try_exists(target).await.map_err(|error| {
        office_error(
            "use.office.install_failed",
            format!("Failed to inspect '{}': {error}", target.display()),
        )
    })?;
    if had_target && !owned_install_directory(target).await? {
        return Err(office_error(
            "use.office.install_path_unowned",
            format!(
                "Refusing to replace unowned Office directory '{}'.",
                target.display()
            ),
        )
        .with_suggestion("Move the directory aside or restore its A3S install receipt."));
    }
    let backup = unique_backup_path(target).await?;
    if had_target {
        tokio::fs::rename(target, &backup).await.map_err(|error| {
            office_error(
                "use.office.install_failed",
                format!("Failed to stage existing Office installation: {error}"),
            )
        })?;
    }
    if let Err(error) = tokio::fs::rename(stage, target).await {
        if had_target {
            let _ = tokio::fs::rename(&backup, target).await;
        }
        return Err(office_error(
            "use.office.install_failed",
            format!(
                "Failed to activate OfficeCLI '{}': {error}",
                target.display()
            ),
        ));
    }
    if had_target {
        remove_entry(&backup).await?;
    }
    Ok(())
}

async fn unique_backup_path(target: &Path) -> UseResult<PathBuf> {
    static NEXT_ID: AtomicU64 = AtomicU64::new(1);
    for _ in 0..32 {
        let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
        let path = target.with_file_name(format!(
            "{BACKUP_PREFIX}{}-{}-{id}",
            SUPPORTED_OFFICECLI_VERSION,
            std::process::id()
        ));
        match tokio::fs::symlink_metadata(&path).await {
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(path),
            Ok(_) => continue,
            Err(error) => {
                return Err(office_error(
                    "use.office.install_failed",
                    format!(
                        "Failed to inspect Office backup path '{}': {error}",
                        path.display()
                    ),
                ))
            }
        }
    }
    Err(office_error(
        "use.office.install_failed",
        "Failed to allocate a unique Office backup path.",
    ))
}

async fn owned_install_directory(directory: &Path) -> UseResult<bool> {
    let path = directory.join(RECEIPT_FILE);
    let bytes = match tokio::fs::read(&path).await {
        Ok(bytes) => bytes,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(false),
        Err(error) => {
            return Err(office_error(
                "use.office.ownership_check_failed",
                format!(
                    "Failed to read Office install receipt '{}': {error}",
                    path.display()
                ),
            ))
        }
    };
    let Ok(receipt) = serde_json::from_slice::<OfficeInstallReceipt>(&bytes) else {
        return Ok(false);
    };
    Ok(receipt.schema_version == 1
        && receipt.provider == "officecli"
        && receipt.version == SUPPORTED_OFFICECLI_VERSION)
}

async fn write_receipt(directory: &Path, receipt: &OfficeInstallReceipt) -> UseResult<()> {
    let bytes = serde_json::to_vec_pretty(receipt).map_err(|error| {
        office_error(
            "use.office.install_failed",
            format!("Failed to encode Office install receipt: {error}"),
        )
    })?;
    let path = directory.join(RECEIPT_FILE);
    let mut file = tokio::fs::OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(&path)
        .await
        .map_err(|error| {
            office_error(
                "use.office.install_failed",
                format!(
                    "Failed to create Office install receipt '{}': {error}",
                    path.display()
                ),
            )
        })?;
    file.write_all(&bytes).await.map_err(|error| {
        office_error(
            "use.office.install_failed",
            format!(
                "Failed to write Office install receipt '{}': {error}",
                path.display()
            ),
        )
    })?;
    file.sync_all().await.map_err(|error| {
        office_error(
            "use.office.install_failed",
            format!(
                "Failed to sync Office install receipt '{}': {error}",
                path.display()
            ),
        )
    })
}

#[cfg(unix)]
async fn make_executable(path: &Path) -> UseResult<()> {
    use std::os::unix::fs::PermissionsExt;
    let mut permissions = tokio::fs::metadata(path)
        .await
        .map_err(|error| {
            office_error(
                "use.office.install_failed",
                format!("Failed to inspect OfficeCLI '{}': {error}", path.display()),
            )
        })?
        .permissions();
    permissions.set_mode(0o755);
    tokio::fs::set_permissions(path, permissions)
        .await
        .map_err(|error| {
            office_error(
                "use.office.install_failed",
                format!(
                    "Failed to set OfficeCLI permissions '{}': {error}",
                    path.display()
                ),
            )
        })?;
    Ok(())
}

#[cfg(not(unix))]
async fn make_executable(_path: &Path) -> UseResult<()> {
    Ok(())
}

async fn remove_if_exists(path: &Path) -> UseResult<bool> {
    match tokio::fs::symlink_metadata(path).await {
        Ok(_) => remove_entry(path).await,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(false),
        Err(error) => Err(office_error(
            "use.office.uninstall_failed",
            format!(
                "Failed to inspect managed Office path '{}': {error}",
                path.display()
            ),
        )),
    }
}

async fn remove_entry(path: &Path) -> UseResult<bool> {
    let metadata = match tokio::fs::symlink_metadata(path).await {
        Ok(metadata) => metadata,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(false),
        Err(error) => {
            return Err(office_error(
                "use.office.uninstall_failed",
                format!(
                    "Failed to inspect managed Office path '{}': {error}",
                    path.display()
                ),
            ))
        }
    };
    let result = if metadata.is_dir() && !metadata.file_type().is_symlink() {
        tokio::fs::remove_dir_all(path).await
    } else {
        tokio::fs::remove_file(path).await
    };
    result.map_err(|error| {
        office_error(
            "use.office.uninstall_failed",
            format!(
                "Failed to remove managed Office path '{}': {error}",
                path.display()
            ),
        )
    })?;
    Ok(true)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn test_receipt() -> OfficeInstallReceipt {
        OfficeInstallReceipt {
            schema_version: 1,
            provider: "officecli".to_string(),
            version: SUPPORTED_OFFICECLI_VERSION.to_string(),
            source_url: "https://github.com/iOfficeAI/OfficeCLI/releases/test".to_string(),
            artifact_sha256: "0".repeat(64),
            artifact_bytes: 1,
        }
    }

    #[test]
    fn pinned_assets_have_publisher_sha256_for_every_supported_target() {
        for (os, arch) in [
            ("macos", "aarch64"),
            ("macos", "x86_64"),
            ("linux", "aarch64"),
            ("linux", "x86_64"),
            ("windows", "aarch64"),
            ("windows", "x86_64"),
        ] {
            let asset = pinned_asset(os, arch).unwrap();
            assert_eq!(asset.sha256.len(), 64);
            assert!(asset.bytes < MAX_ARTIFACT_BYTES);
        }
        assert_eq!(
            pinned_asset("plan9", "x86_64").unwrap_err().code,
            "use.office.platform_unsupported"
        );
    }

    #[tokio::test]
    async fn uninstall_is_idempotent_and_preserves_unowned_files() {
        let temp = tempfile::tempdir().unwrap();
        let version = temp.path().join(SUPPORTED_OFFICECLI_VERSION);
        tokio::fs::create_dir_all(&version).await.unwrap();
        tokio::fs::write(version.join("officecli"), b"managed")
            .await
            .unwrap();
        write_receipt(&version, &test_receipt()).await.unwrap();
        tokio::fs::write(temp.path().join("keep"), b"unowned")
            .await
            .unwrap();

        assert!(uninstall_at(temp.path()).await.unwrap());
        assert!(!uninstall_at(temp.path()).await.unwrap());
        assert!(tokio::fs::try_exists(temp.path().join("keep"))
            .await
            .unwrap());
    }

    #[tokio::test]
    async fn uninstall_preserves_an_unowned_version_directory() {
        let temp = tempfile::tempdir().unwrap();
        let version = temp.path().join(SUPPORTED_OFFICECLI_VERSION);
        tokio::fs::create_dir_all(&version).await.unwrap();
        tokio::fs::write(version.join("keep"), b"unowned")
            .await
            .unwrap();

        assert!(!uninstall_at(temp.path()).await.unwrap());
        assert!(tokio::fs::try_exists(version.join("keep")).await.unwrap());
    }

    #[tokio::test]
    async fn stale_cleanup_preserves_an_unowned_backup() {
        let temp = tempfile::tempdir().unwrap();
        let backup = temp.path().join(format!("{BACKUP_PREFIX}fixture"));
        tokio::fs::create_dir_all(&backup).await.unwrap();
        tokio::fs::write(backup.join("keep"), b"unowned")
            .await
            .unwrap();

        cleanup_stale_paths(temp.path()).await.unwrap();
        assert!(tokio::fs::try_exists(backup.join("keep")).await.unwrap());
        assert!(
            !tokio::fs::try_exists(temp.path().join(SUPPORTED_OFFICECLI_VERSION))
                .await
                .unwrap()
        );
    }

    #[tokio::test]
    async fn stale_cleanup_recovers_an_owned_backup() {
        let temp = tempfile::tempdir().unwrap();
        let backup = temp.path().join(format!("{BACKUP_PREFIX}fixture"));
        tokio::fs::create_dir_all(&backup).await.unwrap();
        tokio::fs::write(backup.join("officecli"), b"managed")
            .await
            .unwrap();
        write_receipt(&backup, &test_receipt()).await.unwrap();

        cleanup_stale_paths(temp.path()).await.unwrap();
        let target = temp.path().join(SUPPORTED_OFFICECLI_VERSION);
        assert!(tokio::fs::try_exists(target.join(RECEIPT_FILE))
            .await
            .unwrap());
        assert!(!tokio::fs::try_exists(&backup).await.unwrap());
    }

    #[tokio::test]
    async fn activation_refuses_to_replace_an_unowned_target() {
        let temp = tempfile::tempdir().unwrap();
        let stage = temp.path().join(format!("{STAGE_PREFIX}fixture"));
        let target = temp.path().join(SUPPORTED_OFFICECLI_VERSION);
        tokio::fs::create_dir_all(&stage).await.unwrap();
        write_receipt(&stage, &test_receipt()).await.unwrap();
        tokio::fs::create_dir_all(&target).await.unwrap();
        tokio::fs::write(target.join("keep"), b"unowned")
            .await
            .unwrap();

        let error = activate(&stage, &target).await.unwrap_err();
        assert_eq!(error.code, "use.office.install_path_unowned");
        assert!(tokio::fs::try_exists(stage.join(RECEIPT_FILE))
            .await
            .unwrap());
        assert!(tokio::fs::try_exists(target.join("keep")).await.unwrap());
    }

    #[test]
    fn managed_binary_must_be_executable() {
        let temp = tempfile::tempdir().unwrap();
        let path = temp.path().join("officecli");
        std::fs::write(&path, b"fixture").unwrap();
        #[cfg(unix)]
        assert!(!executable(&path));
    }
}
