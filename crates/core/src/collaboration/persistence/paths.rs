use std::fs::{self, File};
use std::io::{ErrorKind, Read, Write};
use std::path::{Path, PathBuf};

use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};

use super::{MAX_JSON_BYTES, OPERATION_DIRECTORY};
use crate::collaboration::{collaboration_error, sha256_hex};

pub(super) fn normalized_new_root(path: &Path) -> UseResult<PathBuf> {
    if path.as_os_str().is_empty() || path.file_name().is_none() {
        return Err(collaboration_error(
            "office.collaboration.store_invalid",
            "A collaboration replica path must name a child directory.",
        ));
    }
    let absolute = if path.is_absolute() {
        path.to_path_buf()
    } else {
        std::env::current_dir()
            .map_err(|error| io_error("resolve current directory", Path::new("."), error))?
            .join(path)
    };
    let parent = absolute
        .parent()
        .filter(|parent| parent != &absolute)
        .ok_or_else(|| {
            collaboration_error(
                "office.collaboration.store_invalid",
                "A collaboration replica path must have a parent directory.",
            )
        })?;
    fs::create_dir_all(parent).map_err(|error| io_error("create replica parent", parent, error))?;
    let parent = fs::canonicalize(parent)
        .map_err(|error| io_error("resolve replica parent", parent, error))?;
    let root = parent.join(
        absolute
            .file_name()
            .expect("validated new root has a filename"),
    );
    if root.exists() {
        return Err(collaboration_error(
            "office.collaboration.store_exists",
            format!(
                "Collaboration replica '{}' already exists; refusing to replace it.",
                root.display()
            ),
        )
        .with_detail("store", root.display().to_string()));
    }
    Ok(root)
}

pub(super) fn parse_update_entry_name(path: &Path) -> UseResult<(u64, String)> {
    let stem = path
        .file_stem()
        .and_then(|value| value.to_str())
        .ok_or_else(|| log_corrupt(path, "has a non-UTF-8 filename"))?;
    let (sequence, hash) = stem
        .split_once('-')
        .ok_or_else(|| log_corrupt(path, "has an invalid filename"))?;
    let sequence = sequence
        .parse::<u64>()
        .map_err(|_| log_corrupt(path, "has an invalid sequence number"))?;
    if hash.len() != 64 || !hash.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(log_corrupt(path, "has an invalid update SHA-256"));
    }
    Ok((sequence, hash.to_ascii_lowercase()))
}

pub(super) fn parse_checkpoint_name(path: &Path) -> UseResult<Option<(u64, String, String)>> {
    let Some(extension) = path.extension().and_then(|value| value.to_str()) else {
        return Ok(None);
    };
    if extension != "update" && extension != "state-vector" && extension != "complete" {
        return Ok(None);
    }
    let stem = path
        .file_stem()
        .and_then(|value| value.to_str())
        .ok_or_else(|| log_corrupt(path, "has a non-UTF-8 filename"))?;
    let (sequence, hash) = stem
        .split_once('-')
        .ok_or_else(|| log_corrupt(path, "has an invalid filename"))?;
    let sequence = sequence
        .parse::<u64>()
        .map_err(|_| log_corrupt(path, "has an invalid sequence number"))?;
    if hash.len() != 64 || !hash.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(log_corrupt(path, "has an invalid checkpoint SHA-256"));
    }
    Ok(Some((
        sequence,
        hash.to_ascii_lowercase(),
        extension.to_owned(),
    )))
}

pub(super) fn operation_path(root: &Path, operation_id: &str) -> PathBuf {
    root.join(OPERATION_DIRECTORY)
        .join(format!("{}.json", sha256_hex(operation_id.as_bytes())))
}

pub(super) fn json_bytes(value: &impl Serialize, label: &str) -> UseResult<Vec<u8>> {
    let mut bytes = serde_json::to_vec_pretty(value).map_err(|error| {
        collaboration_error(
            "office.collaboration.store_write_failed",
            format!("Failed to encode the collaboration {label}: {error}"),
        )
    })?;
    bytes.push(b'\n');
    Ok(bytes)
}

pub(super) fn read_json<T: for<'de> Deserialize<'de>>(path: &Path, label: &str) -> UseResult<T> {
    let bytes = read_bounded_file(path, MAX_JSON_BYTES, label)?;
    serde_json::from_slice(&bytes).map_err(|error| {
        collaboration_error(
            "office.collaboration.store_corrupt",
            format!(
                "Collaboration {label} '{}' is invalid JSON: {error}",
                path.display()
            ),
        )
        .with_detail("path", path.display().to_string())
    })
}

pub(super) fn read_bounded_file(path: &Path, max_bytes: u64, label: &str) -> UseResult<Vec<u8>> {
    let mut file = open_bounded_file(path, max_bytes, label)?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes)
        .map_err(|error| io_error(&format!("read {label}"), path, error))?;
    Ok(bytes)
}

pub(super) fn open_bounded_file(path: &Path, max_bytes: u64, label: &str) -> UseResult<File> {
    let metadata = fs::symlink_metadata(path)
        .map_err(|error| io_error(&format!("inspect {label}"), path, error))?;
    if !metadata.is_file() || metadata.file_type().is_symlink() {
        return Err(collaboration_error(
            "office.collaboration.store_corrupt",
            format!(
                "Collaboration {label} '{}' must be a real file.",
                path.display()
            ),
        ));
    }
    if metadata.len() > max_bytes {
        return Err(collaboration_error(
            "office.collaboration.store_corrupt",
            format!(
                "Collaboration {label} '{}' is {} bytes; the limit is {} bytes.",
                path.display(),
                metadata.len(),
                max_bytes
            ),
        ));
    }
    File::open(path).map_err(|error| io_error(&format!("open {label}"), path, error))
}

pub(super) fn read_real_files(directory: &Path) -> UseResult<Vec<PathBuf>> {
    assert_real_directory(directory, "collaboration path")?;
    let mut paths = Vec::new();
    for entry in fs::read_dir(directory)
        .map_err(|error| io_error("read collaboration directory", directory, error))?
    {
        let entry =
            entry.map_err(|error| io_error("read collaboration directory", directory, error))?;
        let file_type = entry
            .file_type()
            .map_err(|error| io_error("inspect collaboration entry", &entry.path(), error))?;
        if file_type.is_symlink() || !file_type.is_file() {
            return Err(collaboration_error(
                "office.collaboration.store_corrupt",
                format!(
                    "Collaboration entry '{}' must be a real file.",
                    entry.path().display()
                ),
            ));
        }
        paths.push(entry.path());
    }
    paths.sort();
    Ok(paths)
}

pub(super) fn assert_real_file(path: &Path, label: &str) -> UseResult<()> {
    let metadata = fs::symlink_metadata(path)
        .map_err(|error| io_error(&format!("inspect {label}"), path, error))?;
    if metadata.is_file() && !metadata.file_type().is_symlink() {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.store_invalid",
        format!(
            "Collaboration {label} '{}' must be a real file.",
            path.display()
        ),
    ))
}

pub(super) fn assert_real_directory(path: &Path, label: &str) -> UseResult<()> {
    let metadata = fs::symlink_metadata(path)
        .map_err(|error| io_error(&format!("inspect {label}"), path, error))?;
    if metadata.is_dir() && !metadata.file_type().is_symlink() {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.store_invalid",
        format!(
            "Collaboration {label} '{}' must be a real directory.",
            path.display()
        ),
    ))
}

pub(super) fn write_new_or_same(path: &Path, bytes: &[u8]) -> UseResult<()> {
    match write_new_bytes(path, bytes) {
        Ok(()) => Ok(()),
        Err(error) if error.code == "office.collaboration.store_entry_exists" => {
            let existing = read_bounded_file(path, bytes.len() as u64, "existing store entry")?;
            if existing == bytes {
                Ok(())
            } else {
                Err(collaboration_error(
                    "office.collaboration.store_corrupt",
                    format!(
                        "Collaboration entry '{}' already exists with different bytes.",
                        path.display()
                    ),
                ))
            }
        }
        Err(error) => Err(error),
    }
}

pub(super) fn write_new_bytes(path: &Path, bytes: &[u8]) -> UseResult<()> {
    let parent = path.parent().ok_or_else(|| {
        collaboration_error(
            "office.collaboration.store_write_failed",
            "A collaboration store entry must have a parent directory.",
        )
    })?;
    let mut temporary = tempfile::NamedTempFile::new_in(parent)
        .map_err(|error| io_error("stage collaboration entry", path, error))?;
    temporary
        .write_all(bytes)
        .map_err(|error| io_error("write collaboration entry", path, error))?;
    temporary
        .as_file()
        .sync_all()
        .map_err(|error| io_error("sync collaboration entry", path, error))?;
    temporary.persist_noclobber(path).map_err(|error| {
        if error.error.kind() == ErrorKind::AlreadyExists {
            collaboration_error(
                "office.collaboration.store_entry_exists",
                format!("Collaboration entry '{}' already exists.", path.display()),
            )
            .with_detail("path", path.display().to_string())
        } else {
            io_error("publish collaboration entry", path, error.error)
        }
    })?;
    Ok(())
}

pub(super) fn remove_scoped_file(root: &Path, directory: &str, path: &Path) -> UseResult<()> {
    let expected_parent = fs::canonicalize(root.join(directory)).map_err(|error| {
        io_error(
            "resolve collaboration directory",
            &root.join(directory),
            error,
        )
    })?;
    let resolved = fs::canonicalize(path)
        .map_err(|error| io_error("resolve collaboration entry", path, error))?;
    if resolved.parent() != Some(expected_parent.as_path()) {
        return Err(collaboration_error(
            "office.collaboration.store_scope_invalid",
            format!(
                "Refusing to remove collaboration entry '{}' outside '{}'.",
                resolved.display(),
                expected_parent.display()
            ),
        ));
    }
    fs::remove_file(&resolved)
        .map_err(|error| io_error("remove compacted collaboration entry", &resolved, error))
}

pub(super) fn sync_directory_best_effort(directory: &Path) -> UseResult<()> {
    match File::open(directory).and_then(|file| file.sync_all()) {
        Ok(()) => Ok(()),
        Err(error)
            if cfg!(windows)
                && matches!(
                    error.kind(),
                    ErrorKind::PermissionDenied | ErrorKind::InvalidInput
                ) =>
        {
            Ok(())
        }
        Err(error) => Err(io_error("sync collaboration directory", directory, error)),
    }
}

pub(super) fn log_corrupt(path: &Path, reason: impl Into<String>) -> UseError {
    collaboration_error(
        "office.collaboration.log_corrupt",
        format!(
            "Collaboration log entry '{}' {}.",
            path.display(),
            reason.into()
        ),
    )
    .with_detail("path", path.display().to_string())
}

pub(super) fn io_error(action: &str, path: &Path, error: std::io::Error) -> UseError {
    collaboration_error(
        "office.collaboration.store_io_failed",
        format!("Failed to {action} '{}': {error}", path.display()),
    )
    .with_detail("path", path.display().to_string())
}
