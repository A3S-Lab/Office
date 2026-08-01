use std::io::Write as _;
use std::path::Path;

use a3s_use_core::UseResult;
use sha2::{Digest, Sha256};
use tokio::io::AsyncReadExt as _;

use crate::layout::{layout_error, NativeOfficeLayoutRaster};
use crate::{NativeOfficeImage, NativeOfficeImageFormat, PackageRevision};

pub(in crate::layout) async fn ensure_output_available(output: &Path) -> UseResult<()> {
    match tokio::fs::symlink_metadata(output).await {
        Ok(_) => Err(layout_error(
            "use.office.layout_output_exists",
            "Office layout output already exists; refusing to overwrite it.",
        )),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(_) => Err(layout_error(
            "use.office.layout_output_invalid",
            "Office layout output could not be inspected.",
        )),
    }
}

pub(in crate::layout) struct StagedOutput(tempfile::NamedTempFile);

pub(in crate::layout) async fn stage_output(
    output: &Path,
    bytes: Vec<u8>,
) -> UseResult<StagedOutput> {
    let parent = output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."))
        .to_path_buf();
    tokio::task::spawn_blocking(move || {
        let mut temporary = tempfile::Builder::new()
            .prefix(".a3s-office-layout-")
            .tempfile_in(parent)
            .map_err(|_| layout_output_invalid())?;
        temporary
            .write_all(&bytes)
            .and_then(|()| temporary.as_file().sync_all())
            .map_err(|_| layout_output_invalid())?;
        Ok(StagedOutput(temporary))
    })
    .await
    .map_err(|_| layout_output_invalid())?
}

pub(in crate::layout) async fn publish_output(
    staged: StagedOutput,
    output: &Path,
) -> UseResult<()> {
    let output = output.to_path_buf();
    tokio::task::spawn_blocking(move || {
        staged.0.persist_noclobber(output).map_err(|error| {
            if error.error.kind() == std::io::ErrorKind::AlreadyExists {
                layout_error(
                    "use.office.layout_output_exists",
                    "Office layout output already exists; refusing to overwrite it.",
                )
            } else {
                layout_output_invalid()
            }
        })?;
        Ok(())
    })
    .await
    .map_err(|_| layout_output_invalid())?
}

pub(in crate::layout) async fn validate_published_output(
    output: &Path,
    max_output_bytes: u64,
    expected_width: u32,
    expected_height: u32,
    expected_sha256: &str,
    expected_rotation_degrees: u16,
) -> UseResult<NativeOfficeLayoutRaster> {
    let metadata = tokio::fs::symlink_metadata(output)
        .await
        .map_err(|_| layout_output_invalid())?;
    if !metadata.is_file()
        || metadata.file_type().is_symlink()
        || metadata.len() == 0
        || metadata.len() > max_output_bytes
    {
        return Err(layout_output_invalid());
    }
    let bytes = tokio::fs::read(output)
        .await
        .map_err(|_| layout_output_invalid())?;
    let image = NativeOfficeImage::inspect_bytes(&bytes).map_err(|_| layout_output_invalid())?;
    let sha256 = format!("{:x}", Sha256::digest(&bytes));
    if image.format != NativeOfficeImageFormat::Png
        || image.width_px != expected_width
        || image.height_px != expected_height
        || sha256 != expected_sha256
        || u64::try_from(bytes.len()).unwrap_or(u64::MAX) != metadata.len()
    {
        return Err(layout_output_invalid());
    }
    Ok(NativeOfficeLayoutRaster {
        output_path: output.to_path_buf(),
        media_type: "image/png".to_string(),
        width_px: image.width_px,
        height_px: image.height_px,
        byte_length: metadata.len(),
        sha256,
        rotation_degrees: expected_rotation_degrees,
    })
}

pub(in crate::layout) async fn verify_source_revision(
    path: &Path,
    expected: &PackageRevision,
) -> UseResult<()> {
    let metadata = tokio::fs::symlink_metadata(path)
        .await
        .map_err(|_| source_mutated())?;
    if !metadata.is_file()
        || metadata.file_type().is_symlink()
        || metadata.len() != expected.archive_bytes
    {
        return Err(source_mutated());
    }
    let sha256 = hash_regular_file(path, Some(expected.archive_bytes))
        .await
        .map_err(|_| source_mutated())?;
    let final_metadata = tokio::fs::symlink_metadata(path)
        .await
        .map_err(|_| source_mutated())?;
    if !final_metadata.is_file()
        || final_metadata.file_type().is_symlink()
        || final_metadata.len() != expected.archive_bytes
        || sha256 != expected.sha256
    {
        return Err(source_mutated());
    }
    Ok(())
}

pub(in crate::layout) async fn hash_regular_file(
    path: &Path,
    expected_bytes: Option<u64>,
) -> UseResult<String> {
    let mut file = tokio::fs::File::open(path)
        .await
        .map_err(|_| source_mutated())?;
    let mut digest = Sha256::new();
    let mut bytes_read = 0_u64;
    let mut buffer = vec![0_u8; 64 * 1024];
    loop {
        let count = file.read(&mut buffer).await.map_err(|_| source_mutated())?;
        if count == 0 {
            break;
        }
        bytes_read = bytes_read
            .checked_add(u64::try_from(count).map_err(|_| source_mutated())?)
            .ok_or_else(source_mutated)?;
        if expected_bytes.is_some_and(|expected| bytes_read > expected) {
            return Err(source_mutated());
        }
        digest.update(&buffer[..count]);
    }
    if expected_bytes.is_some_and(|expected| bytes_read != expected) {
        return Err(source_mutated());
    }
    Ok(format!("{:x}", digest.finalize()))
}

pub(in crate::layout) fn source_mutated() -> a3s_use_core::UseError {
    layout_error(
        "use.office.layout_source_mutated",
        "The Office layout source no longer matches its immutable byte length and SHA-256.",
    )
}

fn layout_output_invalid() -> a3s_use_core::UseError {
    layout_error(
        "use.office.layout_output_invalid",
        "Office layout output is missing, invalid, or does not match the published receipt.",
    )
}
