#[cfg(windows)]
use std::io::Read as _;
#[cfg(not(windows))]
use std::io::Write as _;
use std::path::Path;
#[cfg(windows)]
use std::path::PathBuf;

use a3s_use_core::UseResult;
use pdfium_render::prelude::{Pdfium, PdfiumLibraryBindings};
use sha2::{Digest, Sha256};

use crate::layout::layout_error;

const MAX_PDFIUM_LIBRARY_BYTES: u64 = 256 * 1024 * 1024;

pub(super) struct PdfiumLibraryInput {
    bytes: Vec<u8>,
    #[cfg(windows)]
    canonical_path: PathBuf,
    #[cfg(windows)]
    source_guard: std::fs::File,
}

impl PdfiumLibraryInput {
    pub(super) fn sha256(&self) -> String {
        format!("{:x}", Sha256::digest(&self.bytes))
    }
}

pub(super) struct BoundPdfiumLibrary {
    pub(super) bindings: Box<dyn PdfiumLibraryBindings>,
    #[cfg(windows)]
    pub(super) source_guard: std::fs::File,
    #[cfg(not(windows))]
    pub(super) staged_library: tempfile::TempDir,
}

#[cfg(not(windows))]
pub(super) async fn read_library(path: &Path) -> UseResult<PdfiumLibraryInput> {
    let metadata = tokio::fs::symlink_metadata(path)
        .await
        .map_err(|_| pdfium_library_invalid())?;
    validate_library_metadata(&metadata)?;
    let bytes = tokio::fs::read(path)
        .await
        .map_err(|_| pdfium_library_invalid())?;
    if u64::try_from(bytes.len()).unwrap_or(u64::MAX) != metadata.len() {
        return Err(pdfium_library_invalid());
    }
    Ok(PdfiumLibraryInput { bytes })
}

#[cfg(windows)]
pub(super) async fn read_library(path: &Path) -> UseResult<PdfiumLibraryInput> {
    let path = path.to_path_buf();
    tokio::task::spawn_blocking(move || read_locked_windows_library(&path))
        .await
        .map_err(|_| pdfium_library_invalid())?
}

#[cfg(windows)]
fn read_locked_windows_library(path: &Path) -> UseResult<PdfiumLibraryInput> {
    use std::os::windows::fs::OpenOptionsExt as _;

    const FILE_SHARE_READ: u32 = 0x0000_0001;

    let requested_metadata =
        std::fs::symlink_metadata(path).map_err(|_| pdfium_library_invalid())?;
    validate_library_metadata(&requested_metadata)?;
    let canonical_path = std::fs::canonicalize(path).map_err(|_| pdfium_library_invalid())?;
    let mut source_guard = std::fs::OpenOptions::new()
        .read(true)
        .share_mode(FILE_SHARE_READ)
        .open(&canonical_path)
        .map_err(|_| pdfium_library_invalid())?;
    let locked_metadata = source_guard
        .metadata()
        .map_err(|_| pdfium_library_invalid())?;
    validate_library_metadata(&locked_metadata)?;
    if locked_metadata.len() != requested_metadata.len() {
        return Err(pdfium_library_invalid());
    }
    let mut bytes = Vec::with_capacity(
        usize::try_from(locked_metadata.len()).map_err(|_| pdfium_library_invalid())?,
    );
    source_guard
        .read_to_end(&mut bytes)
        .map_err(|_| pdfium_library_invalid())?;
    if u64::try_from(bytes.len()).unwrap_or(u64::MAX) != locked_metadata.len()
        || source_guard
            .metadata()
            .map_err(|_| pdfium_library_invalid())?
            .len()
            != locked_metadata.len()
    {
        return Err(pdfium_library_invalid());
    }
    Ok(PdfiumLibraryInput {
        bytes,
        canonical_path,
        source_guard,
    })
}

#[cfg(windows)]
pub(super) fn bind_library(
    input: PdfiumLibraryInput,
    _binary_sha256: &str,
) -> UseResult<BoundPdfiumLibrary> {
    let bindings =
        Pdfium::bind_to_library(&input.canonical_path).map_err(|_| pdfium_library_invalid())?;
    Ok(BoundPdfiumLibrary {
        bindings,
        source_guard: input.source_guard,
    })
}

#[cfg(not(windows))]
pub(super) fn bind_library(
    input: PdfiumLibraryInput,
    binary_sha256: &str,
) -> UseResult<BoundPdfiumLibrary> {
    let staged_library = tempfile::Builder::new()
        .prefix("a3s-office-pdfium-")
        .tempdir()
        .map_err(|_| pdfium_library_invalid())?;
    let staged_path = staged_library
        .path()
        .join(Pdfium::pdfium_platform_library_name());
    let mut file = std::fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&staged_path)
        .map_err(|_| pdfium_library_invalid())?;
    file.write_all(&input.bytes)
        .and_then(|()| file.sync_all())
        .map_err(|_| pdfium_library_invalid())?;
    drop(file);
    let staged_bytes = std::fs::read(&staged_path).map_err(|_| pdfium_library_invalid())?;
    if format!("{:x}", Sha256::digest(&staged_bytes)) != binary_sha256 {
        return Err(pdfium_library_invalid());
    }
    let bindings = Pdfium::bind_to_library(&staged_path).map_err(|_| pdfium_library_invalid())?;
    Ok(BoundPdfiumLibrary {
        bindings,
        staged_library,
    })
}

fn validate_library_metadata(metadata: &std::fs::Metadata) -> UseResult<()> {
    if metadata.is_file()
        && !metadata.file_type().is_symlink()
        && (1..=MAX_PDFIUM_LIBRARY_BYTES).contains(&metadata.len())
    {
        Ok(())
    } else {
        Err(pdfium_library_invalid())
    }
}

pub(super) fn pdfium_library_invalid() -> a3s_use_core::UseError {
    layout_error(
        "use.office.pdfium_library_invalid",
        "The explicit PDFium library is missing, unsafe, oversized, or incompatible with build 7881.",
    )
}

#[cfg(all(test, windows))]
#[path = "library_tests.rs"]
mod windows_tests;
