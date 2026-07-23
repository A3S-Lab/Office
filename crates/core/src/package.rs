use std::collections::{BTreeMap, BTreeSet};
use std::io::{Read, Seek, Write};
use std::path::{Path, PathBuf};
use std::sync::Arc;

use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use tempfile::NamedTempFile;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::discovery::office_error;
use crate::xml::{LosslessXmlPart, XmlLimits};
use crate::DocumentKind;

const CONTENT_TYPES_PART: &str = "[Content_Types].xml";
const ROOT_RELATIONSHIPS_PART: &str = "_rels/.rels";
const COMPRESSION_RATIO_MINIMUM_BYTES: u64 = 1024 * 1024;

/// Resource limits applied before an OOXML package is admitted to the native engine.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PackageLimits {
    pub max_archive_bytes: u64,
    pub max_entries: usize,
    pub max_part_bytes: u64,
    pub max_uncompressed_bytes: u64,
    pub max_compression_ratio: u64,
}

/// Content revision used to prevent one writer from overwriting another.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PackageRevision {
    pub archive_bytes: u64,
    pub sha256: String,
}

#[derive(Debug, Clone)]
enum WriteTarget {
    Replace,
    Revision(PackageRevision),
    CreateNew,
}

impl Default for PackageLimits {
    fn default() -> Self {
        Self {
            max_archive_bytes: 512 * 1024 * 1024,
            max_entries: 16_384,
            max_part_bytes: 128 * 1024 * 1024,
            max_uncompressed_bytes: 1024 * 1024 * 1024,
            max_compression_ratio: 250,
        }
    }
}

/// An A3S-owned, loss-preserving OOXML package.
///
/// The package kernel understands OPC container boundaries, document kind,
/// resource limits, and atomic persistence. Format-specific semantic engines
/// build on top of its part API without discarding unknown package parts.
#[derive(Debug, Clone)]
pub struct NativeOfficePackage {
    path: PathBuf,
    kind: DocumentKind,
    limits: PackageLimits,
    parts: BTreeMap<String, Arc<[u8]>>,
    source_revision: PackageRevision,
    dirty: bool,
}

impl NativeOfficePackage {
    pub async fn open(path: impl AsRef<Path>) -> UseResult<Self> {
        Self::open_with_limits(path, PackageLimits::default()).await
    }

    pub async fn open_with_limits(
        path: impl AsRef<Path>,
        limits: PackageLimits,
    ) -> UseResult<Self> {
        validate_limits(limits)?;
        let path = absolute(path.as_ref())?;
        let error_path = path.clone();
        tokio::task::spawn_blocking(move || read_package(path, limits))
            .await
            .map_err(|error| {
                package_error(
                    "use.office.package_open_failed",
                    format!(
                        "Native Office package task failed for '{}': {error}",
                        error_path.display()
                    ),
                )
            })?
    }

    /// Creates a blank OOXML package selected by the destination extension.
    ///
    /// Creation is atomic and fails if the destination already exists.
    pub async fn create(path: impl AsRef<Path>) -> UseResult<Self> {
        Self::create_with_limits(path, PackageLimits::default()).await
    }

    pub async fn create_with_limits(
        path: impl AsRef<Path>,
        limits: PackageLimits,
    ) -> UseResult<Self> {
        validate_limits(limits)?;
        let path = absolute(path.as_ref())?;
        let kind = kind_from_extension(&path)?;
        let parts = crate::template::blank_parts(kind);
        validate_structure(kind, &parts)?;
        validate_part_limits(&parts, limits)?;
        let task_path = path.clone();
        let task_parts = parts.clone();
        let max_archive_bytes = limits.max_archive_bytes;
        let error_path = path.clone();
        let source_revision = tokio::task::spawn_blocking(move || {
            write_package_atomically(
                &task_path,
                &task_parts,
                WriteTarget::CreateNew,
                max_archive_bytes,
            )
        })
        .await
        .map_err(|error| {
            package_error(
                "use.office.package_create_failed",
                format!(
                    "Native Office create task failed for '{}': {error}",
                    error_path.display()
                ),
            )
        })??;
        Ok(Self {
            path,
            kind,
            limits,
            parts,
            source_revision,
            dirty: false,
        })
    }

    pub(crate) fn blank_in_memory(kind: DocumentKind, limits: PackageLimits) -> UseResult<Self> {
        validate_limits(limits)?;
        let parts = crate::template::blank_parts(kind);
        validate_structure(kind, &parts)?;
        validate_part_limits(&parts, limits)?;
        Ok(Self {
            path: PathBuf::from(format!("a3s-native-replay.{}", kind.extension())),
            kind,
            limits,
            source_revision: PackageRevision {
                archive_bytes: 0,
                sha256: content_sha256(&parts),
            },
            parts,
            dirty: false,
        })
    }

    pub fn path(&self) -> &Path {
        &self.path
    }

    pub fn kind(&self) -> DocumentKind {
        self.kind
    }

    pub fn limits(&self) -> PackageLimits {
        self.limits
    }

    pub fn source_revision(&self) -> &PackageRevision {
        &self.source_revision
    }

    /// SHA-256 of the ordered, uncompressed OOXML part map.
    ///
    /// This fingerprint is independent of ZIP compression and entry metadata.
    /// It changes when a part name or any part byte changes.
    pub fn content_sha256(&self) -> String {
        content_sha256(&self.parts)
    }

    pub fn part_names(&self) -> impl Iterator<Item = &str> {
        self.parts.keys().map(String::as_str)
    }

    pub fn contains_part(&self, name: &str) -> bool {
        query_part_name(name)
            .ok()
            .is_some_and(|name| self.parts.contains_key(&name))
    }

    pub fn part(&self, name: &str) -> UseResult<&[u8]> {
        let name = query_part_name(name)?;
        self.parts.get(&name).map(AsRef::as_ref).ok_or_else(|| {
            package_error(
                "use.office.package_part_missing",
                format!("OOXML package part '{name}' does not exist."),
            )
        })
    }

    pub fn xml_part(&self, name: &str) -> UseResult<LosslessXmlPart> {
        self.xml_part_with_limits(name, XmlLimits::default())
    }

    pub fn xml_part_with_limits(
        &self,
        name: &str,
        limits: XmlLimits,
    ) -> UseResult<LosslessXmlPart> {
        let name = query_part_name(name)?;
        let bytes = self.parts.get(&name).cloned().ok_or_else(|| {
            package_error(
                "use.office.package_part_missing",
                format!("OOXML package part '{name}' does not exist."),
            )
        })?;
        LosslessXmlPart::parse_with_limits(name, bytes, limits)
    }

    pub fn set_part(&mut self, name: &str, bytes: Vec<u8>) -> UseResult<()> {
        let name = query_part_name(name)?;
        if self
            .parts
            .keys()
            .any(|existing| existing != &name && existing.eq_ignore_ascii_case(&name))
        {
            return Err(duplicate_part(&name));
        }
        if !self.parts.contains_key(&name) && self.parts.len() >= self.limits.max_entries {
            return Err(package_error(
                "use.office.package_entry_limit",
                format!(
                    "OOXML package cannot add part '{name}'; the {}-entry limit is reached.",
                    self.limits.max_entries
                ),
            ));
        }
        let bytes_len = u64::try_from(bytes.len()).map_err(|_| {
            package_error(
                "use.office.package_part_too_large",
                format!("OOXML package part '{name}' is too large for this platform."),
            )
        })?;
        if bytes_len > self.limits.max_part_bytes {
            return Err(package_error(
                "use.office.package_part_too_large",
                format!(
                    "OOXML package part '{name}' exceeds the {}-byte limit.",
                    self.limits.max_part_bytes
                ),
            ));
        }
        let previous_len = self
            .parts
            .get(&name)
            .map_or(0, |part| u64::try_from(part.len()).unwrap_or(u64::MAX));
        let total = package_bytes(&self.parts)?
            .saturating_sub(previous_len)
            .checked_add(bytes_len)
            .ok_or_else(|| package_size_error(self.limits.max_uncompressed_bytes))?;
        if total > self.limits.max_uncompressed_bytes {
            return Err(package_size_error(self.limits.max_uncompressed_bytes));
        }
        self.parts.insert(name, Arc::from(bytes));
        self.dirty = true;
        Ok(())
    }

    pub fn remove_part(&mut self, name: &str) -> UseResult<bool> {
        let name = query_part_name(name)?;
        let changed = self.parts.remove(&name).is_some();
        self.dirty |= changed;
        Ok(changed)
    }

    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    pub async fn save(&mut self) -> UseResult<()> {
        if !self.dirty {
            return Ok(());
        }
        let path = self.path.clone();
        self.write_to(path).await
    }

    pub async fn save_as(&mut self, path: impl AsRef<Path>) -> UseResult<()> {
        let path = absolute(path.as_ref())?;
        require_extension_kind(&path, self.kind)?;
        self.write_to(path).await
    }

    /// Atomically saves to a new path and refuses to replace any existing
    /// filesystem entry.
    pub async fn save_as_new(&mut self, path: impl AsRef<Path>) -> UseResult<()> {
        let path = absolute(path.as_ref())?;
        require_extension_kind(&path, self.kind)?;
        self.write_to_target(path, WriteTarget::CreateNew).await
    }

    async fn write_to(&mut self, path: PathBuf) -> UseResult<()> {
        let target = if path == self.path {
            WriteTarget::Revision(self.source_revision.clone())
        } else {
            WriteTarget::Replace
        };
        self.write_to_target(path, target).await
    }

    async fn write_to_target(&mut self, path: PathBuf, target: WriteTarget) -> UseResult<()> {
        validate_structure(self.kind, &self.parts)?;
        validate_part_limits(&self.parts, self.limits)?;
        let parts = self.parts.clone();
        let max_archive_bytes = self.limits.max_archive_bytes;
        let error_path = path.clone();
        let revision = tokio::task::spawn_blocking(move || {
            write_package_atomically(&path, &parts, target, max_archive_bytes)
        })
        .await
        .map_err(|error| {
            package_error(
                "use.office.package_save_failed",
                format!(
                    "Native Office save task failed for '{}': {error}",
                    error_path.display()
                ),
            )
        })??;
        self.path = error_path;
        self.source_revision = revision;
        self.dirty = false;
        Ok(())
    }
}

fn content_sha256(parts: &BTreeMap<String, Arc<[u8]>>) -> String {
    let mut digest = Sha256::new();
    digest.update(b"a3s-office-package-v1\0");
    digest.update((parts.len() as u64).to_be_bytes());
    for (name, bytes) in parts {
        digest.update((name.len() as u64).to_be_bytes());
        digest.update(name.as_bytes());
        digest.update((bytes.len() as u64).to_be_bytes());
        digest.update(bytes);
    }
    format!("{:x}", digest.finalize())
}

fn read_package(path: PathBuf, limits: PackageLimits) -> UseResult<NativeOfficePackage> {
    let metadata = std::fs::symlink_metadata(&path).map_err(|error| {
        package_error(
            "use.office.package_open_failed",
            format!(
                "Failed to inspect Office document '{}': {error}",
                path.display()
            ),
        )
    })?;
    if !metadata.is_file() || metadata.file_type().is_symlink() {
        return Err(package_error(
            "use.office.package_open_failed",
            format!(
                "Office document '{}' is not a regular, non-symlink file.",
                path.display()
            ),
        ));
    }
    if metadata.len() > limits.max_archive_bytes {
        return Err(package_error(
            "use.office.package_too_large",
            format!(
                "Office document '{}' exceeds the {}-byte archive limit.",
                path.display(),
                limits.max_archive_bytes
            ),
        ));
    }
    let mut file = std::fs::File::open(&path).map_err(|error| {
        package_error(
            "use.office.package_open_failed",
            format!(
                "Failed to open Office document '{}': {error}",
                path.display()
            ),
        )
    })?;
    let source_revision = revision_from_reader(&mut file, metadata.len(), &path)?;
    file.rewind().map_err(|error| {
        package_error(
            "use.office.package_open_failed",
            format!(
                "Failed to rewind Office document '{}': {error}",
                path.display()
            ),
        )
    })?;
    let mut archive = ZipArchive::new(file).map_err(|error| {
        package_error(
            "use.office.package_invalid",
            format!(
                "Office document '{}' is not a valid ZIP package: {error}",
                path.display()
            ),
        )
    })?;
    if archive.len() > limits.max_entries {
        return Err(package_error(
            "use.office.package_too_many_parts",
            format!(
                "Office document '{}' contains {} entries; the limit is {}.",
                path.display(),
                archive.len(),
                limits.max_entries
            ),
        ));
    }

    let mut parts = BTreeMap::new();
    let mut folded_names = BTreeSet::new();
    let mut total = 0_u64;
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).map_err(|error| {
            package_error(
                "use.office.package_invalid",
                format!("Failed to read OOXML ZIP entry {index}: {error}"),
            )
        })?;
        if entry.is_dir() {
            let directory = entry.name().strip_suffix('/').unwrap_or(entry.name());
            archive_part_name(directory)?;
            continue;
        }
        let name = archive_part_name(entry.name())?;
        if entry.encrypted() {
            return Err(package_error(
                "use.office.package_encrypted",
                format!("Encrypted OOXML package part '{name}' is not supported."),
            ));
        }
        if entry
            .unix_mode()
            .is_some_and(|mode| mode & 0o170000 == 0o120000)
        {
            return Err(package_error(
                "use.office.package_part_invalid",
                format!("OOXML package part '{name}' cannot be a symbolic link."),
            ));
        }
        let folded = name.to_ascii_lowercase();
        if !folded_names.insert(folded) {
            return Err(duplicate_part(&name));
        }
        if entry.size() > limits.max_part_bytes {
            return Err(package_error(
                "use.office.package_part_too_large",
                format!(
                    "OOXML package part '{name}' exceeds the {}-byte limit.",
                    limits.max_part_bytes
                ),
            ));
        }
        enforce_compression_ratio(&name, entry.size(), entry.compressed_size(), limits)?;
        total = total
            .checked_add(entry.size())
            .ok_or_else(|| package_size_error(limits.max_uncompressed_bytes))?;
        if total > limits.max_uncompressed_bytes {
            return Err(package_size_error(limits.max_uncompressed_bytes));
        }
        let capacity = usize::try_from(entry.size()).map_err(|_| {
            package_error(
                "use.office.package_part_too_large",
                format!("OOXML package part '{name}' is too large for this platform."),
            )
        })?;
        let mut bytes = Vec::with_capacity(capacity);
        entry.read_to_end(&mut bytes).map_err(|error| {
            package_error(
                "use.office.package_invalid",
                format!("Failed to decompress OOXML package part '{name}': {error}"),
            )
        })?;
        if u64::try_from(bytes.len()).unwrap_or(u64::MAX) != entry.size() {
            return Err(package_error(
                "use.office.package_invalid",
                format!("OOXML package part '{name}' has an inconsistent expanded size."),
            ));
        }
        parts.insert(name, Arc::from(bytes));
    }

    let kind = detect_kind(&parts)?;
    require_extension_kind(&path, kind)?;
    validate_structure(kind, &parts)?;
    Ok(NativeOfficePackage {
        path,
        kind,
        limits,
        parts,
        source_revision,
        dirty: false,
    })
}

fn write_package_atomically(
    path: &Path,
    parts: &BTreeMap<String, Arc<[u8]>>,
    target: WriteTarget,
    max_archive_bytes: u64,
) -> UseResult<PackageRevision> {
    let parent = path.parent().ok_or_else(|| {
        package_error(
            "use.office.package_save_failed",
            format!(
                "Office document path '{}' has no parent directory.",
                path.display()
            ),
        )
    })?;
    if !parent.is_dir() {
        return Err(package_error(
            "use.office.package_save_failed",
            format!(
                "Office document directory '{}' does not exist.",
                parent.display()
            ),
        ));
    }
    verify_write_target(path, &target, max_archive_bytes)?;
    let mut temporary = NamedTempFile::new_in(parent).map_err(|error| {
        package_error(
            "use.office.package_save_failed",
            format!("Failed to create an atomic Office save file: {error}"),
        )
    })?;
    {
        let mut writer = ZipWriter::new(temporary.as_file_mut());
        let options = SimpleFileOptions::default()
            .compression_method(CompressionMethod::Deflated)
            .unix_permissions(0o644);
        for (name, bytes) in parts {
            writer.start_file(name, options).map_err(|error| {
                package_error(
                    "use.office.package_save_failed",
                    format!("Failed to start OOXML package part '{name}': {error}"),
                )
            })?;
            writer.write_all(bytes).map_err(|error| {
                package_error(
                    "use.office.package_save_failed",
                    format!("Failed to write OOXML package part '{name}': {error}"),
                )
            })?;
        }
        writer.finish().map_err(|error| {
            package_error(
                "use.office.package_save_failed",
                format!(
                    "Failed to finish OOXML package '{}': {error}",
                    path.display()
                ),
            )
        })?;
    }
    temporary.as_file().sync_all().map_err(|error| {
        package_error(
            "use.office.package_save_failed",
            format!("Failed to sync OOXML package '{}': {error}", path.display()),
        )
    })?;
    let temporary_bytes = temporary
        .as_file()
        .metadata()
        .map_err(|error| {
            package_error(
                "use.office.package_save_failed",
                format!("Failed to inspect temporary OOXML package: {error}"),
            )
        })?
        .len();
    if temporary_bytes > max_archive_bytes {
        return Err(package_error(
            "use.office.package_too_large",
            format!("Saved OOXML package would exceed the {max_archive_bytes}-byte archive limit."),
        ));
    }
    verify_write_target(path, &target, max_archive_bytes)?;
    if matches!(target, WriteTarget::CreateNew) {
        temporary.persist_noclobber(path).map_err(|error| {
            if error.error.kind() == std::io::ErrorKind::AlreadyExists {
                package_exists(path)
            } else {
                package_error(
                    "use.office.package_save_failed",
                    format!(
                        "Failed to atomically create '{}': {}",
                        path.display(),
                        error.error
                    ),
                )
            }
        })?;
    } else {
        temporary.persist(path).map_err(|error| {
            package_error(
                "use.office.package_save_failed",
                format!(
                    "Failed to atomically replace '{}': {}",
                    path.display(),
                    error.error
                ),
            )
        })?;
    }
    sync_parent(parent, path)?;
    file_revision(path, max_archive_bytes)
}

fn verify_write_target(path: &Path, target: &WriteTarget, max_archive_bytes: u64) -> UseResult<()> {
    match target {
        WriteTarget::Replace => Ok(()),
        WriteTarget::Revision(expected) => ensure_revision(path, expected, max_archive_bytes),
        WriteTarget::CreateNew => match std::fs::symlink_metadata(path) {
            Ok(_) => Err(package_exists(path)),
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
            Err(error) => Err(package_error(
                "use.office.package_create_failed",
                format!(
                    "Failed to inspect native Office destination '{}': {error}",
                    path.display()
                ),
            )),
        },
    }
}

fn ensure_revision(
    path: &Path,
    expected: &PackageRevision,
    max_archive_bytes: u64,
) -> UseResult<()> {
    let current = file_revision(path, max_archive_bytes)
        .map_err(|error| save_conflict(path, expected, Some(error.message)))?;
    if &current == expected {
        Ok(())
    } else {
        Err(save_conflict(
            path,
            expected,
            Some(format!(
                "current revision is {} bytes with SHA-256 {}",
                current.archive_bytes, current.sha256
            )),
        ))
    }
}

fn file_revision(path: &Path, max_archive_bytes: u64) -> UseResult<PackageRevision> {
    let metadata = std::fs::symlink_metadata(path).map_err(|error| {
        package_error(
            "use.office.package_revision_failed",
            format!(
                "Failed to inspect Office document '{}': {error}",
                path.display()
            ),
        )
    })?;
    if !metadata.is_file()
        || metadata.file_type().is_symlink()
        || metadata.len() > max_archive_bytes
    {
        return Err(package_error(
            "use.office.package_revision_failed",
            format!(
                "Office document '{}' is missing, not regular, or exceeds the archive limit.",
                path.display()
            ),
        ));
    }
    let mut file = std::fs::File::open(path).map_err(|error| {
        package_error(
            "use.office.package_revision_failed",
            format!(
                "Failed to open Office document '{}': {error}",
                path.display()
            ),
        )
    })?;
    revision_from_reader(&mut file, metadata.len(), path)
}

fn revision_from_reader(
    reader: &mut std::fs::File,
    archive_bytes: u64,
    path: &Path,
) -> UseResult<PackageRevision> {
    let mut digest = Sha256::new();
    let mut buffer = [0_u8; 64 * 1024];
    loop {
        let read = reader.read(&mut buffer).map_err(|error| {
            package_error(
                "use.office.package_revision_failed",
                format!(
                    "Failed to hash Office document '{}': {error}",
                    path.display()
                ),
            )
        })?;
        if read == 0 {
            break;
        }
        digest.update(&buffer[..read]);
    }
    Ok(PackageRevision {
        archive_bytes,
        sha256: format!("{:x}", digest.finalize()),
    })
}

fn save_conflict(path: &Path, expected: &PackageRevision, current: Option<String>) -> UseError {
    package_error(
        "use.office.save_conflict",
        format!(
            "Office document '{}' changed after it was opened; refusing to overwrite it.",
            path.display()
        ),
    )
    .with_suggestion("Reopen the document, inspect the newer revision, and reapply the mutation.")
    .with_detail("expectedSha256", expected.sha256.clone())
    .with_detail("expectedBytes", expected.archive_bytes)
    .with_detail(
        "current",
        current.unwrap_or_else(|| "unavailable".to_string()),
    )
}

#[cfg(unix)]
fn sync_parent(parent: &Path, path: &Path) -> UseResult<()> {
    std::fs::File::open(parent)
        .and_then(|directory| directory.sync_all())
        .map_err(|error| {
            package_error(
                "use.office.package_save_failed",
                format!(
                    "Failed to sync the directory for '{}': {error}",
                    path.display()
                ),
            )
        })
}

#[cfg(not(unix))]
fn sync_parent(_parent: &Path, _path: &Path) -> UseResult<()> {
    Ok(())
}

fn validate_structure(kind: DocumentKind, parts: &BTreeMap<String, Arc<[u8]>>) -> UseResult<()> {
    for required in [
        CONTENT_TYPES_PART,
        ROOT_RELATIONSHIPS_PART,
        kind.main_part(),
    ] {
        if !parts.contains_key(required) {
            return Err(package_error(
                "use.office.package_part_missing",
                format!("OOXML package is missing required part '{required}'."),
            ));
        }
    }
    Ok(())
}

fn detect_kind(parts: &BTreeMap<String, Arc<[u8]>>) -> UseResult<DocumentKind> {
    let kinds: Vec<_> = [
        DocumentKind::Word,
        DocumentKind::Spreadsheet,
        DocumentKind::Presentation,
    ]
    .into_iter()
    .filter(|kind| parts.contains_key(kind.main_part()))
    .collect();
    match kinds.as_slice() {
        [kind] => Ok(*kind),
        [] => Err(package_error(
            "use.office.package_kind_unknown",
            "OOXML package has no supported Word, Spreadsheet, or Presentation main part.",
        )),
        _ => Err(package_error(
            "use.office.package_kind_ambiguous",
            "OOXML package contains more than one supported document main part.",
        )),
    }
}

fn require_extension_kind(path: &Path, kind: DocumentKind) -> UseResult<()> {
    let extension = path
        .extension()
        .and_then(|value| value.to_str())
        .map(str::to_ascii_lowercase);
    if extension.as_deref() == Some(kind.extension()) {
        return Ok(());
    }
    Err(package_error(
        "use.office.package_kind_mismatch",
        format!(
            "Office document '{}' contains a {:?} package but must use the .{} extension.",
            path.display(),
            kind,
            kind.extension()
        ),
    ))
}

fn kind_from_extension(path: &Path) -> UseResult<DocumentKind> {
    match path
        .extension()
        .and_then(|value| value.to_str())
        .map(str::to_ascii_lowercase)
        .as_deref()
    {
        Some("docx") => Ok(DocumentKind::Word),
        Some("xlsx") => Ok(DocumentKind::Spreadsheet),
        Some("pptx") => Ok(DocumentKind::Presentation),
        _ => Err(package_error(
            "use.office.package_extension_unsupported",
            format!(
                "Native Office creation requires a .docx, .xlsx, or .pptx destination; received '{}'.",
                path.display()
            ),
        )),
    }
}

fn archive_part_name(name: &str) -> UseResult<String> {
    if name.starts_with('/') {
        return Err(invalid_part(name));
    }
    validate_part_name(name)
}

fn query_part_name(name: &str) -> UseResult<String> {
    let name = name.strip_prefix('/').unwrap_or(name);
    validate_part_name(name)
}

fn validate_part_name(name: &str) -> UseResult<String> {
    if name.is_empty()
        || name.contains('\\')
        || name.contains('\0')
        || name.chars().any(char::is_control)
        || name
            .split('/')
            .any(|segment| segment.is_empty() || segment == "." || segment == "..")
    {
        return Err(invalid_part(name));
    }
    Ok(name.to_string())
}

fn invalid_part(name: &str) -> UseError {
    package_error(
        "use.office.package_part_invalid",
        format!("OOXML package part name '{name}' is unsafe or invalid."),
    )
}

fn duplicate_part(name: &str) -> UseError {
    package_error(
        "use.office.package_part_duplicate",
        format!("OOXML package contains an ambiguous duplicate part '{name}'."),
    )
}

fn enforce_compression_ratio(
    name: &str,
    expanded: u64,
    compressed: u64,
    limits: PackageLimits,
) -> UseResult<()> {
    if expanded < COMPRESSION_RATIO_MINIMUM_BYTES {
        return Ok(());
    }
    if compressed == 0 || expanded / compressed.max(1) > limits.max_compression_ratio {
        return Err(package_error(
            "use.office.package_compression_ratio",
            format!(
                "OOXML package part '{name}' exceeds the {}x compression-ratio limit.",
                limits.max_compression_ratio
            ),
        ));
    }
    Ok(())
}

fn package_bytes(parts: &BTreeMap<String, Arc<[u8]>>) -> UseResult<u64> {
    parts.values().try_fold(0_u64, |total, part| {
        let len = u64::try_from(part.len()).map_err(|_| package_size_error(u64::MAX))?;
        total
            .checked_add(len)
            .ok_or_else(|| package_size_error(u64::MAX))
    })
}

fn validate_part_limits(
    parts: &BTreeMap<String, Arc<[u8]>>,
    limits: PackageLimits,
) -> UseResult<()> {
    if parts.len() > limits.max_entries {
        return Err(package_error(
            "use.office.package_entry_limit",
            format!(
                "OOXML package contains {} parts; the limit is {}.",
                parts.len(),
                limits.max_entries
            ),
        ));
    }
    for (name, bytes) in parts {
        let bytes = u64::try_from(bytes.len()).map_err(|_| {
            package_error(
                "use.office.package_part_too_large",
                format!("OOXML package part '{name}' is too large for this platform."),
            )
        })?;
        if bytes > limits.max_part_bytes {
            return Err(package_error(
                "use.office.package_part_too_large",
                format!(
                    "OOXML package part '{name}' exceeds the {}-byte limit.",
                    limits.max_part_bytes
                ),
            ));
        }
    }
    if package_bytes(parts)? > limits.max_uncompressed_bytes {
        return Err(package_size_error(limits.max_uncompressed_bytes));
    }
    Ok(())
}

fn validate_limits(limits: PackageLimits) -> UseResult<()> {
    if limits.max_archive_bytes == 0
        || limits.max_entries == 0
        || limits.max_part_bytes == 0
        || limits.max_uncompressed_bytes == 0
        || limits.max_compression_ratio == 0
        || limits.max_part_bytes > limits.max_uncompressed_bytes
    {
        return Err(package_error(
            "use.office.package_limits_invalid",
            "Native Office package limits must be positive and internally consistent.",
        ));
    }
    Ok(())
}

fn package_size_error(limit: u64) -> UseError {
    package_error(
        "use.office.package_too_large",
        format!("OOXML package expands beyond the {limit}-byte limit."),
    )
}

fn package_exists(path: &Path) -> UseError {
    package_error(
        "use.office.package_exists",
        format!(
            "Native Office destination '{}' already exists; refusing to overwrite it.",
            path.display()
        ),
    )
    .with_suggestion("Choose a new destination or mutate the existing document explicitly.")
}

fn absolute(path: &Path) -> UseResult<PathBuf> {
    if path.is_absolute() {
        return Ok(path.to_path_buf());
    }
    std::env::current_dir()
        .map(|directory| directory.join(path))
        .map_err(|error| {
            package_error(
                "use.office.path_resolution_failed",
                format!("Failed to resolve Office document path: {error}"),
            )
        })
}

fn package_error(code: &str, message: impl Into<String>) -> UseError {
    office_error(code, message)
}

impl DocumentKind {
    fn extension(self) -> &'static str {
        match self {
            Self::Word => "docx",
            Self::Spreadsheet => "xlsx",
            Self::Presentation => "pptx",
        }
    }

    pub(crate) fn main_part(self) -> &'static str {
        match self {
            Self::Word => "word/document.xml",
            Self::Spreadsheet => "xl/workbook.xml",
            Self::Presentation => "ppt/presentation.xml",
        }
    }
}
