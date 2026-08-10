#[cfg(feature = "pdfium")]
mod pdfium;
mod pptx_image;

use std::path::PathBuf;

use a3s_use_core::{UseError, UseResult};
use async_trait::async_trait;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use crate::{DocumentKind, NativeOfficeUnit, PackageRevision};

#[cfg(feature = "pdfium")]
pub use pdfium::{
    NativeOfficePdfOutline, NativeOfficePdfOutlineEntry, NativeOfficePdfOutlineOptions,
    NativeOfficePdfPageBox, NativeOfficePdfPageGeometry, NativeOfficePdfPageInventory,
    NativeOfficePdfPageInventoryOptions, NativeOfficePdfPageTextLayer,
    NativeOfficePdfTextCharacter, NativeOfficePdfTextLayerOptions,
    NativeOfficePdfiumLayoutRenderer, DEFAULT_NATIVE_OFFICE_PDF_OUTLINE_DEPTH,
    DEFAULT_NATIVE_OFFICE_PDF_OUTLINE_ENTRIES, DEFAULT_NATIVE_OFFICE_PDF_OUTLINE_TITLE_BYTES,
    DEFAULT_NATIVE_OFFICE_PDF_PAGE_LIMIT, DEFAULT_NATIVE_OFFICE_PDF_TEXT_CHARACTERS,
    DEFAULT_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES, MAX_NATIVE_OFFICE_PDF_OUTLINE_DEPTH,
    MAX_NATIVE_OFFICE_PDF_OUTLINE_ENTRIES, MAX_NATIVE_OFFICE_PDF_OUTLINE_TITLE_BYTES,
    MAX_NATIVE_OFFICE_PDF_PAGE_LIMIT, MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES,
    MAX_NATIVE_OFFICE_PDF_TEXT_CHARACTERS, MAX_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES,
    NATIVE_OFFICE_PDF_TEXT_SCHEMA_VERSION,
};
pub use pptx_image::NativeOfficePptxImageLayoutRenderer;

/// Version of the canonical native Office layout-profile contract.
pub const NATIVE_OFFICE_LAYOUT_PROFILE_SCHEMA_VERSION: u32 = 1;

pub(crate) const MAX_LAYOUT_TEXT_BYTES: usize = 128;
pub(crate) const MAX_LAYOUT_TIMEOUT_MS: u64 = 120_000;
pub(crate) const MAX_LAYOUT_OUTPUT_BYTES: u64 = 64 * 1024 * 1024;
const LAYOUT_PROFILE_DOMAIN: &[u8] = b"a3s-office-layout-profile-v1\0";
const MAX_DEVICE_SCALE_FACTOR_MILLI: u32 = 16_000;
const MAX_DPI_MILLI: u32 = 2_400_000;

/// Fidelity authority carried by one visual receipt.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeLayoutAuthority {
    /// Pixels preserve the renderer's declared source-layout surface.
    SourceLayout,
    /// Pixels are a semantic preview and cannot ground source coordinates.
    SemanticPreview,
}

/// Source formats admitted by the read-only layout boundary.
///
/// This remains separate from [`DocumentKind`], which identifies editable
/// OOXML packages owned by the native Office document model.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeLayoutSourceKind {
    Word,
    Spreadsheet,
    Presentation,
    #[cfg(feature = "pdfium")]
    Pdf,
}

impl From<DocumentKind> for NativeOfficeLayoutSourceKind {
    fn from(kind: DocumentKind) -> Self {
        match kind {
            DocumentKind::Word => Self::Word,
            DocumentKind::Spreadsheet => Self::Spreadsheet,
            DocumentKind::Presentation => Self::Presentation,
        }
    }
}

/// Stable identity and transfer behavior of one layout renderer.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeLayoutRendererDescriptor {
    pub id: String,
    pub version: String,
    pub sends_source_off_device: bool,
}

impl NativeOfficeLayoutRendererDescriptor {
    pub fn validate(&self) -> UseResult<()> {
        if valid_identifier(&self.id) && bounded_text(&self.version) {
            return Ok(());
        }
        Err(layout_error(
            "use.office.layout_renderer_invalid",
            "An Office layout renderer requires bounded stable identity and version fields.",
        ))
    }
}

/// Explicit locale and timezone inputs used while preparing a render profile.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeLayoutEnvironment {
    pub locale: String,
    pub timezone: String,
}

impl NativeOfficeLayoutEnvironment {
    pub fn new(locale: impl Into<String>, timezone: impl Into<String>) -> Self {
        Self {
            locale: locale.into(),
            timezone: timezone.into(),
        }
    }

    pub(crate) fn validate(&self) -> UseResult<()> {
        if bounded_text(&self.locale) && bounded_text(&self.timezone) {
            return Ok(());
        }
        Err(invalid_profile())
    }
}

/// Complete content-addressed inputs for one source-layout raster.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeLayoutProfile {
    pub schema_version: u32,
    pub authority: NativeOfficeLayoutAuthority,
    pub renderer_id: String,
    pub renderer_version: String,
    pub engine_name: String,
    pub engine_version: String,
    pub engine_binary_sha256: String,
    pub viewport_width_px: u32,
    pub viewport_height_px: u32,
    pub device_scale_factor_milli: u32,
    pub dpi_x_milli: u32,
    pub dpi_y_milli: u32,
    pub surface_width_micrometers: u32,
    pub surface_height_micrometers: u32,
    pub locale: String,
    pub timezone: String,
    pub font_manifest_sha256: String,
    pub renderer_config_sha256: String,
    pub output_media_type: String,
    pub output_width_px: u32,
    pub output_height_px: u32,
    pub rotation_degrees: u16,
}

impl NativeOfficeLayoutProfile {
    pub fn validate(&self) -> UseResult<()> {
        if self.schema_version != NATIVE_OFFICE_LAYOUT_PROFILE_SCHEMA_VERSION
            || self.authority != NativeOfficeLayoutAuthority::SourceLayout
            || !valid_identifier(&self.renderer_id)
            || !bounded_text(&self.renderer_version)
            || !bounded_text(&self.engine_name)
            || !bounded_text(&self.engine_version)
            || !is_sha256(&self.engine_binary_sha256)
            || self.viewport_width_px == 0
            || self.viewport_height_px == 0
            || !(1..=MAX_DEVICE_SCALE_FACTOR_MILLI).contains(&self.device_scale_factor_milli)
            || !(1_000..=MAX_DPI_MILLI).contains(&self.dpi_x_milli)
            || !(1_000..=MAX_DPI_MILLI).contains(&self.dpi_y_milli)
            || self.surface_width_micrometers == 0
            || self.surface_height_micrometers == 0
            || !dimension_matches_dpi(
                self.output_width_px,
                self.dpi_x_milli,
                self.surface_width_micrometers,
            )
            || !dimension_matches_dpi(
                self.output_height_px,
                self.dpi_y_milli,
                self.surface_height_micrometers,
            )
            || !bounded_text(&self.locale)
            || !bounded_text(&self.timezone)
            || !is_sha256(&self.font_manifest_sha256)
            || !is_sha256(&self.renderer_config_sha256)
            || self.output_media_type != "image/png"
            || self.output_width_px == 0
            || self.output_height_px == 0
            || !matches!(self.rotation_degrees, 0 | 90 | 180 | 270)
        {
            return Err(invalid_profile());
        }
        Ok(())
    }

    /// Hashes the validated profile with a versioned domain separator.
    pub fn sha256(&self) -> UseResult<String> {
        self.validate()?;
        let encoded = serde_json::to_vec(self).map_err(|_| invalid_profile())?;
        let mut digest = Sha256::new();
        digest.update(LAYOUT_PROFILE_DOMAIN);
        digest.update(encoded);
        Ok(format!("{:x}", digest.finalize()))
    }
}

/// Prepared, source-bound inputs for one exact-unit render.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeLayoutInspection {
    pub source_path: PathBuf,
    pub source_revision: PackageRevision,
    pub unit: NativeOfficeUnit,
    pub profile: NativeOfficeLayoutProfile,
}

impl NativeOfficeLayoutInspection {
    pub fn into_render_request(
        self,
        output: impl Into<PathBuf>,
        max_output_bytes: u64,
        timeout_ms: u64,
    ) -> NativeOfficeLayoutRenderRequest {
        NativeOfficeLayoutRenderRequest {
            source_path: self.source_path,
            source_revision: self.source_revision,
            unit: self.unit,
            output: output.into(),
            max_output_bytes,
            timeout_ms,
            profile: self.profile,
        }
    }
}

/// Fully explicit exact-unit source-layout request.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeLayoutRenderRequest {
    pub source_path: PathBuf,
    pub source_revision: PackageRevision,
    pub unit: NativeOfficeUnit,
    pub output: PathBuf,
    pub max_output_bytes: u64,
    pub timeout_ms: u64,
    pub profile: NativeOfficeLayoutProfile,
}

/// Rehashed pixels published by one layout renderer.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeLayoutRaster {
    pub output_path: PathBuf,
    pub media_type: String,
    pub width_px: u32,
    pub height_px: u32,
    pub byte_length: u64,
    pub sha256: String,
    pub rotation_degrees: u16,
}

/// Receipt binding one immutable Office source unit to exact output pixels.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeLayoutReceipt {
    pub source_revision: PackageRevision,
    pub render_input_sha256: String,
    pub unit: NativeOfficeUnit,
    pub profile: NativeOfficeLayoutProfile,
    pub profile_sha256: String,
    pub raster: NativeOfficeLayoutRaster,
}

impl NativeOfficeLayoutReceipt {
    pub fn validate(&self) -> UseResult<()> {
        validate_revision(&self.source_revision)?;
        validate_unit(&self.unit)?;
        self.profile.validate()?;
        if self.render_input_sha256 != self.source_revision.sha256
            || self.profile.sha256()? != self.profile_sha256
            || self.raster.output_path.as_os_str().is_empty()
            || self.raster.media_type != self.profile.output_media_type
            || self.raster.width_px != self.profile.output_width_px
            || self.raster.height_px != self.profile.output_height_px
            || self.raster.rotation_degrees != self.profile.rotation_degrees
            || self.raster.byte_length == 0
            || !is_sha256(&self.raster.sha256)
        {
            return Err(layout_error(
                "use.office.layout_receipt_invalid",
                "The Office layout receipt does not bind one source, unit, profile, and raster.",
            ));
        }
        Ok(())
    }
}

/// Browser-neutral boundary for exact-unit Office layout rendering.
#[async_trait]
pub trait NativeOfficeLayoutRenderer: Send + Sync {
    fn descriptor(&self) -> NativeOfficeLayoutRendererDescriptor;
    fn supports(&self, kind: NativeOfficeLayoutSourceKind) -> bool;

    async fn render(
        &self,
        request: NativeOfficeLayoutRenderRequest,
    ) -> UseResult<NativeOfficeLayoutReceipt>;
}

pub(crate) fn validate_revision(revision: &PackageRevision) -> UseResult<()> {
    if revision.archive_bytes > 0 && is_sha256(&revision.sha256) {
        return Ok(());
    }
    Err(layout_error(
        "use.office.layout_source_invalid",
        "An Office layout request requires an immutable source byte length and SHA-256.",
    ))
}

pub(crate) fn validate_unit(unit: &NativeOfficeUnit) -> UseResult<()> {
    let ordinal_matches = match &unit.locator {
        crate::NativeOfficeUnitLocator::Document => unit.ordinal == 1 && unit.path == "/",
        crate::NativeOfficeUnitLocator::Worksheet { index, name } => {
            unit.ordinal == *index && !name.is_empty() && unit.path == format!("/{name}")
        }
        crate::NativeOfficeUnitLocator::Slide { number } => {
            unit.ordinal == *number && unit.path == format!("/slide[{number}]")
        }
        #[cfg(feature = "pdfium")]
        crate::NativeOfficeUnitLocator::Page { number } => {
            unit.ordinal == *number && unit.path == format!("/page[{number}]")
        }
    };
    if unit.ordinal > 0 && ordinal_matches && !unit.path.chars().any(char::is_control) {
        return Ok(());
    }
    Err(layout_error(
        "use.office.layout_unit_mismatch",
        "The Office layout unit locator, ordinal, and semantic path do not agree.",
    ))
}

pub(crate) fn bounded_text(value: &str) -> bool {
    !value.is_empty()
        && value.trim() == value
        && value.len() <= MAX_LAYOUT_TEXT_BYTES
        && !value.chars().any(char::is_control)
}

pub(crate) fn is_sha256(value: &str) -> bool {
    value.len() == 64
        && value
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
}

pub(crate) fn dimension_matches_dpi(pixels: u32, dpi_milli: u32, micrometers: u32) -> bool {
    const MICROMETERS_PER_INCH_MILLI: u128 = 25_400_000;

    let declared = u128::from(micrometers) * u128::from(dpi_milli);
    let derived = u128::from(pixels) * MICROMETERS_PER_INCH_MILLI;
    // The physical size and pixel count are independently rounded from the
    // same exact surface. Their cross-multiplied difference can therefore
    // contain up to half a micrometer plus half a pixel of quantization.
    let quantization_tolerance = (u128::from(dpi_milli) + MICROMETERS_PER_INCH_MILLI).div_ceil(2);
    declared.abs_diff(derived) <= quantization_tolerance
}

pub(crate) fn layout_error(code: &'static str, message: impl Into<String>) -> UseError {
    UseError::new(code, message)
}

fn invalid_profile() -> UseError {
    layout_error(
        "use.office.layout_profile_invalid",
        "An Office source-layout profile must identify every layout-changing input and exact output.",
    )
}

fn valid_identifier(value: &str) -> bool {
    !value.is_empty()
        && value.len() <= MAX_LAYOUT_TEXT_BYTES
        && value.bytes().all(|byte| {
            byte.is_ascii_lowercase() || byte.is_ascii_digit() || matches!(byte, b'.' | b'-' | b'_')
        })
}
