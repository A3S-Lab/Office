mod candidate;
pub(in crate::layout) mod io;

use std::path::Path;
use std::time::Duration;

use a3s_use_core::UseResult;
use async_trait::async_trait;
use sha2::{Digest, Sha256};

use candidate::{dpi_milli, load_candidate, ExactPngSlide};
use io::{
    ensure_output_available, hash_regular_file, publish_output, stage_output,
    validate_published_output, verify_source_revision,
};

use super::{
    is_sha256, layout_error, validate_revision, validate_unit, NativeOfficeLayoutAuthority,
    NativeOfficeLayoutEnvironment, NativeOfficeLayoutInspection, NativeOfficeLayoutProfile,
    NativeOfficeLayoutReceipt, NativeOfficeLayoutRenderRequest, NativeOfficeLayoutRenderer,
    NativeOfficeLayoutRendererDescriptor, NativeOfficeLayoutSourceKind, MAX_LAYOUT_OUTPUT_BYTES,
    MAX_LAYOUT_TIMEOUT_MS, NATIVE_OFFICE_LAYOUT_PROFILE_SCHEMA_VERSION,
};
use crate::{NativeOfficeUnit, PackageRevision};

const RENDERER_ID: &str = "a3s-office-native-pptx-image";
const ENGINE_NAME: &str = "a3s-office-png-passthrough";
const ENGINE_VERSION: &str = "1";
const DEVICE_SCALE_FACTOR_MILLI: u32 = 1_000;

/// Native exact-layout renderer for opaque, full-slide PPTX PNG pictures.
///
/// This deliberately narrow provider performs no reflow, resampling, or SVG
/// conversion. It returns the embedded PNG bytes only when those bytes are the
/// complete visible slide surface. Every richer slide fails closed.
#[derive(Debug, Clone)]
pub struct NativeOfficePptxImageLayoutRenderer {
    engine_binary_sha256: String,
}

impl NativeOfficePptxImageLayoutRenderer {
    pub fn new(engine_binary_sha256: impl Into<String>) -> UseResult<Self> {
        let engine_binary_sha256 = engine_binary_sha256.into();
        if !is_sha256(&engine_binary_sha256) {
            return Err(layout_error(
                "use.office.layout_renderer_invalid",
                "The native PPTX image renderer requires the containing engine binary SHA-256.",
            ));
        }
        Ok(Self {
            engine_binary_sha256,
        })
    }

    /// Binds the provider identity to the executable containing this renderer.
    pub async fn from_current_executable() -> UseResult<Self> {
        let executable = std::env::current_exe().map_err(|_| {
            layout_error(
                "use.office.layout_renderer_invalid",
                "The native PPTX image renderer could not resolve its containing executable.",
            )
        })?;
        let sha256 = hash_regular_file(&executable, None).await.map_err(|_| {
            layout_error(
                "use.office.layout_renderer_invalid",
                "The native PPTX image renderer could not hash its containing executable.",
            )
        })?;
        Self::new(sha256)
    }

    /// Inspects one exact source unit and freezes every profile input before a
    /// render request can be issued.
    pub async fn inspect_unit(
        &self,
        source_path: impl AsRef<Path>,
        source_revision: PackageRevision,
        unit: NativeOfficeUnit,
        environment: NativeOfficeLayoutEnvironment,
        timeout_ms: u64,
    ) -> UseResult<NativeOfficeLayoutInspection> {
        validate_timeout(timeout_ms)?;
        self.descriptor().validate()?;
        validate_revision(&source_revision)?;
        validate_unit(&unit)?;
        environment.validate()?;
        let source_path = source_path.as_ref().to_path_buf();
        let candidate = match tokio::time::timeout(
            Duration::from_millis(timeout_ms),
            load_candidate(&source_path, &source_revision, &unit),
        )
        .await
        {
            Ok(result) => result?,
            Err(_) => return Err(layout_timeout(timeout_ms)),
        };
        verify_source_revision(&candidate.source_path, &source_revision).await?;
        let profile = self.profile(&candidate, environment);
        profile.validate()?;
        Ok(NativeOfficeLayoutInspection {
            source_path: candidate.source_path,
            source_revision,
            unit,
            profile,
        })
    }

    fn profile(
        &self,
        candidate: &ExactPngSlide,
        environment: NativeOfficeLayoutEnvironment,
    ) -> NativeOfficeLayoutProfile {
        NativeOfficeLayoutProfile {
            schema_version: NATIVE_OFFICE_LAYOUT_PROFILE_SCHEMA_VERSION,
            authority: NativeOfficeLayoutAuthority::SourceLayout,
            renderer_id: RENDERER_ID.to_string(),
            renderer_version: env!("CARGO_PKG_VERSION").to_string(),
            engine_name: ENGINE_NAME.to_string(),
            engine_version: ENGINE_VERSION.to_string(),
            engine_binary_sha256: self.engine_binary_sha256.clone(),
            viewport_width_px: candidate.width_px,
            viewport_height_px: candidate.height_px,
            device_scale_factor_milli: DEVICE_SCALE_FACTOR_MILLI,
            dpi_x_milli: dpi_milli(candidate.width_px, candidate.surface_width_micrometers),
            dpi_y_milli: dpi_milli(candidate.height_px, candidate.surface_height_micrometers),
            surface_width_micrometers: candidate.surface_width_micrometers,
            surface_height_micrometers: candidate.surface_height_micrometers,
            locale: environment.locale,
            timezone: environment.timezone,
            font_manifest_sha256: empty_font_manifest_sha256(),
            renderer_config_sha256: renderer_config_sha256(),
            output_media_type: "image/png".to_string(),
            output_width_px: candidate.width_px,
            output_height_px: candidate.height_px,
            rotation_degrees: 0,
        }
    }

    async fn render_inner(
        &self,
        request: NativeOfficeLayoutRenderRequest,
    ) -> UseResult<NativeOfficeLayoutReceipt> {
        self.render_inner_with_hook(request, || {}).await
    }

    async fn render_inner_with_hook<F>(
        &self,
        request: NativeOfficeLayoutRenderRequest,
        before_final_source_check: F,
    ) -> UseResult<NativeOfficeLayoutReceipt>
    where
        F: FnOnce() + Send,
    {
        ensure_output_available(&request.output).await?;
        let candidate = load_candidate(
            &request.source_path,
            &request.source_revision,
            &request.unit,
        )
        .await?;
        let expected_profile = self.profile(
            &candidate,
            NativeOfficeLayoutEnvironment::new(
                request.profile.locale.clone(),
                request.profile.timezone.clone(),
            ),
        );
        if request.profile != expected_profile {
            return Err(layout_error(
                "use.office.layout_profile_invalid",
                "The requested layout profile does not match the exact source slide and renderer.",
            ));
        }
        let byte_length = u64::try_from(candidate.png.len()).unwrap_or(u64::MAX);
        if byte_length == 0 || byte_length > request.max_output_bytes {
            return Err(layout_output_too_large(request.max_output_bytes));
        }

        let staged = stage_output(&request.output, candidate.png.clone()).await?;
        before_final_source_check();
        verify_source_revision(&candidate.source_path, &request.source_revision).await?;
        publish_output(staged, &request.output).await?;
        let raster = validate_published_output(
            &request.output,
            request.max_output_bytes,
            candidate.width_px,
            candidate.height_px,
            &candidate.png_sha256,
            0,
        )
        .await?;
        verify_source_revision(&candidate.source_path, &request.source_revision).await?;

        let profile_sha256 = request.profile.sha256()?;
        let receipt = NativeOfficeLayoutReceipt {
            source_revision: request.source_revision.clone(),
            render_input_sha256: request.source_revision.sha256,
            unit: request.unit,
            profile: request.profile,
            profile_sha256,
            raster,
        };
        receipt.validate()?;
        Ok(receipt)
    }

    #[cfg(test)]
    pub(crate) async fn render_with_before_final_source_check<F>(
        &self,
        request: NativeOfficeLayoutRenderRequest,
        before_final_source_check: F,
    ) -> UseResult<NativeOfficeLayoutReceipt>
    where
        F: FnOnce() + Send,
    {
        validate_request(self, &request).await?;
        self.render_inner_with_hook(request, before_final_source_check)
            .await
    }
}

#[async_trait]
impl NativeOfficeLayoutRenderer for NativeOfficePptxImageLayoutRenderer {
    fn descriptor(&self) -> NativeOfficeLayoutRendererDescriptor {
        NativeOfficeLayoutRendererDescriptor {
            id: RENDERER_ID.to_string(),
            version: env!("CARGO_PKG_VERSION").to_string(),
            sends_source_off_device: false,
        }
    }

    fn supports(&self, kind: NativeOfficeLayoutSourceKind) -> bool {
        kind == NativeOfficeLayoutSourceKind::Presentation
    }

    async fn render(
        &self,
        request: NativeOfficeLayoutRenderRequest,
    ) -> UseResult<NativeOfficeLayoutReceipt> {
        validate_request(self, &request).await?;
        let timeout_ms = request.timeout_ms;
        match tokio::time::timeout(
            Duration::from_millis(timeout_ms),
            self.render_inner(request),
        )
        .await
        {
            Ok(result) => result,
            Err(_) => Err(layout_timeout(timeout_ms)),
        }
    }
}

async fn validate_request(
    renderer: &NativeOfficePptxImageLayoutRenderer,
    request: &NativeOfficeLayoutRenderRequest,
) -> UseResult<()> {
    renderer.descriptor().validate()?;
    validate_revision(&request.source_revision)?;
    validate_unit(&request.unit)?;
    request.profile.validate()?;
    validate_timeout(request.timeout_ms)?;
    if request.source_path.as_os_str().is_empty()
        || request.output.as_os_str().is_empty()
        || request.output == Path::new("-")
        || !request
            .output
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("png"))
        || !(1..=MAX_LAYOUT_OUTPUT_BYTES).contains(&request.max_output_bytes)
        || request.profile.renderer_id != RENDERER_ID
        || request.profile.renderer_version != env!("CARGO_PKG_VERSION")
        || request.profile.engine_name != ENGINE_NAME
        || request.profile.engine_version != ENGINE_VERSION
        || request.profile.engine_binary_sha256 != renderer.engine_binary_sha256
        || request.profile.font_manifest_sha256 != empty_font_manifest_sha256()
        || request.profile.renderer_config_sha256 != renderer_config_sha256()
        || request.profile.device_scale_factor_milli != DEVICE_SCALE_FACTOR_MILLI
    {
        return Err(layout_error(
            "use.office.layout_request_invalid",
            "The native PPTX layout request is incomplete, unbounded, or conflicts with its renderer.",
        ));
    }
    ensure_output_available(&request.output).await
}

fn validate_timeout(timeout_ms: u64) -> UseResult<()> {
    if (1..=MAX_LAYOUT_TIMEOUT_MS).contains(&timeout_ms) {
        return Ok(());
    }
    Err(layout_error(
        "use.office.layout_timeout_invalid",
        format!("Office layout timeout must be between 1 and {MAX_LAYOUT_TIMEOUT_MS} ms."),
    ))
}

fn empty_font_manifest_sha256() -> String {
    format!("{:x}", Sha256::digest(b"a3s-office-font-manifest-v1\0[]"))
}

fn renderer_config_sha256() -> String {
    format!(
        "{:x}",
        Sha256::digest(
            b"a3s-office-pptx-image-layout-v1\0opaque-png\0full-surface\0no-crop\0no-transform\0no-resampling"
        )
    )
}

fn layout_timeout(timeout_ms: u64) -> a3s_use_core::UseError {
    layout_error(
        "use.office.layout_timeout",
        format!("Office layout rendering exceeded {timeout_ms} ms."),
    )
}

fn layout_output_too_large(max_output_bytes: u64) -> a3s_use_core::UseError {
    layout_error(
        "use.office.layout_output_too_large",
        format!("Office layout raster exceeds the {max_output_bytes}-byte output limit."),
    )
}
