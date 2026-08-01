use std::fmt::{Debug, Formatter};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::Duration;

#[cfg(test)]
use std::sync::atomic::{AtomicUsize, Ordering};

use a3s_use_core::UseResult;
use async_trait::async_trait;
use sha2::{Digest, Sha256};

use super::engine::{PdfiumEngine, PDFIUM_ENGINE_VERSION};
use super::{
    pdf_page_identity_mismatch, NativeOfficePdfOutline, NativeOfficePdfOutlineOptions,
    NativeOfficePdfPageGeometry, NativeOfficePdfPageInventory, NativeOfficePdfPageInventoryOptions,
    NativeOfficePdfPageTextLayer, NativeOfficePdfTextLayerOptions,
    MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES,
};
use crate::layout::pptx_image::io::{
    ensure_output_available, hash_regular_file, publish_output, stage_output,
    validate_published_output, verify_source_revision,
};
use crate::layout::{
    is_sha256, layout_error, validate_revision, validate_unit, NativeOfficeLayoutAuthority,
    NativeOfficeLayoutEnvironment, NativeOfficeLayoutInspection, NativeOfficeLayoutProfile,
    NativeOfficeLayoutReceipt, NativeOfficeLayoutRenderRequest, NativeOfficeLayoutRenderer,
    NativeOfficeLayoutRendererDescriptor, NativeOfficeLayoutSourceKind, MAX_LAYOUT_OUTPUT_BYTES,
    MAX_LAYOUT_TIMEOUT_MS, NATIVE_OFFICE_LAYOUT_PROFILE_SCHEMA_VERSION,
};
use crate::{NativeOfficeUnit, NativeOfficeUnitLocator, PackageRevision};

const RENDERER_ID: &str = "a3s-office-native-pdfium";
const ENGINE_NAME: &str = "pdfium";
const DEVICE_SCALE_FACTOR_MILLI: u32 = 1_000;

/// Native, browser-neutral PDF page renderer backed by a host-supplied PDFium 7881 library.
///
/// The provider never downloads a runtime. The first successful construction
/// binds one content-addressed PDFium binary process-wide; later constructions
/// must name the same binary identity.
#[derive(Clone)]
pub struct NativeOfficePdfiumLayoutRenderer {
    engine: Arc<PdfiumEngine>,
    font_manifest_sha256: String,
    #[cfg(test)]
    inventory_calls: Arc<AtomicUsize>,
}

impl NativeOfficePdfiumLayoutRenderer {
    /// Loads an explicit PDFium 7881 library and binds the render profile to a
    /// caller-maintained host font manifest.
    pub async fn from_library(
        library_path: impl AsRef<Path>,
        font_manifest_sha256: impl Into<String>,
    ) -> UseResult<Self> {
        let font_manifest_sha256 = font_manifest_sha256.into();
        if !is_sha256(&font_manifest_sha256) {
            return Err(layout_error(
                "use.office.layout_renderer_invalid",
                "The PDFium renderer requires a SHA-256 host font manifest.",
            ));
        }
        let engine = PdfiumEngine::from_library(library_path.as_ref()).await?;
        let renderer = Self {
            engine,
            font_manifest_sha256,
            #[cfg(test)]
            inventory_calls: Arc::new(AtomicUsize::new(0)),
        };
        renderer.descriptor().validate()?;
        Ok(renderer)
    }

    /// Hashes and validates one regular PDF source under an explicit byte and
    /// deadline bound before any inventory state is created.
    pub async fn source_revision(
        &self,
        source_path: impl AsRef<Path>,
        max_source_bytes: u64,
        timeout_ms: u64,
    ) -> UseResult<PackageRevision> {
        validate_source_bounds(max_source_bytes, timeout_ms)?;
        let source_path = absolute(source_path.as_ref())?;
        match tokio::time::timeout(
            Duration::from_millis(timeout_ms),
            observe_source_revision(&source_path, max_source_bytes),
        )
        .await
        {
            Ok(result) => result,
            Err(_) => Err(layout_timeout(timeout_ms)),
        }
    }

    /// Inventories every PDF page without truncation and reports exact source,
    /// box, rotation, physical-surface, and target-pixel geometry.
    pub async fn inventory_pages(
        &self,
        source_path: impl AsRef<Path>,
        source_revision: PackageRevision,
        options: NativeOfficePdfPageInventoryOptions,
    ) -> UseResult<NativeOfficePdfPageInventory> {
        options.validate()?;
        self.descriptor().validate()?;
        validate_pdf_revision(&source_revision, options.max_source_bytes)?;
        #[cfg(test)]
        self.inventory_calls.fetch_add(1, Ordering::Relaxed);
        let source_path = absolute(source_path.as_ref())?;
        let timeout_ms = options.timeout_ms;
        let engine = Arc::clone(&self.engine);
        let revision_for_task = source_revision.clone();
        let source_for_task = source_path.clone();
        let task = async move {
            let bytes = read_source_bytes(
                &source_for_task,
                &revision_for_task,
                options.max_source_bytes,
            )
            .await?;
            let pages = tokio::task::spawn_blocking(move || {
                engine.inspect_pages(bytes, options.max_pages, options.dpi_milli)
            })
            .await
            .map_err(|_| pdfium_task_failed())??;
            verify_source_revision(&source_for_task, &revision_for_task).await?;
            Ok::<_, a3s_use_core::UseError>(pages)
        };
        let pages = match tokio::time::timeout(Duration::from_millis(timeout_ms), task).await {
            Ok(result) => result?,
            Err(_) => return Err(layout_timeout(timeout_ms)),
        };
        let inventory = NativeOfficePdfPageInventory {
            kind: NativeOfficeLayoutSourceKind::Pdf,
            source_revision,
            max_pages: options.max_pages,
            total_pages: pages.len(),
            dpi_milli: options.dpi_milli,
            pages,
        };
        inventory.validate()?;
        Ok(inventory)
    }

    /// Extracts one complete, bounded native text layer from a previously
    /// inventoried PDF page without rerunning page inventory.
    pub async fn extract_page_text(
        &self,
        source_path: impl AsRef<Path>,
        inventory: &NativeOfficePdfPageInventory,
        unit: NativeOfficeUnit,
        options: NativeOfficePdfTextLayerOptions,
    ) -> UseResult<NativeOfficePdfPageTextLayer> {
        self.descriptor().validate()?;
        inventory.validate()?;
        options.validate()?;
        inventory.validated_page(&unit)?;
        validate_pdf_revision(
            &inventory.source_revision,
            MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES,
        )?;
        let source_path = absolute(source_path.as_ref())?;
        let timeout_ms = options.timeout_ms;
        let engine = Arc::clone(&self.engine);
        let inventory_for_task = inventory.clone();
        let revision_for_task = inventory.source_revision.clone();
        let source_for_task = source_path.clone();
        let task = async move {
            let bytes = read_source_bytes(
                &source_for_task,
                &revision_for_task,
                MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES,
            )
            .await?;
            let layer = tokio::task::spawn_blocking(move || {
                engine.extract_page_text(bytes, &inventory_for_task, &unit, options)
            })
            .await
            .map_err(|_| pdfium_task_failed())??;
            verify_source_revision(&source_for_task, &revision_for_task).await?;
            Ok::<_, a3s_use_core::UseError>(layer)
        };
        let layer = match tokio::time::timeout(Duration::from_millis(timeout_ms), task).await {
            Ok(result) => result?,
            Err(_) => return Err(layout_timeout(timeout_ms)),
        };
        layer.validate(inventory)?;
        Ok(layer)
    }

    /// Extracts one complete, bounded native document outline against a
    /// previously admitted page inventory without rerunning inventory.
    pub async fn extract_outline(
        &self,
        source_path: impl AsRef<Path>,
        inventory: &NativeOfficePdfPageInventory,
        options: NativeOfficePdfOutlineOptions,
    ) -> UseResult<NativeOfficePdfOutline> {
        self.descriptor().validate()?;
        inventory.validate()?;
        options.validate()?;
        validate_pdf_revision(
            &inventory.source_revision,
            MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES,
        )?;
        let source_path = absolute(source_path.as_ref())?;
        let timeout_ms = options.timeout_ms;
        let engine = Arc::clone(&self.engine);
        let inventory_for_task = inventory.clone();
        let revision_for_task = inventory.source_revision.clone();
        let source_for_task = source_path.clone();
        let task = async move {
            let bytes = read_source_bytes(
                &source_for_task,
                &revision_for_task,
                MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES,
            )
            .await?;
            let outline = tokio::task::spawn_blocking(move || {
                engine.extract_outline(bytes, &inventory_for_task, options)
            })
            .await
            .map_err(|_| pdfium_task_failed())??;
            verify_source_revision(&source_for_task, &revision_for_task).await?;
            Ok::<_, a3s_use_core::UseError>(outline)
        };
        let outline = match tokio::time::timeout(Duration::from_millis(timeout_ms), task).await {
            Ok(result) => result?,
            Err(_) => return Err(layout_timeout(timeout_ms)),
        };
        outline.validate(inventory)?;
        Ok(outline)
    }

    /// Inventories the PDF and freezes the exact profile for one selected page.
    /// Reuse [`Self::inspect_inventoried_page`] after a complete inventory has
    /// already been admitted.
    pub async fn inspect_page(
        &self,
        source_path: impl AsRef<Path>,
        source_revision: PackageRevision,
        unit: NativeOfficeUnit,
        environment: NativeOfficeLayoutEnvironment,
        options: NativeOfficePdfPageInventoryOptions,
    ) -> UseResult<NativeOfficeLayoutInspection> {
        validate_unit(&unit)?;
        environment.validate()?;
        if !matches!(unit.locator, NativeOfficeUnitLocator::Page { .. }) {
            return Err(pdf_page_identity_mismatch());
        }
        let source_path = absolute(source_path.as_ref())?;
        let inventory = self
            .inventory_pages(&source_path, source_revision.clone(), options)
            .await?;
        self.inspect_inventoried_page(source_path, &inventory, unit, environment)
    }

    /// Freezes one page from a previously validated complete inventory without
    /// reopening or reinventorying the PDF. Rendering still revalidates the
    /// immutable source and the selected page's exact profile before output.
    pub fn inspect_inventoried_page(
        &self,
        source_path: impl AsRef<Path>,
        inventory: &NativeOfficePdfPageInventory,
        unit: NativeOfficeUnit,
        environment: NativeOfficeLayoutEnvironment,
    ) -> UseResult<NativeOfficeLayoutInspection> {
        self.descriptor().validate()?;
        validate_unit(&unit)?;
        environment.validate()?;
        validate_pdf_revision(
            &inventory.source_revision,
            MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES,
        )?;
        let source_path = absolute(source_path.as_ref())?;
        let page = inventory.validated_page(&unit)?;
        let profile = self.profile(page, environment, inventory.dpi_milli);
        profile.validate()?;
        Ok(NativeOfficeLayoutInspection {
            source_path,
            source_revision: inventory.source_revision.clone(),
            unit,
            profile,
        })
    }

    #[cfg(test)]
    pub(crate) fn inventory_call_count(&self) -> usize {
        self.inventory_calls.load(Ordering::Relaxed)
    }

    fn profile(
        &self,
        page: &NativeOfficePdfPageGeometry,
        environment: NativeOfficeLayoutEnvironment,
        dpi_milli: u32,
    ) -> NativeOfficeLayoutProfile {
        NativeOfficeLayoutProfile {
            schema_version: NATIVE_OFFICE_LAYOUT_PROFILE_SCHEMA_VERSION,
            authority: NativeOfficeLayoutAuthority::SourceLayout,
            renderer_id: RENDERER_ID.to_string(),
            renderer_version: env!("CARGO_PKG_VERSION").to_string(),
            engine_name: ENGINE_NAME.to_string(),
            engine_version: PDFIUM_ENGINE_VERSION.to_string(),
            engine_binary_sha256: self.engine.binary_sha256().to_string(),
            viewport_width_px: page.output_width_px,
            viewport_height_px: page.output_height_px,
            device_scale_factor_milli: DEVICE_SCALE_FACTOR_MILLI,
            dpi_x_milli: dpi_milli,
            dpi_y_milli: dpi_milli,
            surface_width_micrometers: page.surface_width_micrometers,
            surface_height_micrometers: page.surface_height_micrometers,
            locale: environment.locale,
            timezone: environment.timezone,
            font_manifest_sha256: self.font_manifest_sha256.clone(),
            renderer_config_sha256: renderer_config_sha256(),
            output_media_type: "image/png".to_string(),
            output_width_px: page.output_width_px,
            output_height_px: page.output_height_px,
            rotation_degrees: page.rotation_degrees,
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
        let bytes = read_source_bytes(
            &request.source_path,
            &request.source_revision,
            MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES,
        )
        .await?;
        let engine = Arc::clone(&self.engine);
        let unit = request.unit.clone();
        let dpi_milli = request.profile.dpi_x_milli;
        let max_output_bytes = request.max_output_bytes;
        let (page, png) = tokio::task::spawn_blocking(move || {
            engine.render_page(bytes, &unit, dpi_milli, max_output_bytes)
        })
        .await
        .map_err(|_| pdfium_task_failed())??;
        let expected_profile = self.profile(
            &page,
            NativeOfficeLayoutEnvironment::new(
                request.profile.locale.clone(),
                request.profile.timezone.clone(),
            ),
            request.profile.dpi_x_milli,
        );
        if request.profile != expected_profile {
            return Err(layout_error(
                "use.office.layout_profile_invalid",
                "The requested layout profile does not match the exact PDF page and renderer.",
            ));
        }
        let png_sha256 = format!("{:x}", Sha256::digest(&png));
        let staged = stage_output(&request.output, png).await?;
        before_final_source_check();
        verify_source_revision(&request.source_path, &request.source_revision).await?;
        publish_output(staged, &request.output).await?;
        let raster = validate_published_output(
            &request.output,
            request.max_output_bytes,
            page.output_width_px,
            page.output_height_px,
            &png_sha256,
            page.rotation_degrees,
        )
        .await?;
        verify_source_revision(&request.source_path, &request.source_revision).await?;

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
        let timeout_ms = request.timeout_ms;
        match tokio::time::timeout(
            Duration::from_millis(timeout_ms),
            self.render_inner_with_hook(request, before_final_source_check),
        )
        .await
        {
            Ok(result) => result,
            Err(_) => Err(layout_timeout(timeout_ms)),
        }
    }
}

impl Debug for NativeOfficePdfiumLayoutRenderer {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("NativeOfficePdfiumLayoutRenderer")
            .field("engine_binary_sha256", &self.engine.binary_sha256())
            .field("font_manifest_sha256", &self.font_manifest_sha256)
            .finish()
    }
}

#[async_trait]
impl NativeOfficeLayoutRenderer for NativeOfficePdfiumLayoutRenderer {
    fn descriptor(&self) -> NativeOfficeLayoutRendererDescriptor {
        NativeOfficeLayoutRendererDescriptor {
            id: RENDERER_ID.to_string(),
            version: env!("CARGO_PKG_VERSION").to_string(),
            sends_source_off_device: false,
        }
    }

    fn supports(&self, kind: NativeOfficeLayoutSourceKind) -> bool {
        kind == NativeOfficeLayoutSourceKind::Pdf
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
    renderer: &NativeOfficePdfiumLayoutRenderer,
    request: &NativeOfficeLayoutRenderRequest,
) -> UseResult<()> {
    renderer.descriptor().validate()?;
    validate_pdf_revision(&request.source_revision, MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES)?;
    validate_unit(&request.unit)?;
    request.profile.validate()?;
    validate_timeout(request.timeout_ms)?;
    let is_page = matches!(request.unit.locator, NativeOfficeUnitLocator::Page { .. });
    if !is_page
        || request.source_path.as_os_str().is_empty()
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
        || request.profile.engine_version != PDFIUM_ENGINE_VERSION
        || request.profile.engine_binary_sha256 != renderer.engine.binary_sha256()
        || request.profile.font_manifest_sha256 != renderer.font_manifest_sha256
        || request.profile.renderer_config_sha256 != renderer_config_sha256()
        || request.profile.device_scale_factor_milli != DEVICE_SCALE_FACTOR_MILLI
        || request.profile.dpi_x_milli != request.profile.dpi_y_milli
    {
        return Err(layout_error(
            "use.office.layout_request_invalid",
            "The PDFium layout request is incomplete, unbounded, or conflicts with its renderer.",
        ));
    }
    ensure_output_available(&request.output).await
}

async fn observe_source_revision(path: &Path, max_bytes: u64) -> UseResult<PackageRevision> {
    let metadata = tokio::fs::symlink_metadata(path)
        .await
        .map_err(|_| pdf_source_invalid())?;
    if !metadata.is_file()
        || metadata.file_type().is_symlink()
        || metadata.len() == 0
        || metadata.len() > max_bytes
    {
        return Err(pdf_source_invalid());
    }
    let sha256 = hash_regular_file(path, Some(metadata.len()))
        .await
        .map_err(|_| pdf_source_invalid())?;
    let revision = PackageRevision {
        archive_bytes: metadata.len(),
        sha256,
    };
    verify_source_revision(path, &revision).await?;
    Ok(revision)
}

async fn read_source_bytes(
    path: &Path,
    revision: &PackageRevision,
    max_source_bytes: u64,
) -> UseResult<Vec<u8>> {
    validate_pdf_revision(revision, max_source_bytes)?;
    verify_source_revision(path, revision).await?;
    let bytes = tokio::fs::read(path)
        .await
        .map_err(|_| pdf_source_invalid())?;
    if u64::try_from(bytes.len()).unwrap_or(u64::MAX) != revision.archive_bytes
        || format!("{:x}", Sha256::digest(&bytes)) != revision.sha256
    {
        return Err(layout_error(
            "use.office.layout_source_mutated",
            "The PDF layout source changed while its immutable bytes were being captured.",
        ));
    }
    verify_source_revision(path, revision).await?;
    Ok(bytes)
}

fn renderer_config_sha256() -> String {
    format!(
        "{:x}",
        Sha256::digest(
            b"a3s-office-pdfium-layout-v1\0pdfium-7881\0rgba8\0png\0white-background\0annotations\0forms\0no-fetch"
        )
    )
}

fn validate_pdf_revision(revision: &PackageRevision, max_source_bytes: u64) -> UseResult<()> {
    validate_revision(revision)?;
    if revision.archive_bytes <= max_source_bytes
        && max_source_bytes <= MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES
    {
        return Ok(());
    }
    Err(pdf_source_invalid())
}

fn validate_source_bounds(max_source_bytes: u64, timeout_ms: u64) -> UseResult<()> {
    if (1..=MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES).contains(&max_source_bytes) {
        validate_timeout(timeout_ms)
    } else {
        Err(pdf_source_invalid())
    }
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

fn absolute(path: &Path) -> UseResult<PathBuf> {
    if path.as_os_str().is_empty() {
        return Err(pdf_source_invalid());
    }
    if path.is_absolute() {
        return Ok(path.to_path_buf());
    }
    std::env::current_dir()
        .map(|directory| directory.join(path))
        .map_err(|_| pdf_source_invalid())
}

fn layout_timeout(timeout_ms: u64) -> a3s_use_core::UseError {
    layout_error(
        "use.office.layout_timeout",
        format!("Office layout rendering exceeded {timeout_ms} ms."),
    )
}

fn pdf_source_invalid() -> a3s_use_core::UseError {
    layout_error(
        "use.office.pdf_source_invalid",
        "The PDF source is missing, unsafe, empty, or exceeds its explicit byte bound.",
    )
}

fn pdfium_task_failed() -> a3s_use_core::UseError {
    layout_error(
        "use.office.pdfium_unavailable",
        "The bounded PDFium task did not complete successfully.",
    )
}
