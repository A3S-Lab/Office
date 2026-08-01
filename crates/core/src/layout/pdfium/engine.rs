use std::io::Write as _;
use std::path::Path;
use std::sync::{Arc, Mutex, OnceLock};

use a3s_use_core::UseResult;
use pdfium_render::prelude::{
    PdfPageBoundaryBox, PdfPageRenderRotation, PdfRenderConfig, Pdfium, PdfiumError,
    PdfiumInternalError,
};
use sha2::{Digest, Sha256};

use super::{
    invalid_page_geometry, millipoints_to_micrometers, millipoints_to_pixels,
    NativeOfficePdfOutline, NativeOfficePdfOutlineOptions, NativeOfficePdfPageBox,
    NativeOfficePdfPageGeometry, NativeOfficePdfPageInventory, NativeOfficePdfPageTextLayer,
    NativeOfficePdfTextLayerOptions,
};
use crate::layout::layout_error;
use crate::{NativeOfficeUnit, NativeOfficeUnitLocator};

pub(super) const PDFIUM_ENGINE_VERSION: &str = "chromium/7881";
const MAX_PDFIUM_LIBRARY_BYTES: u64 = 256 * 1024 * 1024;
const MAX_PDF_BITMAP_BYTES: u64 = 256 * 1024 * 1024;

static PDFIUM_ENGINE: OnceLock<Arc<PdfiumEngine>> = OnceLock::new();
static PDFIUM_INIT_LOCK: Mutex<()> = Mutex::new(());

pub(super) struct PdfiumEngine {
    pdfium: Pdfium,
    binary_sha256: String,
    operation_lock: Mutex<()>,
    _staged_library: tempfile::TempDir,
}

impl PdfiumEngine {
    pub(super) async fn from_library(path: &Path) -> UseResult<Arc<Self>> {
        let bytes = read_library(path).await?;
        let binary_sha256 = format!("{:x}", Sha256::digest(&bytes));
        if let Some(engine) = PDFIUM_ENGINE.get() {
            return engine_for_identity(engine, &binary_sha256);
        }

        tokio::task::spawn_blocking(move || initialize_engine(bytes, binary_sha256))
            .await
            .map_err(|_| pdfium_library_invalid())?
    }

    pub(super) fn binary_sha256(&self) -> &str {
        &self.binary_sha256
    }

    pub(super) fn inspect_pages(
        &self,
        bytes: Vec<u8>,
        max_pages: usize,
        dpi_milli: u32,
    ) -> UseResult<Vec<NativeOfficePdfPageGeometry>> {
        let _guard = self
            .operation_lock
            .lock()
            .map_err(|_| pdfium_unavailable())?;
        let document = self
            .pdfium
            .load_pdf_from_byte_vec(bytes, None)
            .map_err(map_load_error)?;
        let total_pages = page_count(document.pages().len())?;
        validate_page_count(total_pages, max_pages)?;

        let mut pages = Vec::with_capacity(total_pages);
        for offset in 0..total_pages {
            let index =
                i32::try_from(offset).map_err(|_| pdf_page_limit(total_pages, max_pages))?;
            let page = document.pages().get(index).map_err(map_page_error)?;
            pages.push(observe_page(&page, offset, dpi_milli)?);
        }
        Ok(pages)
    }

    pub(super) fn render_page(
        &self,
        bytes: Vec<u8>,
        unit: &NativeOfficeUnit,
        dpi_milli: u32,
        max_output_bytes: u64,
    ) -> UseResult<(NativeOfficePdfPageGeometry, Vec<u8>)> {
        let _guard = self
            .operation_lock
            .lock()
            .map_err(|_| pdfium_unavailable())?;
        let document = self
            .pdfium
            .load_pdf_from_byte_vec(bytes, None)
            .map_err(map_load_error)?;
        let total_pages = page_count(document.pages().len())?;
        validate_page_count(total_pages, total_pages.max(1))?;
        let offset =
            usize::try_from(unit.ordinal.saturating_sub(1)).map_err(|_| pdf_page_not_found())?;
        if offset >= total_pages {
            return Err(pdf_page_not_found());
        }
        let index = i32::try_from(offset).map_err(|_| pdf_page_not_found())?;
        let page = document.pages().get(index).map_err(map_page_error)?;
        let observed = observe_page(&page, offset, dpi_milli)?;
        if &observed.unit != unit {
            return Err(layout_error(
                "use.office.pdf_page_identity_mismatch",
                "The selected PDF page no longer matches its inventoried identity and geometry.",
            ));
        }

        let width = i32::try_from(observed.output_width_px).map_err(|_| invalid_page_geometry())?;
        let height =
            i32::try_from(observed.output_height_px).map_err(|_| invalid_page_geometry())?;
        let raw_bytes = u64::from(observed.output_width_px)
            .checked_mul(u64::from(observed.output_height_px))
            .and_then(|pixels| pixels.checked_mul(4))
            .filter(|bytes| *bytes > 0 && *bytes <= MAX_PDF_BITMAP_BYTES)
            .ok_or_else(pdf_bitmap_too_large)?;

        let bitmap = page
            .render_with_config(&PdfRenderConfig::new().set_fixed_size(width, height))
            .map_err(map_page_error)?;
        if bitmap.width() != width || bitmap.height() != height {
            return Err(invalid_page_geometry());
        }
        let rgba = bitmap.as_rgba_bytes();
        if u64::try_from(rgba.len()).unwrap_or(u64::MAX) != raw_bytes {
            return Err(layout_error(
                "use.office.pdf_render_invalid",
                "PDFium returned an incomplete or inconsistent page bitmap.",
            ));
        }
        let png = encode_png(
            observed.output_width_px,
            observed.output_height_px,
            &rgba,
            max_output_bytes,
        )?;
        Ok((observed, png))
    }

    pub(super) fn extract_page_text(
        &self,
        bytes: Vec<u8>,
        inventory: &NativeOfficePdfPageInventory,
        unit: &NativeOfficeUnit,
        options: NativeOfficePdfTextLayerOptions,
    ) -> UseResult<NativeOfficePdfPageTextLayer> {
        inventory.validate()?;
        options.validate()?;
        let _guard = self
            .operation_lock
            .lock()
            .map_err(|_| pdfium_unavailable())?;
        let document = self
            .pdfium
            .load_pdf_from_byte_vec(bytes, None)
            .map_err(map_load_error)?;
        let total_pages = page_count(document.pages().len())?;
        if total_pages != inventory.total_pages {
            return Err(layout_error(
                "use.office.pdf_inventory_source_mismatch",
                "The PDF page inventory no longer matches the immutable source.",
            ));
        }
        let expected = inventory.validated_page(unit)?;
        let offset =
            usize::try_from(unit.ordinal.saturating_sub(1)).map_err(|_| pdf_page_not_found())?;
        let index = i32::try_from(offset).map_err(|_| pdf_page_not_found())?;
        let page = document.pages().get(index).map_err(map_page_error)?;
        let observed = observe_page(&page, offset, inventory.dpi_milli)?;
        if observed != *expected {
            return Err(layout_error(
                "use.office.pdf_page_identity_mismatch",
                "The selected PDF text page no longer matches its inventoried geometry.",
            ));
        }
        let layer = super::text::extract_page_text(
            &page,
            inventory.source_revision.clone(),
            observed,
            options,
        )?;
        layer.validate(inventory)?;
        Ok(layer)
    }

    pub(super) fn extract_outline(
        &self,
        bytes: Vec<u8>,
        inventory: &NativeOfficePdfPageInventory,
        options: NativeOfficePdfOutlineOptions,
    ) -> UseResult<NativeOfficePdfOutline> {
        inventory.validate()?;
        options.validate()?;
        let _guard = self
            .operation_lock
            .lock()
            .map_err(|_| pdfium_unavailable())?;
        let document = self
            .pdfium
            .load_pdf_from_byte_vec(bytes, None)
            .map_err(map_load_error)?;
        let total_pages = page_count(document.pages().len())?;
        if total_pages != inventory.total_pages {
            return Err(layout_error(
                "use.office.pdf_inventory_source_mismatch",
                "The PDF page inventory no longer matches the immutable source.",
            ));
        }
        super::outline::extract_outline(&document, inventory, options)
    }
}

async fn read_library(path: &Path) -> UseResult<Vec<u8>> {
    let metadata = tokio::fs::symlink_metadata(path)
        .await
        .map_err(|_| pdfium_library_invalid())?;
    if !metadata.is_file()
        || metadata.file_type().is_symlink()
        || metadata.len() == 0
        || metadata.len() > MAX_PDFIUM_LIBRARY_BYTES
    {
        return Err(pdfium_library_invalid());
    }
    let bytes = tokio::fs::read(path)
        .await
        .map_err(|_| pdfium_library_invalid())?;
    if u64::try_from(bytes.len()).unwrap_or(u64::MAX) != metadata.len() {
        return Err(pdfium_library_invalid());
    }
    Ok(bytes)
}

fn initialize_engine(bytes: Vec<u8>, binary_sha256: String) -> UseResult<Arc<PdfiumEngine>> {
    let _guard = PDFIUM_INIT_LOCK
        .lock()
        .map_err(|_| pdfium_library_invalid())?;
    if let Some(engine) = PDFIUM_ENGINE.get() {
        return engine_for_identity(engine, &binary_sha256);
    }

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
    file.write_all(&bytes)
        .and_then(|()| file.sync_all())
        .map_err(|_| pdfium_library_invalid())?;
    drop(file);
    let staged_bytes = std::fs::read(&staged_path).map_err(|_| pdfium_library_invalid())?;
    if format!("{:x}", Sha256::digest(&staged_bytes)) != binary_sha256 {
        return Err(pdfium_library_invalid());
    }

    let bindings = Pdfium::bind_to_library(&staged_path).map_err(|_| pdfium_library_invalid())?;
    let pdfium = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| Pdfium::new(bindings)))
        .map_err(|_| pdfium_library_invalid())?;
    let engine = Arc::new(PdfiumEngine {
        pdfium,
        binary_sha256,
        operation_lock: Mutex::new(()),
        _staged_library: staged_library,
    });
    PDFIUM_ENGINE
        .set(Arc::clone(&engine))
        .map_err(|_| pdfium_library_invalid())?;
    Ok(engine)
}

fn engine_for_identity(
    engine: &Arc<PdfiumEngine>,
    binary_sha256: &str,
) -> UseResult<Arc<PdfiumEngine>> {
    if engine.binary_sha256 == binary_sha256 {
        return Ok(Arc::clone(engine));
    }
    Err(layout_error(
        "use.office.pdfium_library_conflict",
        "A different PDFium binary is already bound in this process.",
    ))
}

fn observe_page(
    page: &pdfium_render::prelude::PdfPage<'_>,
    offset: usize,
    dpi_milli: u32,
) -> UseResult<NativeOfficePdfPageGeometry> {
    let number = u32::try_from(offset + 1).map_err(|_| invalid_page_geometry())?;
    let media = page_box(page.boundaries().media().map_err(map_page_error)?)?;
    let crop = match page.boundaries().crop() {
        Ok(boundary) => page_box(boundary)?,
        Err(_) => media,
    };
    let rotation_degrees = rotation_degrees(page.rotation().map_err(map_page_error)?);
    let (width_millipoints, height_millipoints) = if matches!(rotation_degrees, 90 | 270) {
        (crop.height_millipoints(), crop.width_millipoints())
    } else {
        (crop.width_millipoints(), crop.height_millipoints())
    };
    let observed_width = points_to_millipoints(page.width().value)?;
    let observed_height = points_to_millipoints(page.height().value)?;
    if observed_width.abs_diff(width_millipoints) > 2
        || observed_height.abs_diff(height_millipoints) > 2
    {
        return Err(invalid_page_geometry());
    }
    let geometry = NativeOfficePdfPageGeometry {
        unit: NativeOfficeUnit {
            ordinal: number,
            locator: NativeOfficeUnitLocator::Page { number },
            path: format!("/page[{number}]"),
        },
        media_box: media,
        crop_box: crop,
        rotation_degrees,
        surface_width_micrometers: millipoints_to_micrometers(width_millipoints)?,
        surface_height_micrometers: millipoints_to_micrometers(height_millipoints)?,
        output_width_px: millipoints_to_pixels(width_millipoints, dpi_milli)?,
        output_height_px: millipoints_to_pixels(height_millipoints, dpi_milli)?,
    };
    geometry.validate(dpi_milli)?;
    Ok(geometry)
}

fn page_box(boundary: PdfPageBoundaryBox) -> UseResult<NativeOfficePdfPageBox> {
    Ok(NativeOfficePdfPageBox {
        left_millipoints: signed_points_to_millipoints(boundary.bounds.left().value)?,
        bottom_millipoints: signed_points_to_millipoints(boundary.bounds.bottom().value)?,
        right_millipoints: signed_points_to_millipoints(boundary.bounds.right().value)?,
        top_millipoints: signed_points_to_millipoints(boundary.bounds.top().value)?,
    })
}

pub(super) fn signed_points_to_millipoints(points: f32) -> UseResult<i64> {
    let value = f64::from(points) * 1_000.0;
    if !value.is_finite() || value < i64::MIN as f64 || value > i64::MAX as f64 {
        return Err(invalid_page_geometry());
    }
    Ok(value.round() as i64)
}

fn points_to_millipoints(points: f32) -> UseResult<u64> {
    u64::try_from(signed_points_to_millipoints(points)?)
        .ok()
        .filter(|value| *value > 0)
        .ok_or_else(invalid_page_geometry)
}

fn rotation_degrees(rotation: PdfPageRenderRotation) -> u16 {
    match rotation {
        PdfPageRenderRotation::None => 0,
        PdfPageRenderRotation::Degrees90 => 90,
        PdfPageRenderRotation::Degrees180 => 180,
        PdfPageRenderRotation::Degrees270 => 270,
    }
}

fn encode_png(width: u32, height: u32, rgba: &[u8], max_bytes: u64) -> UseResult<Vec<u8>> {
    let mut bytes = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut bytes, width, height);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        let mut writer = encoder.write_header().map_err(|_| pdf_render_invalid())?;
        writer
            .write_image_data(rgba)
            .map_err(|_| pdf_render_invalid())?;
    }
    let byte_length = u64::try_from(bytes.len()).unwrap_or(u64::MAX);
    if byte_length == 0 || byte_length > max_bytes {
        return Err(layout_error(
            "use.office.layout_output_too_large",
            format!("Office layout raster exceeds the {max_bytes}-byte output limit."),
        ));
    }
    Ok(bytes)
}

fn validate_page_count(total_pages: usize, max_pages: usize) -> UseResult<()> {
    if total_pages == 0 {
        return Err(layout_error(
            "use.office.pdf_zero_pages",
            "PDFium found no pages in the PDF document.",
        ));
    }
    if total_pages > max_pages {
        return Err(pdf_page_limit(total_pages, max_pages));
    }
    Ok(())
}

fn page_count(value: i32) -> UseResult<usize> {
    usize::try_from(value).map_err(|_| {
        layout_error(
            "use.office.pdf_corrupt",
            "PDFium returned an invalid page count for the PDF document.",
        )
    })
}

fn map_load_error(error: PdfiumError) -> a3s_use_core::UseError {
    match error {
        PdfiumError::PdfiumLibraryInternalError(PdfiumInternalError::PasswordError)
        | PdfiumError::PdfiumLibraryInternalError(PdfiumInternalError::SecurityError) => {
            layout_error(
                "use.office.pdf_password_required",
                "The PDF is encrypted or requires a password.",
            )
        }
        PdfiumError::PdfiumLibraryInternalError(PdfiumInternalError::FormatError)
        | PdfiumError::PdfiumLibraryInternalError(PdfiumInternalError::FileError) => layout_error(
            "use.office.pdf_corrupt",
            "PDFium could not parse the PDF document.",
        ),
        _ => map_page_error(error),
    }
}

fn map_page_error(_error: PdfiumError) -> a3s_use_core::UseError {
    layout_error(
        "use.office.pdf_unsupported",
        "PDFium could not inspect or render the requested PDF feature.",
    )
}

fn pdfium_library_invalid() -> a3s_use_core::UseError {
    layout_error(
        "use.office.pdfium_library_invalid",
        "The explicit PDFium library is missing, unsafe, oversized, or incompatible with build 7881.",
    )
}

fn pdfium_unavailable() -> a3s_use_core::UseError {
    layout_error(
        "use.office.pdfium_unavailable",
        "The process-wide PDFium engine is unavailable.",
    )
}

fn pdf_page_limit(total_pages: usize, max_pages: usize) -> a3s_use_core::UseError {
    layout_error(
        "use.office.pdf_page_limit",
        format!("PDF contains {total_pages} pages; the configured limit is {max_pages}."),
    )
}

fn pdf_page_not_found() -> a3s_use_core::UseError {
    layout_error(
        "use.office.pdf_page_not_found",
        "The requested one-based PDF page does not exist.",
    )
}

fn pdf_bitmap_too_large() -> a3s_use_core::UseError {
    layout_error(
        "use.office.pdf_bitmap_too_large",
        "The requested PDF page bitmap exceeds the in-memory rendering limit.",
    )
}

fn pdf_render_invalid() -> a3s_use_core::UseError {
    layout_error(
        "use.office.pdf_render_invalid",
        "The PDF page bitmap could not be encoded as a deterministic PNG.",
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pdfium_load_failures_have_stable_typed_codes() {
        let password = map_load_error(PdfiumError::PdfiumLibraryInternalError(
            PdfiumInternalError::PasswordError,
        ));
        assert_eq!(password.code, "use.office.pdf_password_required");

        let security = map_load_error(PdfiumError::PdfiumLibraryInternalError(
            PdfiumInternalError::SecurityError,
        ));
        assert_eq!(security.code, "use.office.pdf_password_required");

        let corrupt = map_load_error(PdfiumError::PdfiumLibraryInternalError(
            PdfiumInternalError::FormatError,
        ));
        assert_eq!(corrupt.code, "use.office.pdf_corrupt");
    }
}
