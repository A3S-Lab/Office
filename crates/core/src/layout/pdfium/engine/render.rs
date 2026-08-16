use a3s_use_core::UseResult;
use pdfium_render::prelude::{PdfDocument, PdfPage, PdfRenderConfig};

use super::{
    invalid_page_geometry, map_load_error, map_page_error, observe_page, page_count,
    pdf_bitmap_too_large, pdf_page_not_found, validate_page_count, PdfiumEngine,
};
use crate::layout::layout_error;
use crate::{NativeOfficePdfPageGeometry, NativeOfficeUnit};

const MAX_PDF_BITMAP_BYTES: u64 = 256 * 1024 * 1024;

pub(in crate::layout::pdfium) struct PdfPageRenderInput {
    pub unit: NativeOfficeUnit,
    pub dpi_milli: u32,
    pub max_output_bytes: u64,
}

pub(in crate::layout::pdfium) type PdfPageRenderResult =
    UseResult<(NativeOfficePdfPageGeometry, Vec<u8>)>;

impl PdfiumEngine {
    pub(in crate::layout::pdfium) fn render_page(
        &self,
        bytes: Vec<u8>,
        unit: &NativeOfficeUnit,
        dpi_milli: u32,
        max_output_bytes: u64,
    ) -> UseResult<(NativeOfficePdfPageGeometry, Vec<u8>)> {
        let _guard = self
            .operation_lock
            .lock()
            .map_err(|_| super::pdfium_unavailable())?;
        let document = self
            .pdfium
            .load_pdf_from_byte_vec(bytes, None)
            .map_err(map_load_error)?;
        render_document_page(&document, unit, dpi_milli, max_output_bytes)
    }

    pub(in crate::layout::pdfium) fn render_page_batch(
        &self,
        bytes: Vec<u8>,
        inputs: &[PdfPageRenderInput],
    ) -> UseResult<Vec<PdfPageRenderResult>> {
        let _guard = self
            .operation_lock
            .lock()
            .map_err(|_| super::pdfium_unavailable())?;
        let document = self
            .pdfium
            .load_pdf_from_byte_vec(bytes, None)
            .map_err(map_load_error)?;
        let total_pages = page_count(document.pages().len())?;
        validate_page_count(total_pages, total_pages.max(1))?;
        Ok(inputs
            .iter()
            .map(|input| {
                render_document_page(
                    &document,
                    &input.unit,
                    input.dpi_milli,
                    input.max_output_bytes,
                )
            })
            .collect())
    }
}

fn render_document_page(
    document: &PdfDocument<'_>,
    unit: &NativeOfficeUnit,
    dpi_milli: u32,
    max_output_bytes: u64,
) -> UseResult<(NativeOfficePdfPageGeometry, Vec<u8>)> {
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
    let png = render_png(&page, &observed, max_output_bytes)?;
    Ok((observed, png))
}

fn render_png(
    page: &PdfPage<'_>,
    observed: &NativeOfficePdfPageGeometry,
    max_output_bytes: u64,
) -> UseResult<Vec<u8>> {
    let width = i32::try_from(observed.output_width_px).map_err(|_| invalid_page_geometry())?;
    let height = i32::try_from(observed.output_height_px).map_err(|_| invalid_page_geometry())?;
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
    super::encode_png(
        observed.output_width_px,
        observed.output_height_px,
        &rgba,
        max_output_bytes,
    )
}
