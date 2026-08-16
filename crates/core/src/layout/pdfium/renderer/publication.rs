use a3s_use_core::UseResult;
use sha2::{Digest, Sha256};

use super::NativeOfficePdfiumLayoutRenderer;
use crate::layout::pdfium::NativeOfficePdfPageGeometry;
use crate::layout::pptx_image::io::{
    publish_output, stage_output, validate_published_output, verify_source_revision, StagedOutput,
};
use crate::layout::{
    layout_error, NativeOfficeLayoutEnvironment, NativeOfficeLayoutReceipt,
    NativeOfficeLayoutRenderRequest,
};

pub(super) struct StagedRenderedPage {
    output: StagedOutput,
    png_sha256: String,
    profile_sha256: String,
}

impl NativeOfficePdfiumLayoutRenderer {
    pub(super) async fn stage_rendered_page(
        &self,
        request: &NativeOfficeLayoutRenderRequest,
        page: &NativeOfficePdfPageGeometry,
        png: Vec<u8>,
    ) -> UseResult<StagedRenderedPage> {
        let expected_profile = self.profile(
            page,
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
        let profile_sha256 = request.profile.sha256()?;
        let output = stage_output(&request.output, png).await?;
        Ok(StagedRenderedPage {
            output,
            png_sha256,
            profile_sha256,
        })
    }

    pub(super) async fn publish_staged_rendered_page(
        &self,
        request: NativeOfficeLayoutRenderRequest,
        page: NativeOfficePdfPageGeometry,
        staged: StagedRenderedPage,
    ) -> UseResult<NativeOfficeLayoutReceipt> {
        publish_output(staged.output, &request.output).await?;
        let raster = validate_published_output(
            &request.output,
            request.max_output_bytes,
            page.output_width_px,
            page.output_height_px,
            &staged.png_sha256,
            page.rotation_degrees,
        )
        .await?;
        let receipt = NativeOfficeLayoutReceipt {
            source_revision: request.source_revision.clone(),
            render_input_sha256: request.source_revision.sha256,
            unit: request.unit,
            profile: request.profile,
            profile_sha256: staged.profile_sha256,
            raster,
        };
        receipt.validate()?;
        Ok(receipt)
    }

    pub(super) async fn publish_rendered_page_with_hook<F>(
        &self,
        request: NativeOfficeLayoutRenderRequest,
        page: NativeOfficePdfPageGeometry,
        png: Vec<u8>,
        before_final_source_check: F,
    ) -> UseResult<NativeOfficeLayoutReceipt>
    where
        F: FnOnce() + Send,
    {
        let staged = self.stage_rendered_page(&request, &page, png).await?;
        before_final_source_check();
        verify_source_revision(&request.source_path, &request.source_revision).await?;
        let source_path = request.source_path.clone();
        let source_revision = request.source_revision.clone();
        let receipt = self
            .publish_staged_rendered_page(request, page, staged)
            .await?;
        verify_source_revision(&source_path, &source_revision).await?;
        Ok(receipt)
    }
}
