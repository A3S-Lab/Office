use std::sync::Arc;
use std::time::Duration;

use a3s_use_core::UseResult;

use super::{
    pdfium_task_failed, read_source_bytes, validate_request, verify_source_revision,
    NativeOfficePdfiumLayoutRenderer, MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES,
};
use crate::layout::pdfium::engine::render::PdfPageRenderInput;
use crate::layout::pdfium::render_batch::validate_render_batch_request;
use crate::layout::pdfium::{
    NativeOfficePdfPageRenderBatch, NativeOfficePdfRenderBatchOptions,
    NativeOfficePdfRenderBatchSlot, NativeOfficePdfRenderBatchSlotOutcome,
};
use crate::layout::NativeOfficeLayoutRenderRequest;

impl NativeOfficePdfiumLayoutRenderer {
    /// Renders one bounded ordered page batch after one immutable source read
    /// and one PDFium document open. Page rendering and publication failures
    /// remain isolated in their exact input slots.
    pub async fn render_batch(
        &self,
        requests: Vec<NativeOfficeLayoutRenderRequest>,
        options: NativeOfficePdfRenderBatchOptions,
    ) -> UseResult<NativeOfficePdfPageRenderBatch> {
        self.render_batch_with_hook(requests, options, || {}).await
    }

    async fn render_batch_with_hook<F>(
        &self,
        requests: Vec<NativeOfficeLayoutRenderRequest>,
        options: NativeOfficePdfRenderBatchOptions,
        before_publish_source_check: F,
    ) -> UseResult<NativeOfficePdfPageRenderBatch>
    where
        F: FnOnce() + Send,
    {
        validate_render_batch_request(&requests, options)?;
        for request in &requests {
            validate_request(self, request).await?;
        }
        let timeout_ms = options.timeout_ms;
        match tokio::time::timeout(
            Duration::from_millis(timeout_ms),
            self.render_batch_inner(requests, options, before_publish_source_check),
        )
        .await
        {
            Ok(result) => result,
            Err(_) => Err(super::layout_timeout(timeout_ms)),
        }
    }

    async fn render_batch_inner(
        &self,
        requests: Vec<NativeOfficeLayoutRenderRequest>,
        options: NativeOfficePdfRenderBatchOptions,
        before_publish_source_check: impl FnOnce() + Send,
    ) -> UseResult<NativeOfficePdfPageRenderBatch> {
        let first = &requests[0];
        let bytes = read_source_bytes(
            &first.source_path,
            &first.source_revision,
            MAX_NATIVE_OFFICE_PDF_SOURCE_BYTES,
        )
        .await?;
        let inputs = requests
            .iter()
            .map(|request| PdfPageRenderInput {
                unit: request.unit.clone(),
                dpi_milli: request.profile.dpi_x_milli,
                max_output_bytes: request.max_output_bytes,
            })
            .collect::<Vec<_>>();
        let engine = Arc::clone(&self.engine);
        let rendered =
            tokio::task::spawn_blocking(move || engine.render_page_batch(bytes, &inputs))
                .await
                .map_err(|_| pdfium_task_failed())??;

        let mut outcomes = (0..requests.len()).map(|_| None).collect::<Vec<_>>();
        let mut staged_pages = Vec::with_capacity(requests.len());
        for (index, rendered) in rendered.into_iter().enumerate() {
            match rendered {
                Ok((page, png)) => {
                    match self.stage_rendered_page(&requests[index], &page, png).await {
                        Ok(staged) => staged_pages.push((index, page, staged)),
                        Err(error) => {
                            outcomes[index] =
                                Some(NativeOfficePdfRenderBatchSlotOutcome::Failed { error });
                        }
                    }
                }
                Err(error) => {
                    outcomes[index] = Some(NativeOfficePdfRenderBatchSlotOutcome::Failed { error });
                }
            }
        }

        before_publish_source_check();
        verify_source_revision(&first.source_path, &first.source_revision).await?;
        for (index, page, staged) in staged_pages {
            let outcome = match self
                .publish_staged_rendered_page(requests[index].clone(), page, staged)
                .await
            {
                Ok(receipt) => NativeOfficePdfRenderBatchSlotOutcome::Completed {
                    receipt: Box::new(receipt),
                },
                Err(error) => NativeOfficePdfRenderBatchSlotOutcome::Failed { error },
            };
            outcomes[index] = Some(outcome);
        }
        verify_source_revision(&first.source_path, &first.source_revision).await?;

        let mut slots = Vec::with_capacity(requests.len());
        for (request, outcome) in requests.iter().zip(outcomes) {
            let outcome = outcome.ok_or_else(pdfium_task_failed)?;
            slots.push(NativeOfficePdfRenderBatchSlot {
                unit: request.unit.clone(),
                outcome,
            });
        }
        let batch = NativeOfficePdfPageRenderBatch {
            source_revision: requests[0].source_revision.clone(),
            options,
            slots,
        };
        batch.validate_for(&requests, options)?;
        Ok(batch)
    }

    #[cfg(test)]
    pub(crate) async fn render_batch_with_before_publish_source_check<F>(
        &self,
        requests: Vec<NativeOfficeLayoutRenderRequest>,
        options: NativeOfficePdfRenderBatchOptions,
        before_publish_source_check: F,
    ) -> UseResult<NativeOfficePdfPageRenderBatch>
    where
        F: FnOnce() + Send,
    {
        self.render_batch_with_hook(requests, options, before_publish_source_check)
            .await
    }
}
