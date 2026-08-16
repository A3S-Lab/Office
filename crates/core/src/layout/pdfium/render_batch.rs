use std::collections::BTreeSet;

use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};

use crate::layout::{
    layout_error, validate_revision, NativeOfficeLayoutReceipt, NativeOfficeLayoutRenderRequest,
};
use crate::{NativeOfficeUnit, PackageRevision};

/// Default maximum number of ordered pages accepted by one render batch.
pub const DEFAULT_NATIVE_OFFICE_PDF_RENDER_BATCH_PAGES: usize = 16;
/// Hard maximum number of ordered pages accepted by one render batch.
pub const MAX_NATIVE_OFFICE_PDF_RENDER_BATCH_PAGES: usize = 64;
/// Default maximum sum of successful PNG bytes retained by one render batch.
pub const DEFAULT_NATIVE_OFFICE_PDF_RENDER_BATCH_BYTES: u64 = 512 * 1024 * 1024;
/// Hard maximum sum of successful PNG bytes retained by one render batch.
pub const MAX_NATIVE_OFFICE_PDF_RENDER_BATCH_BYTES: u64 = 2 * 1024 * 1024 * 1024;

/// Explicit cardinality, aggregate output, and wall-time bounds for one
/// ordered native PDF render batch.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfRenderBatchOptions {
    pub max_pages: usize,
    pub max_total_output_bytes: u64,
    pub timeout_ms: u64,
}

impl Default for NativeOfficePdfRenderBatchOptions {
    fn default() -> Self {
        Self {
            max_pages: DEFAULT_NATIVE_OFFICE_PDF_RENDER_BATCH_PAGES,
            max_total_output_bytes: DEFAULT_NATIVE_OFFICE_PDF_RENDER_BATCH_BYTES,
            timeout_ms: super::super::MAX_LAYOUT_TIMEOUT_MS,
        }
    }
}

impl NativeOfficePdfRenderBatchOptions {
    pub fn validate(&self) -> UseResult<()> {
        if (1..=MAX_NATIVE_OFFICE_PDF_RENDER_BATCH_PAGES).contains(&self.max_pages)
            && (1..=MAX_NATIVE_OFFICE_PDF_RENDER_BATCH_BYTES).contains(&self.max_total_output_bytes)
            && (1..=super::super::MAX_LAYOUT_TIMEOUT_MS).contains(&self.timeout_ms)
        {
            return Ok(());
        }
        Err(layout_error(
            "use.office.pdf_render_batch_options_invalid",
            "PDF render batch page, aggregate output, and timeout bounds are invalid.",
        ))
    }
}

/// Isolated result for one exact page in a native PDF render batch.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "status",
    rename_all = "kebab-case",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum NativeOfficePdfRenderBatchSlotOutcome {
    Completed {
        receipt: Box<NativeOfficeLayoutReceipt>,
    },
    Failed {
        error: UseError,
    },
}

/// One exact ordered page identity and its isolated render outcome.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfRenderBatchSlot {
    pub unit: NativeOfficeUnit,
    pub outcome: NativeOfficePdfRenderBatchSlotOutcome,
}

/// Ordered page rasters produced from one immutable PDF source read and one
/// PDFium document open.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfPageRenderBatch {
    pub source_revision: PackageRevision,
    pub options: NativeOfficePdfRenderBatchOptions,
    pub slots: Vec<NativeOfficePdfRenderBatchSlot>,
}

impl NativeOfficePdfPageRenderBatch {
    pub fn validate_for(
        &self,
        requests: &[NativeOfficeLayoutRenderRequest],
        options: NativeOfficePdfRenderBatchOptions,
    ) -> UseResult<()> {
        validate_render_batch_request(requests, options)?;
        if self.source_revision != requests[0].source_revision
            || self.options != options
            || self.slots.len() != requests.len()
        {
            return Err(invalid_render_batch_output());
        }
        let mut total_output_bytes = 0_u64;
        for (slot, request) in self.slots.iter().zip(requests) {
            if slot.unit != request.unit {
                return Err(invalid_render_batch_output());
            }
            match &slot.outcome {
                NativeOfficePdfRenderBatchSlotOutcome::Completed { receipt } => {
                    receipt.validate()?;
                    if receipt.source_revision != request.source_revision
                        || receipt.unit != request.unit
                        || receipt.profile != request.profile
                        || receipt.raster.output_path != request.output
                        || receipt.raster.byte_length > request.max_output_bytes
                    {
                        return Err(invalid_render_batch_output());
                    }
                    total_output_bytes = total_output_bytes
                        .checked_add(receipt.raster.byte_length)
                        .ok_or_else(invalid_render_batch_output)?;
                }
                NativeOfficePdfRenderBatchSlotOutcome::Failed { error } => {
                    if error.code.trim().is_empty()
                        || error.code.len() > 256
                        || error.message.trim().is_empty()
                        || error.message.len() > 4 * 1024
                        || error.suggestion.is_some()
                        || !error.details.is_empty()
                    {
                        return Err(invalid_render_batch_output());
                    }
                }
            }
        }
        if total_output_bytes > options.max_total_output_bytes {
            return Err(invalid_render_batch_output());
        }
        Ok(())
    }
}

pub(super) fn validate_render_batch_request(
    requests: &[NativeOfficeLayoutRenderRequest],
    options: NativeOfficePdfRenderBatchOptions,
) -> UseResult<()> {
    options.validate()?;
    if requests.is_empty() || requests.len() > options.max_pages {
        return Err(invalid_render_batch_request());
    }
    let source_path = &requests[0].source_path;
    let source_revision = &requests[0].source_revision;
    validate_revision(source_revision)?;
    let mut units = BTreeSet::new();
    let mut outputs = BTreeSet::new();
    let mut reserved_output_bytes = 0_u64;
    for request in requests {
        reserved_output_bytes = reserved_output_bytes
            .checked_add(request.max_output_bytes)
            .ok_or_else(invalid_render_batch_request)?;
        if &request.source_path != source_path
            || &request.source_revision != source_revision
            || request.timeout_ms > options.timeout_ms
            || !units.insert(&request.unit)
            || !outputs.insert(&request.output)
        {
            return Err(invalid_render_batch_request());
        }
    }
    if reserved_output_bytes > options.max_total_output_bytes {
        return Err(invalid_render_batch_request());
    }
    Ok(())
}

fn invalid_render_batch_request() -> UseError {
    layout_error(
        "use.office.pdf_render_batch_invalid",
        "A PDF render batch must contain bounded, unique pages and outputs from one immutable source.",
    )
}

fn invalid_render_batch_output() -> UseError {
    layout_error(
        "use.office.pdf_render_batch_output_invalid",
        "A PDF render batch changed source, options, order, page identity, output, or aggregate bounds.",
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn render_batch_options_have_a_strict_bounded_wire_contract() {
        let options = NativeOfficePdfRenderBatchOptions::default();
        options.validate().unwrap();
        let value = serde_json::to_value(options).unwrap();
        assert_eq!(
            value["maxPages"],
            DEFAULT_NATIVE_OFFICE_PDF_RENDER_BATCH_PAGES
        );
        assert_eq!(
            serde_json::from_value::<NativeOfficePdfRenderBatchOptions>(value.clone()).unwrap(),
            options
        );
        let mut unknown = value;
        unknown["unbounded"] = serde_json::json!(true);
        assert!(serde_json::from_value::<NativeOfficePdfRenderBatchOptions>(unknown).is_err());
    }

    #[test]
    fn public_render_batch_contracts_are_send_and_sync() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<NativeOfficePdfRenderBatchOptions>();
        assert_send_sync::<NativeOfficePdfRenderBatchSlotOutcome>();
        assert_send_sync::<NativeOfficePdfRenderBatchSlot>();
        assert_send_sync::<NativeOfficePdfPageRenderBatch>();
    }
}
