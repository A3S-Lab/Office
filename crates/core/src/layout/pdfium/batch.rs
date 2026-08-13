use std::collections::BTreeSet;

use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};

use super::{
    NativeOfficePdfPageInventory, NativeOfficePdfPageTextLayer, NativeOfficePdfTextLayerOptions,
    MAX_NATIVE_OFFICE_PDF_TEXT_CHARACTERS, MAX_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES,
};
use crate::layout::layout_error;
use crate::{NativeOfficeUnit, PackageRevision};

/// Default maximum number of ordered pages accepted by one text batch.
pub const DEFAULT_NATIVE_OFFICE_PDF_TEXT_BATCH_PAGES: usize = 64;
/// Hard maximum number of ordered pages accepted by one text batch.
pub const MAX_NATIVE_OFFICE_PDF_TEXT_BATCH_PAGES: usize = 512;
/// Default maximum number of successful characters retained by one batch.
pub const DEFAULT_NATIVE_OFFICE_PDF_TEXT_BATCH_CHARACTERS: usize = 5_000_000;
/// Hard maximum number of successful characters retained by one batch.
pub const MAX_NATIVE_OFFICE_PDF_TEXT_BATCH_CHARACTERS: usize = 20_000_000;
/// Default maximum number of successful UTF-8 text bytes retained by one batch.
pub const DEFAULT_NATIVE_OFFICE_PDF_TEXT_BATCH_TEXT_BYTES: usize = 64 * 1024 * 1024;
/// Hard maximum number of successful UTF-8 text bytes retained by one batch.
pub const MAX_NATIVE_OFFICE_PDF_TEXT_BATCH_TEXT_BYTES: usize = 256 * 1024 * 1024;

/// Explicit per-page, aggregate, cardinality, and deadline bounds for one
/// ordered native PDF text extraction batch.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfTextBatchOptions {
    pub max_pages: usize,
    pub max_characters_per_page: usize,
    pub max_text_bytes_per_page: usize,
    pub max_total_characters: usize,
    pub max_total_text_bytes: usize,
    pub timeout_ms: u64,
}

impl Default for NativeOfficePdfTextBatchOptions {
    fn default() -> Self {
        let page = NativeOfficePdfTextLayerOptions::default();
        Self {
            max_pages: DEFAULT_NATIVE_OFFICE_PDF_TEXT_BATCH_PAGES,
            max_characters_per_page: page.max_characters,
            max_text_bytes_per_page: page.max_text_bytes,
            max_total_characters: DEFAULT_NATIVE_OFFICE_PDF_TEXT_BATCH_CHARACTERS,
            max_total_text_bytes: DEFAULT_NATIVE_OFFICE_PDF_TEXT_BATCH_TEXT_BYTES,
            timeout_ms: page.timeout_ms,
        }
    }
}

impl NativeOfficePdfTextBatchOptions {
    pub fn validate(&self) -> UseResult<()> {
        if (1..=MAX_NATIVE_OFFICE_PDF_TEXT_BATCH_PAGES).contains(&self.max_pages)
            && (1..=MAX_NATIVE_OFFICE_PDF_TEXT_CHARACTERS).contains(&self.max_characters_per_page)
            && (1..=MAX_NATIVE_OFFICE_PDF_TEXT_PAGE_BYTES).contains(&self.max_text_bytes_per_page)
            && (1..=MAX_NATIVE_OFFICE_PDF_TEXT_BATCH_CHARACTERS)
                .contains(&self.max_total_characters)
            && (1..=MAX_NATIVE_OFFICE_PDF_TEXT_BATCH_TEXT_BYTES)
                .contains(&self.max_total_text_bytes)
            && (1..=super::super::MAX_LAYOUT_TIMEOUT_MS).contains(&self.timeout_ms)
        {
            return Ok(());
        }
        Err(layout_error(
            "use.office.pdf_text_batch_options_invalid",
            "PDF text batch page, character, byte, and timeout bounds are invalid.",
        ))
    }

    pub(super) fn page_options(self) -> NativeOfficePdfTextLayerOptions {
        NativeOfficePdfTextLayerOptions {
            max_characters: self.max_characters_per_page,
            max_text_bytes: self.max_text_bytes_per_page,
            timeout_ms: self.timeout_ms,
        }
    }
}

/// Isolated result for one exact page in a native PDF text batch.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "status",
    rename_all = "kebab-case",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum NativeOfficePdfTextBatchSlotOutcome {
    Completed {
        layer: Box<NativeOfficePdfPageTextLayer>,
    },
    Failed {
        error: UseError,
    },
}

/// One exact ordered page identity and its isolated text extraction outcome.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfTextBatchSlot {
    pub unit: NativeOfficeUnit,
    pub outcome: NativeOfficePdfTextBatchSlotOutcome,
}

/// Ordered native PDF text layers produced from one immutable source read and
/// one PDFium document open.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfPageTextBatch {
    pub source_revision: PackageRevision,
    pub options: NativeOfficePdfTextBatchOptions,
    pub slots: Vec<NativeOfficePdfTextBatchSlot>,
}

impl NativeOfficePdfPageTextBatch {
    pub fn validate_for(
        &self,
        inventory: &NativeOfficePdfPageInventory,
        requested_units: &[NativeOfficeUnit],
        options: NativeOfficePdfTextBatchOptions,
    ) -> UseResult<()> {
        validate_batch_request(inventory, requested_units, options)?;
        if self.source_revision != inventory.source_revision
            || self.options != options
            || self.slots.len() != requested_units.len()
        {
            return Err(invalid_batch_output());
        }
        let mut total_characters = 0_usize;
        let mut total_text_bytes = 0_usize;
        for (slot, requested) in self.slots.iter().zip(requested_units) {
            if &slot.unit != requested {
                return Err(invalid_batch_output());
            }
            match &slot.outcome {
                NativeOfficePdfTextBatchSlotOutcome::Completed { layer } => {
                    layer.validate(inventory)?;
                    if layer.unit != slot.unit
                        || layer.max_characters != options.max_characters_per_page
                        || layer.max_text_bytes != options.max_text_bytes_per_page
                    {
                        return Err(invalid_batch_output());
                    }
                    total_characters = total_characters
                        .checked_add(layer.characters.len())
                        .ok_or_else(invalid_batch_output)?;
                    total_text_bytes = total_text_bytes
                        .checked_add(layer.text.len())
                        .ok_or_else(invalid_batch_output)?;
                }
                NativeOfficePdfTextBatchSlotOutcome::Failed { error } => {
                    if error.code.trim().is_empty()
                        || error.code.len() > 256
                        || error.message.trim().is_empty()
                        || error.message.len() > 4 * 1024
                        || error.suggestion.is_some()
                        || !error.details.is_empty()
                    {
                        return Err(invalid_batch_output());
                    }
                }
            }
        }
        if total_characters > options.max_total_characters
            || total_text_bytes > options.max_total_text_bytes
        {
            return Err(invalid_batch_output());
        }
        Ok(())
    }
}

pub(super) fn validate_batch_request(
    inventory: &NativeOfficePdfPageInventory,
    units: &[NativeOfficeUnit],
    options: NativeOfficePdfTextBatchOptions,
) -> UseResult<()> {
    inventory.validate()?;
    options.validate()?;
    if units.is_empty() || units.len() > options.max_pages {
        return Err(layout_error(
            "use.office.pdf_text_batch_invalid",
            "A PDF text batch must contain one bounded, unique, inventoried page sequence.",
        ));
    }
    let mut unique = BTreeSet::new();
    for unit in units {
        inventory.validated_page(unit)?;
        if !unique.insert(unit) {
            return Err(layout_error(
                "use.office.pdf_text_batch_invalid",
                "A PDF text batch must contain one bounded, unique, inventoried page sequence.",
            ));
        }
    }
    Ok(())
}

pub(super) fn total_character_limit() -> UseError {
    layout_error(
        "use.office.pdf_text_batch_character_limit",
        "The successful PDF text batch characters exceed the configured aggregate limit.",
    )
}

pub(super) fn total_text_byte_limit() -> UseError {
    layout_error(
        "use.office.pdf_text_batch_byte_limit",
        "The successful PDF text batch bytes exceed the configured aggregate limit.",
    )
}

fn invalid_batch_output() -> UseError {
    layout_error(
        "use.office.pdf_text_batch_output_invalid",
        "A PDF text batch changed source, options, order, page identity, or aggregate bounds.",
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn batch_options_have_a_strict_bounded_wire_contract() {
        let options = NativeOfficePdfTextBatchOptions::default();
        options.validate().unwrap();
        let value = serde_json::to_value(options).unwrap();
        assert_eq!(
            value["maxPages"],
            DEFAULT_NATIVE_OFFICE_PDF_TEXT_BATCH_PAGES
        );
        assert_eq!(
            serde_json::from_value::<NativeOfficePdfTextBatchOptions>(value.clone()).unwrap(),
            options
        );
        let mut unknown = value;
        unknown["unbounded"] = serde_json::json!(true);
        assert!(serde_json::from_value::<NativeOfficePdfTextBatchOptions>(unknown).is_err());
    }

    #[test]
    fn public_batch_contracts_are_send_and_sync() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<NativeOfficePdfTextBatchOptions>();
        assert_send_sync::<NativeOfficePdfTextBatchSlotOutcome>();
        assert_send_sync::<NativeOfficePdfTextBatchSlot>();
        assert_send_sync::<NativeOfficePdfPageTextBatch>();
    }
}
