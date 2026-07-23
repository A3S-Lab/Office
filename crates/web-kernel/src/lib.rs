//! Browser-safe A3S Office layout kernel.
//!
//! The crate intentionally has no filesystem, network, async-runtime, or DOM
//! dependency. Native tests exercise the same deterministic layout function
//! exported through the raw WebAssembly ABI.

use std::collections::HashSet;

use serde::{Deserialize, Serialize};

pub const OFFICE_KERNEL_PROTOCOL_VERSION: u32 = 1;
const MAX_LAYOUT_BLOCKS: usize = 10_000;
const MAX_LAYOUT_EXTENT: f64 = 1_000_000.0;

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutPageMetrics {
    pub width: f64,
    pub height: f64,
    pub margin_top: f64,
    pub margin_right: f64,
    pub margin_bottom: f64,
    pub margin_left: f64,
    #[serde(default)]
    pub header_height: f64,
    #[serde(default)]
    pub footer_height: f64,
    #[serde(default = "default_page_gap")]
    pub page_gap: f64,
}

impl LayoutPageMetrics {
    fn available_height(&self) -> f64 {
        (self.height
            - self.margin_top
            - self.margin_bottom
            - self.header_height
            - self.footer_height)
            .max(1.0)
    }

    fn validate(&self) -> Result<(), KernelError> {
        for (name, value) in [
            ("page.width", self.width),
            ("page.height", self.height),
            ("page.marginTop", self.margin_top),
            ("page.marginRight", self.margin_right),
            ("page.marginBottom", self.margin_bottom),
            ("page.marginLeft", self.margin_left),
            ("page.headerHeight", self.header_height),
            ("page.footerHeight", self.footer_height),
            ("page.pageGap", self.page_gap),
        ] {
            validate_extent(name, value)?;
        }
        if self.width <= self.margin_left + self.margin_right {
            return Err(KernelError::invalid(
                "office.kernel.page_width_invalid",
                "Page width must be greater than its horizontal margins.",
            ));
        }
        if self.height
            <= self.margin_top + self.margin_bottom + self.header_height + self.footer_height
        {
            return Err(KernelError::invalid(
                "office.kernel.page_height_invalid",
                "Page height must leave a positive body area.",
            ));
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutBlock {
    pub id: String,
    pub height: f64,
    #[serde(default)]
    pub break_before: bool,
    #[serde(default)]
    pub break_after: bool,
    #[serde(default)]
    pub keep_together: bool,
    #[serde(default)]
    pub keep_with_next: bool,
}

impl LayoutBlock {
    fn validate(&self) -> Result<(), KernelError> {
        if self.id.trim().is_empty() || self.id.len() > 256 {
            return Err(KernelError::invalid(
                "office.kernel.block_id_invalid",
                "Every layout block requires a non-empty ID of at most 256 bytes.",
            ));
        }
        validate_extent("block.height", self.height)
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutRequest {
    pub protocol: u32,
    pub kind: String,
    pub request_id: u32,
    pub revision: u32,
    pub page: LayoutPageMetrics,
    pub blocks: Vec<LayoutBlock>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutPlacement {
    pub block_id: String,
    pub y: f64,
    pub height: f64,
    pub overflow: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutPage {
    pub index: u32,
    pub used_height: f64,
    pub available_height: f64,
    pub placements: Vec<LayoutPlacement>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutBreak {
    pub before_block_id: String,
    pub page_index: u32,
    pub spacer_height: f64,
    pub remaining_body_height: f64,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutResult {
    pub protocol: u32,
    pub kind: String,
    pub request_id: u32,
    pub revision: u32,
    pub engine: String,
    pub pages: Vec<LayoutPage>,
    pub breaks: Vec<LayoutBreak>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct KernelError {
    pub code: String,
    pub message: String,
}

impl KernelError {
    fn invalid(code: impl Into<String>, message: impl Into<String>) -> Self {
        Self {
            code: code.into(),
            message: message.into(),
        }
    }
}

#[cfg(target_arch = "wasm32")]
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct KernelErrorResponse {
    protocol: u32,
    kind: String,
    request_id: u32,
    revision: u32,
    engine: String,
    error: KernelError,
}

pub fn layout_document(request: &LayoutRequest) -> Result<LayoutResult, KernelError> {
    validate_request(request)?;
    let available_height = request.page.available_height();
    let mut pages = vec![empty_page(0, available_height)];

    for (index, block) in request.blocks.iter().enumerate() {
        let current_has_content = pages.last().is_some_and(|page| !page.placements.is_empty());
        if block.break_before && current_has_content {
            pages.push(empty_page(pages.len() as u32, available_height));
        }

        let current = pages.last().expect("layout always contains one page");
        let remaining = (available_height - current.used_height).max(0.0);
        let next_height = request
            .blocks
            .get(index + 1)
            .filter(|next| !next.break_before)
            .map_or(0.0, |next| next.height);
        let grouped_height = block.height
            + if block.keep_with_next {
                next_height
            } else {
                0.0
            };
        let should_advance = current_has_content
            && (block.height > remaining
                || ((block.keep_together || block.keep_with_next)
                    && grouped_height <= available_height
                    && grouped_height > remaining));
        if should_advance {
            pages.push(empty_page(pages.len() as u32, available_height));
        }

        let current = pages.last_mut().expect("layout always contains one page");
        let y = current.used_height;
        current.placements.push(LayoutPlacement {
            block_id: block.id.clone(),
            y,
            height: block.height,
            overflow: block.height > available_height || y + block.height > available_height,
        });
        current.used_height += block.height;

        if block.break_after && index + 1 < request.blocks.len() {
            pages.push(empty_page(pages.len() as u32, available_height));
        }
    }

    if pages.len() > 1 && pages.last().is_some_and(|page| page.placements.is_empty()) {
        pages.pop();
    }
    let breaks = layout_breaks(&pages, &request.page);
    Ok(LayoutResult {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: "layoutResult".into(),
        request_id: request.request_id,
        revision: request.revision,
        engine: "wasm".into(),
        pages,
        breaks,
    })
}

fn validate_request(request: &LayoutRequest) -> Result<(), KernelError> {
    if request.protocol != OFFICE_KERNEL_PROTOCOL_VERSION {
        return Err(KernelError::invalid(
            "office.kernel.protocol_unsupported",
            format!(
                "Office kernel protocol {} is unsupported; expected {}.",
                request.protocol, OFFICE_KERNEL_PROTOCOL_VERSION
            ),
        ));
    }
    if request.kind != "layout" {
        return Err(KernelError::invalid(
            "office.kernel.request_kind_invalid",
            "The Office kernel only accepts layout requests at this boundary.",
        ));
    }
    if request.blocks.len() > MAX_LAYOUT_BLOCKS {
        return Err(KernelError::invalid(
            "office.kernel.block_limit_exceeded",
            format!("A layout request may contain at most {MAX_LAYOUT_BLOCKS} blocks."),
        ));
    }
    request.page.validate()?;
    let mut block_ids = HashSet::with_capacity(request.blocks.len());
    for block in &request.blocks {
        block.validate()?;
        if !block_ids.insert(block.id.as_str()) {
            return Err(KernelError::invalid(
                "office.kernel.block_id_duplicate",
                format!("Layout block ID '{}' is duplicated.", block.id),
            ));
        }
    }
    Ok(())
}

fn empty_page(index: u32, available_height: f64) -> LayoutPage {
    LayoutPage {
        index,
        used_height: 0.0,
        available_height,
        placements: Vec::new(),
    }
}

fn layout_breaks(pages: &[LayoutPage], metrics: &LayoutPageMetrics) -> Vec<LayoutBreak> {
    pages
        .iter()
        .skip(1)
        .filter_map(|page| {
            let before_block_id = page.placements.first()?.block_id.clone();
            let previous = &pages[page.index.saturating_sub(1) as usize];
            let remaining_body_height = (previous.available_height - previous.used_height).max(0.0);
            Some(LayoutBreak {
                before_block_id,
                page_index: page.index,
                spacer_height: remaining_body_height
                    + metrics.margin_bottom
                    + metrics.footer_height
                    + metrics.page_gap
                    + metrics.margin_top
                    + metrics.header_height,
                remaining_body_height,
            })
        })
        .collect()
}

fn validate_extent(name: &str, value: f64) -> Result<(), KernelError> {
    if !value.is_finite() || !(0.0..=MAX_LAYOUT_EXTENT).contains(&value) {
        return Err(KernelError::invalid(
            "office.kernel.extent_invalid",
            format!("{name} must be a finite non-negative number."),
        ));
    }
    Ok(())
}

const fn default_page_gap() -> f64 {
    24.0
}

#[cfg(target_arch = "wasm32")]
mod wasm_abi {
    use std::cell::RefCell;

    use super::{
        layout_document, KernelError, KernelErrorResponse, LayoutRequest,
        OFFICE_KERNEL_PROTOCOL_VERSION,
    };

    thread_local! {
        static LAST_RESULT: RefCell<Vec<u8>> = const { RefCell::new(Vec::new()) };
    }

    #[no_mangle]
    pub extern "C" fn office_kernel_abi_version() -> u32 {
        OFFICE_KERNEL_PROTOCOL_VERSION
    }

    #[no_mangle]
    pub extern "C" fn office_kernel_alloc(length: usize) -> *mut u8 {
        let bytes = vec![0_u8; length].into_boxed_slice();
        Box::into_raw(bytes).cast::<u8>()
    }

    #[no_mangle]
    pub unsafe extern "C" fn office_kernel_dealloc(pointer: *mut u8, length: usize) {
        if pointer.is_null() {
            return;
        }
        let slice = std::ptr::slice_from_raw_parts_mut(pointer, length);
        drop(Box::from_raw(slice));
    }

    #[no_mangle]
    pub unsafe extern "C" fn office_kernel_layout(pointer: *const u8, length: usize) -> i32 {
        let input = if pointer.is_null() {
            &[]
        } else {
            std::slice::from_raw_parts(pointer, length)
        };
        let parsed = serde_json::from_slice::<LayoutRequest>(input);
        let (status, output) = match parsed {
            Ok(request) => match layout_document(&request) {
                Ok(result) => (0, serde_json::to_vec(&result)),
                Err(error) => (
                    1,
                    serde_json::to_vec(&error_response(
                        request.request_id,
                        request.revision,
                        error,
                    )),
                ),
            },
            Err(error) => (
                1,
                serde_json::to_vec(&error_response(
                    0,
                    0,
                    KernelError {
                        code: "office.kernel.request_invalid".into(),
                        message: format!("The layout request is not valid JSON: {error}"),
                    },
                )),
            ),
        };
        LAST_RESULT.with(|result| {
            *result.borrow_mut() = output.unwrap_or_else(|error| {
                format!(
                    "{{\"protocol\":{},\"kind\":\"error\",\"requestId\":0,\"revision\":0,\"engine\":\"wasm\",\"error\":{{\"code\":\"office.kernel.serialization_failed\",\"message\":{:?}}}}}",
                    OFFICE_KERNEL_PROTOCOL_VERSION,
                    error.to_string()
                )
                .into_bytes()
            });
        });
        status
    }

    #[no_mangle]
    pub extern "C" fn office_kernel_result_pointer() -> *const u8 {
        LAST_RESULT.with(|result| result.borrow().as_ptr())
    }

    #[no_mangle]
    pub extern "C" fn office_kernel_result_length() -> usize {
        LAST_RESULT.with(|result| result.borrow().len())
    }

    fn error_response(request_id: u32, revision: u32, error: KernelError) -> KernelErrorResponse {
        KernelErrorResponse {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "error".into(),
            request_id,
            revision,
            engine: "wasm".into(),
            error,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn request(blocks: Vec<LayoutBlock>) -> LayoutRequest {
        LayoutRequest {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "layout".into(),
            request_id: 7,
            revision: 12,
            page: LayoutPageMetrics {
                width: 794.0,
                height: 1123.0,
                margin_top: 80.0,
                margin_right: 80.0,
                margin_bottom: 80.0,
                margin_left: 80.0,
                header_height: 0.0,
                footer_height: 0.0,
                page_gap: 24.0,
            },
            blocks,
        }
    }

    fn block(id: &str, height: f64) -> LayoutBlock {
        LayoutBlock {
            id: id.into(),
            height,
            break_before: false,
            break_after: false,
            keep_together: false,
            keep_with_next: false,
        }
    }

    #[test]
    fn paginates_blocks_and_reports_visual_spacers() {
        let result = layout_document(&request(vec![
            block("one", 500.0),
            block("two", 500.0),
            block("three", 120.0),
        ]))
        .expect("layout");

        assert_eq!(result.pages.len(), 2);
        assert_eq!(result.pages[0].placements.len(), 1);
        assert_eq!(result.pages[1].placements.len(), 2);
        assert_eq!(result.breaks[0].before_block_id, "two");
        assert_eq!(result.breaks[0].spacer_height, 647.0);
    }

    #[test]
    fn keeps_a_heading_with_the_following_block() {
        let mut heading = block("heading", 40.0);
        heading.keep_with_next = true;
        let result = layout_document(&request(vec![
            block("intro", 850.0),
            heading,
            block("paragraph", 100.0),
        ]))
        .expect("layout");

        assert_eq!(result.pages.len(), 2);
        assert_eq!(result.pages[1].placements[0].block_id, "heading");
        assert_eq!(result.pages[1].placements[1].block_id, "paragraph");
    }

    #[test]
    fn rejects_non_finite_or_impossible_page_metrics() {
        let mut invalid = request(vec![block("one", 10.0)]);
        invalid.page.margin_top = invalid.page.height;
        let error = layout_document(&invalid).expect_err("invalid page");
        assert_eq!(error.code, "office.kernel.page_height_invalid");
    }

    #[test]
    fn rejects_duplicate_block_ids() {
        let error = layout_document(&request(vec![block("same", 10.0), block("same", 20.0)]))
            .expect_err("duplicate block ID");
        assert_eq!(error.code, "office.kernel.block_id_duplicate");
    }
}
