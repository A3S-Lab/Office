//! Browser-safe A3S Office layout kernel.
//!
//! The crate intentionally has no filesystem, network, async-runtime, or DOM
//! dependency. Native tests exercise the same deterministic layout function
//! exported through the raw WebAssembly ABI.

use std::collections::HashSet;

use serde::{Deserialize, Serialize};

mod presentation_geometry;
mod text_layout;

pub use presentation_geometry::{
    align_presentation_to_slide, PresentationAlignment, PresentationGeometryElement,
    PresentationGeometryOperation, PresentationGeometryRequest, PresentationGeometryResult,
};
pub use text_layout::{
    layout_text, validate_font, FontRegistry, TextDirection, TextLayoutLine, TextLayoutParagraph,
    TextLayoutParagraphResult, TextLayoutRequest, TextLayoutResult, TextLayoutRun,
    TextTabAlignment, TextTabLayout, TextTabStop, TextWhiteSpace,
};

pub const OFFICE_KERNEL_PROTOCOL_VERSION: u32 = 11;
const MAX_LAYOUT_BLOCKS: usize = 10_000;
const MAX_LAYOUT_EXTENT: f64 = 1_000_000.0;
const MAX_LAYOUT_PAGE_INDEX: u32 = 1_000_000;

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
    #[serde(default)]
    pub flow_id: Option<String>,
    #[serde(default)]
    pub flow_index: Option<u32>,
    #[serde(default)]
    pub flow_count: Option<u32>,
    #[serde(default)]
    pub minimum_fragments_per_page: Option<u32>,
    #[serde(default)]
    pub repeat_header_count: Option<u32>,
    #[serde(default)]
    pub repeat_header_height: Option<f64>,
}

impl LayoutBlock {
    fn validate(&self) -> Result<(), KernelError> {
        if self.id.trim().is_empty() || self.id.len() > 256 {
            return Err(KernelError::invalid(
                "office.kernel.block_id_invalid",
                "Every layout block requires a non-empty ID of at most 256 bytes.",
            ));
        }
        validate_extent("block.height", self.height)?;
        match self.flow_id.as_deref() {
            None => {
                if self.flow_index.is_some()
                    || self.flow_count.is_some()
                    || self.minimum_fragments_per_page.is_some()
                    || self.repeat_header_count.is_some()
                    || self.repeat_header_height.is_some()
                {
                    return Err(KernelError::invalid(
                        "office.kernel.flow_metadata_incomplete",
                        "Flow metadata requires flowId, flowIndex, flowCount, and minimumFragmentsPerPage.",
                    ));
                }
            }
            Some(flow_id) => {
                if flow_id.trim().is_empty() || flow_id.len() > 256 {
                    return Err(KernelError::invalid(
                        "office.kernel.flow_id_invalid",
                        "Every layout flow requires a non-empty ID of at most 256 bytes.",
                    ));
                }
                let Some(flow_index) = self.flow_index else {
                    return Err(KernelError::invalid(
                        "office.kernel.flow_metadata_incomplete",
                        "Flow metadata requires flowId, flowIndex, flowCount, and minimumFragmentsPerPage.",
                    ));
                };
                let Some(flow_count) = self.flow_count else {
                    return Err(KernelError::invalid(
                        "office.kernel.flow_metadata_incomplete",
                        "Flow metadata requires flowId, flowIndex, flowCount, and minimumFragmentsPerPage.",
                    ));
                };
                let Some(minimum) = self.minimum_fragments_per_page else {
                    return Err(KernelError::invalid(
                        "office.kernel.flow_metadata_incomplete",
                        "Flow metadata requires flowId, flowIndex, flowCount, and minimumFragmentsPerPage.",
                    ));
                };
                if flow_count == 0 || flow_index >= flow_count {
                    return Err(KernelError::invalid(
                        "office.kernel.flow_index_invalid",
                        "Flow indices must be within a non-empty flow.",
                    ));
                }
                if minimum == 0 || minimum > flow_count {
                    return Err(KernelError::invalid(
                        "office.kernel.flow_minimum_invalid",
                        "minimumFragmentsPerPage must be within the flow fragment count.",
                    ));
                }
                match (self.repeat_header_count, self.repeat_header_height) {
                    (None, None) => {}
                    (Some(count), Some(height)) => {
                        if count == 0 || count >= flow_count {
                            return Err(KernelError::invalid(
                                "office.kernel.repeat_header_count_invalid",
                                "repeatHeaderCount must identify a non-empty header before the body rows.",
                            ));
                        }
                        validate_extent("block.repeatHeaderHeight", height)?;
                        if height <= 0.0 {
                            return Err(KernelError::invalid(
                                "office.kernel.repeat_header_height_invalid",
                                "repeatHeaderHeight must be greater than zero.",
                            ));
                        }
                    }
                    _ => {
                        return Err(KernelError::invalid(
                            "office.kernel.repeat_header_metadata_incomplete",
                            "Repeated table headers require repeatHeaderCount and repeatHeaderHeight.",
                        ));
                    }
                }
            }
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutRequest {
    pub protocol: u32,
    pub kind: String,
    pub request_id: u32,
    pub revision: u32,
    pub document_revision: u64,
    pub start_page_index: u32,
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
    pub document_revision: u64,
    pub start_page_index: u32,
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
    document_revision: u64,
    engine: String,
    error: KernelError,
}

pub fn layout_document(request: &LayoutRequest) -> Result<LayoutResult, KernelError> {
    validate_request(request)?;
    let available_height = request.page.available_height();
    let mut pages = vec![empty_page(request.start_page_index, available_height)];
    let mut index = 0;
    while index < request.blocks.len() {
        let block = &request.blocks[index];
        if let Some(flow_count) = block.flow_count {
            let end = index + flow_count as usize;
            layout_flow(
                &request.blocks[index..end],
                &request.blocks[end..],
                &mut pages,
                available_height,
                end < request.blocks.len(),
            );
            index = end;
        } else {
            layout_single_block(
                block,
                &request.blocks[index + 1..],
                &mut pages,
                available_height,
                index + 1 < request.blocks.len(),
            );
            index += 1;
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
        document_revision: request.document_revision,
        start_page_index: request.start_page_index,
        engine: "wasm".into(),
        pages,
        breaks,
    })
}

fn layout_single_block(
    block: &LayoutBlock,
    next: &[LayoutBlock],
    pages: &mut Vec<LayoutPage>,
    available_height: f64,
    has_more_blocks: bool,
) {
    let current_has_content = pages.last().is_some_and(|page| !page.placements.is_empty());
    if block.break_before && current_has_content {
        pages.push(empty_page(next_page_index(pages), available_height));
    }

    let current = pages.last().expect("layout always contains one page");
    let remaining = (available_height - current.used_height).max(0.0);
    let next_height = next_block_preview_height(next);
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
        pages.push(empty_page(next_page_index(pages), available_height));
    }

    place_fragment(block, pages, available_height);
    if block.break_after && has_more_blocks {
        pages.push(empty_page(next_page_index(pages), available_height));
    }
}

fn layout_flow(
    blocks: &[LayoutBlock],
    next: &[LayoutBlock],
    pages: &mut Vec<LayoutPage>,
    available_height: f64,
    has_more_blocks: bool,
) {
    let first = blocks.first().expect("validated flow is non-empty");
    let last = blocks.last().expect("validated flow is non-empty");
    if first.break_before && pages.last().is_some_and(|page| !page.placements.is_empty()) {
        pages.push(empty_page(next_page_index(pages), available_height));
    }

    let total_height = blocks.iter().map(|block| block.height).sum::<f64>();
    let next_height = if last.keep_with_next {
        next_block_preview_height(next)
    } else {
        0.0
    };
    let current = pages.last().expect("layout always contains one page");
    let current_has_content = !current.placements.is_empty();
    let remaining = (available_height - current.used_height).max(0.0);
    let repeat_header_count = first.repeat_header_count.unwrap_or(0) as usize;
    let repeat_header_height = first.repeat_header_height.unwrap_or(0.0);
    let grouped_height = total_height + next_height;
    if first.keep_together
        && current_has_content
        && grouped_height <= available_height
        && grouped_height > remaining
    {
        pages.push(empty_page(next_page_index(pages), available_height));
    }
    if repeat_header_count > 0 && repeat_header_count < blocks.len() {
        let leading_height = blocks[..repeat_header_count + 1]
            .iter()
            .map(|block| block.height)
            .sum::<f64>();
        let current = pages.last().expect("layout always contains one page");
        let remaining = (available_height - current.used_height).max(0.0);
        if !current.placements.is_empty()
            && leading_height <= available_height
            && leading_height > remaining
        {
            pages.push(empty_page(next_page_index(pages), available_height));
        }
    }

    let minimum = first.minimum_fragments_per_page.unwrap_or(1).max(1) as usize;
    let mut cursor = 0;
    while cursor < blocks.len() {
        if cursor >= repeat_header_count
            && cursor > 0
            && repeat_header_height > 0.0
            && pages
                .last()
                .is_some_and(|page| page.placements.is_empty() && page.used_height == 0.0)
        {
            pages
                .last_mut()
                .expect("layout always contains one page")
                .used_height = repeat_header_height;
        }
        let current = pages.last().expect("layout always contains one page");
        let current_has_content = !current.placements.is_empty();
        let remaining_height = (available_height - current.used_height).max(0.0);
        let remaining_fragments = blocks.len() - cursor;
        let remaining_flow_height = blocks[cursor..]
            .iter()
            .map(|block| block.height)
            .sum::<f64>();
        if next_height > 0.0
            && current_has_content
            && remaining_flow_height <= remaining_height
            && remaining_flow_height + next_height > remaining_height
            && remaining_flow_height + next_height <= available_height
        {
            pages.push(empty_page(next_page_index(pages), available_height));
            continue;
        }
        let mut fitting = fragments_fitting(&blocks[cursor..], remaining_height);

        if fitting == remaining_fragments {
            place_fragments(&blocks[cursor..], pages, available_height);
            cursor = blocks.len();
            continue;
        }
        if fitting == 0 {
            if current_has_content {
                pages.push(empty_page(next_page_index(pages), available_height));
                continue;
            }
            fitting = 1;
        }

        if remaining_fragments > minimum {
            fitting = fitting.min(remaining_fragments - minimum);
        }
        let minimum_here = minimum.min(remaining_fragments);
        if fitting < minimum_here && current_has_content {
            pages.push(empty_page(next_page_index(pages), available_height));
            continue;
        }
        let fitting = fitting.max(1);
        place_fragments(&blocks[cursor..cursor + fitting], pages, available_height);
        cursor += fitting;
        if cursor < blocks.len() {
            pages.push(empty_page(next_page_index(pages), available_height));
        }
    }

    if last.break_after && has_more_blocks {
        pages.push(empty_page(next_page_index(pages), available_height));
    }
}

fn next_block_preview_height(blocks: &[LayoutBlock]) -> f64 {
    let Some(first) = blocks.first().filter(|block| !block.break_before) else {
        return 0.0;
    };
    let count = first.flow_count.unwrap_or(1).min(2) as usize;
    blocks.iter().take(count).map(|block| block.height).sum()
}

fn fragments_fitting(blocks: &[LayoutBlock], available_height: f64) -> usize {
    let mut used = 0.0;
    blocks
        .iter()
        .take_while(|block| {
            let fits = used + block.height <= available_height;
            if fits {
                used += block.height;
            }
            fits
        })
        .count()
}

fn place_fragments(blocks: &[LayoutBlock], pages: &mut Vec<LayoutPage>, available_height: f64) {
    for block in blocks {
        place_fragment(block, pages, available_height);
    }
}

fn place_fragment(block: &LayoutBlock, pages: &mut Vec<LayoutPage>, available_height: f64) {
    let current = pages.last_mut().expect("layout always contains one page");
    let y = current.used_height;
    current.placements.push(LayoutPlacement {
        block_id: block.id.clone(),
        y,
        height: block.height,
        overflow: block.height > available_height || y + block.height > available_height,
    });
    current.used_height += block.height;
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
    if request.start_page_index > MAX_LAYOUT_PAGE_INDEX {
        return Err(KernelError::invalid(
            "office.kernel.page_index_invalid",
            format!("startPageIndex may not exceed {MAX_LAYOUT_PAGE_INDEX}."),
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
    validate_flows(&request.blocks)?;
    Ok(())
}

fn validate_flows(blocks: &[LayoutBlock]) -> Result<(), KernelError> {
    let mut flow_ids = HashSet::new();
    let mut index = 0;
    while index < blocks.len() {
        let block = &blocks[index];
        let Some(flow_id) = block.flow_id.as_deref() else {
            index += 1;
            continue;
        };
        if block.flow_index != Some(0) {
            return Err(KernelError::invalid(
                "office.kernel.flow_sequence_invalid",
                "A layout flow must begin with fragment index zero.",
            ));
        }
        if !flow_ids.insert(flow_id) {
            return Err(KernelError::invalid(
                "office.kernel.flow_id_duplicate",
                format!("Layout flow ID '{flow_id}' is duplicated."),
            ));
        }
        let count = block.flow_count.unwrap_or_default() as usize;
        let minimum = block.minimum_fragments_per_page;
        let repeat_header_count = block.repeat_header_count;
        let repeat_header_height = block.repeat_header_height;
        let Some(flow) = blocks.get(index..index + count) else {
            return Err(KernelError::invalid(
                "office.kernel.flow_sequence_invalid",
                "A layout flow must contain its declared number of fragments.",
            ));
        };
        for (flow_index, fragment) in flow.iter().enumerate() {
            if fragment.flow_id.as_deref() != Some(flow_id)
                || fragment.flow_index != Some(flow_index as u32)
                || fragment.flow_count != Some(count as u32)
                || fragment.minimum_fragments_per_page != minimum
                || fragment.repeat_header_count != repeat_header_count
                || fragment.repeat_header_height != repeat_header_height
            {
                return Err(KernelError::invalid(
                    "office.kernel.flow_sequence_invalid",
                    "Layout flow fragments must be consecutive and consistently indexed.",
                ));
            }
            if flow_index > 0 && fragment.break_before {
                return Err(KernelError::invalid(
                    "office.kernel.flow_break_invalid",
                    "Only the first flow fragment may request breakBefore.",
                ));
            }
            if flow_index + 1 < count && fragment.break_after {
                return Err(KernelError::invalid(
                    "office.kernel.flow_break_invalid",
                    "Only the last flow fragment may request breakAfter.",
                ));
            }
        }
        index += count;
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

fn next_page_index(pages: &[LayoutPage]) -> u32 {
    pages.last().map_or(0, |page| page.index.saturating_add(1))
}

fn layout_breaks(pages: &[LayoutPage], metrics: &LayoutPageMetrics) -> Vec<LayoutBreak> {
    pages
        .iter()
        .enumerate()
        .skip(1)
        .filter_map(|(local_index, page)| {
            let before_block_id = page.placements.first()?.block_id.clone();
            let previous = &pages[local_index - 1];
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
        align_presentation_to_slide, layout_document, layout_text, validate_font, FontRegistry,
        KernelError, KernelErrorResponse, LayoutRequest, PresentationGeometryRequest,
        TextLayoutRequest, OFFICE_KERNEL_PROTOCOL_VERSION,
    };

    thread_local! {
        static LAST_RESULT: RefCell<Vec<u8>> = const { RefCell::new(Vec::new()) };
        static FONTS: RefCell<FontRegistry> = RefCell::new(FontRegistry::new());
    }

    const MAX_REGISTERED_FONTS: usize = 16;
    const MAX_REGISTERED_FONT_BYTES: usize = 32 * 1024 * 1024;

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
    pub unsafe extern "C" fn office_kernel_register_font(
        id_pointer: *const u8,
        id_length: usize,
        data_pointer: *const u8,
        data_length: usize,
    ) -> i32 {
        if id_pointer.is_null()
            || data_pointer.is_null()
            || id_length == 0
            || data_length == 0
            || data_length > MAX_REGISTERED_FONT_BYTES
        {
            return 1;
        }
        let id_bytes = std::slice::from_raw_parts(id_pointer, id_length);
        let data = std::slice::from_raw_parts(data_pointer, data_length);
        let Ok(id) = std::str::from_utf8(id_bytes) else {
            return 1;
        };
        if validate_font(id, data).is_err() {
            return 1;
        }
        FONTS.with(|fonts| {
            let mut fonts = fonts.borrow_mut();
            if !fonts.contains_key(id) && fonts.len() >= MAX_REGISTERED_FONTS {
                return 1;
            }
            fonts.insert(id.to_owned(), data.to_vec());
            0
        })
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
                        request.document_revision,
                        error,
                    )),
                ),
            },
            Err(error) => (
                1,
                serde_json::to_vec(&error_response(
                    0,
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
                    "{{\"protocol\":{},\"kind\":\"error\",\"requestId\":0,\"revision\":0,\"documentRevision\":0,\"engine\":\"wasm\",\"error\":{{\"code\":\"office.kernel.serialization_failed\",\"message\":{:?}}}}}",
                    OFFICE_KERNEL_PROTOCOL_VERSION,
                    error.to_string()
                )
                .into_bytes()
            });
        });
        status
    }

    #[no_mangle]
    pub unsafe extern "C" fn office_kernel_presentation_geometry(
        pointer: *const u8,
        length: usize,
    ) -> i32 {
        let input = if pointer.is_null() {
            &[]
        } else {
            std::slice::from_raw_parts(pointer, length)
        };
        let parsed = serde_json::from_slice::<PresentationGeometryRequest>(input);
        let (status, output) = match parsed {
            Ok(request) => match align_presentation_to_slide(&request) {
                Ok(result) => (0, serde_json::to_vec(&result)),
                Err(error) => (
                    1,
                    serde_json::to_vec(&error_response(
                        request.request_id,
                        request.revision,
                        request.document_revision,
                        error,
                    )),
                ),
            },
            Err(error) => (
                1,
                serde_json::to_vec(&error_response(
                    0,
                    0,
                    0,
                    KernelError {
                        code: "office.kernel.request_invalid".into(),
                        message: format!(
                            "The presentation geometry request is not valid JSON: {error}"
                        ),
                    },
                )),
            ),
        };
        LAST_RESULT.with(|result| {
            *result.borrow_mut() = output.unwrap_or_else(|error| {
                format!(
                    "{{\"protocol\":{},\"kind\":\"error\",\"requestId\":0,\"revision\":0,\"documentRevision\":0,\"engine\":\"wasm\",\"error\":{{\"code\":\"office.kernel.serialization_failed\",\"message\":{:?}}}}}",
                    OFFICE_KERNEL_PROTOCOL_VERSION,
                    error.to_string()
                )
                .into_bytes()
            });
        });
        status
    }

    #[no_mangle]
    pub unsafe extern "C" fn office_kernel_text_layout(pointer: *const u8, length: usize) -> i32 {
        let input = if pointer.is_null() {
            &[]
        } else {
            std::slice::from_raw_parts(pointer, length)
        };
        let parsed = serde_json::from_slice::<TextLayoutRequest>(input);
        let (status, output) = match parsed {
            Ok(request) => FONTS.with(|fonts| match layout_text(&request, &fonts.borrow()) {
                Ok(result) => (0, serde_json::to_vec(&result)),
                Err(error) => (
                    1,
                    serde_json::to_vec(&error_response(
                        request.request_id,
                        request.revision,
                        request.document_revision,
                        error,
                    )),
                ),
            }),
            Err(error) => (
                1,
                serde_json::to_vec(&error_response(
                    0,
                    0,
                    0,
                    KernelError {
                        code: "office.kernel.request_invalid".into(),
                        message: format!("The text layout request is not valid JSON: {error}"),
                    },
                )),
            ),
        };
        LAST_RESULT.with(|result| {
            *result.borrow_mut() = output.unwrap_or_else(|error| {
                format!(
                    "{{\"protocol\":{},\"kind\":\"error\",\"requestId\":0,\"revision\":0,\"documentRevision\":0,\"engine\":\"wasm\",\"error\":{{\"code\":\"office.kernel.serialization_failed\",\"message\":{:?}}}}}",
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

    fn error_response(
        request_id: u32,
        revision: u32,
        document_revision: u64,
        error: KernelError,
    ) -> KernelErrorResponse {
        KernelErrorResponse {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "error".into(),
            request_id,
            revision,
            document_revision,
            engine: "wasm".into(),
            error,
        }
    }
}

#[cfg(test)]
#[path = "layout_tests.rs"]
mod tests;
