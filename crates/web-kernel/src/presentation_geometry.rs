use std::collections::HashSet;

use serde::{Deserialize, Serialize};

use crate::{KernelError, OFFICE_KERNEL_PROTOCOL_VERSION};

const MAX_PRESENTATION_ELEMENTS: usize = 10_000;
const MAX_PRESENTATION_EXTENT: f64 = 1_000_000.0;
const PRESENTATION_SLIDE_EXTENT: f64 = 100.0;

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum PresentationAlignment {
    Bottom,
    Center,
    Left,
    Middle,
    Right,
    Top,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationGeometryOperation {
    #[serde(rename = "type")]
    pub kind: String,
    pub alignment: PresentationAlignment,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationGeometryElement {
    pub id: String,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationGeometryRequest {
    pub protocol: u32,
    pub kind: String,
    pub request_id: u32,
    pub revision: u32,
    pub document_revision: u64,
    pub operation: PresentationGeometryOperation,
    pub elements: Vec<PresentationGeometryElement>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationGeometryResult {
    pub protocol: u32,
    pub kind: String,
    pub request_id: u32,
    pub revision: u32,
    pub document_revision: u64,
    pub engine: String,
    pub elements: Vec<PresentationGeometryElement>,
}

pub fn align_presentation_to_slide(
    request: &PresentationGeometryRequest,
) -> Result<PresentationGeometryResult, KernelError> {
    validate_request(request)?;
    let elements = request
        .elements
        .iter()
        .map(|element| {
            let mut aligned = element.clone();
            let maximum_x = (PRESENTATION_SLIDE_EXTENT - element.width).max(0.0);
            let maximum_y = (PRESENTATION_SLIDE_EXTENT - element.height).max(0.0);
            match request.operation.alignment {
                PresentationAlignment::Left => aligned.x = 0.0,
                PresentationAlignment::Center => aligned.x = maximum_x / 2.0,
                PresentationAlignment::Right => aligned.x = maximum_x,
                PresentationAlignment::Top => aligned.y = 0.0,
                PresentationAlignment::Middle => aligned.y = maximum_y / 2.0,
                PresentationAlignment::Bottom => aligned.y = maximum_y,
            }
            aligned
        })
        .collect();
    Ok(PresentationGeometryResult {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: "presentationGeometryResult".into(),
        request_id: request.request_id,
        revision: request.revision,
        document_revision: request.document_revision,
        engine: "wasm".into(),
        elements,
    })
}

fn validate_request(request: &PresentationGeometryRequest) -> Result<(), KernelError> {
    if request.protocol != OFFICE_KERNEL_PROTOCOL_VERSION {
        return Err(KernelError::invalid(
            "office.kernel.protocol_unsupported",
            format!(
                "Office kernel protocol {} is unsupported; expected {}.",
                request.protocol, OFFICE_KERNEL_PROTOCOL_VERSION
            ),
        ));
    }
    if request.kind != "presentationGeometry" || request.operation.kind != "alignToSlide" {
        return Err(KernelError::invalid(
            "office.kernel.request_kind_invalid",
            "The presentation geometry kernel only accepts slide-alignment requests.",
        ));
    }
    if request.elements.len() > MAX_PRESENTATION_ELEMENTS {
        return Err(KernelError::invalid(
            "office.kernel.element_limit_exceeded",
            format!(
                "A presentation geometry request may contain at most {MAX_PRESENTATION_ELEMENTS} elements."
            ),
        ));
    }
    let mut ids = HashSet::with_capacity(request.elements.len());
    for element in &request.elements {
        if element.id.trim().is_empty() || element.id.len() > 256 {
            return Err(KernelError::invalid(
                "office.kernel.element_id_invalid",
                "Every presentation element requires a non-empty ID of at most 256 bytes.",
            ));
        }
        if !ids.insert(element.id.as_str()) {
            return Err(KernelError::invalid(
                "office.kernel.element_id_duplicate",
                format!("Presentation element ID '{}' is duplicated.", element.id),
            ));
        }
        for (name, value) in [
            ("element.x", element.x),
            ("element.y", element.y),
            ("element.width", element.width),
            ("element.height", element.height),
        ] {
            validate_extent(name, value)?;
        }
        if element.width <= 0.0 || element.height <= 0.0 {
            return Err(KernelError::invalid(
                "office.kernel.element_size_invalid",
                "Presentation element width and height must be positive.",
            ));
        }
    }
    Ok(())
}

fn validate_extent(name: &str, value: f64) -> Result<(), KernelError> {
    if !value.is_finite() || !(0.0..=MAX_PRESENTATION_EXTENT).contains(&value) {
        return Err(KernelError::invalid(
            "office.kernel.extent_invalid",
            format!("{name} must be a finite non-negative number."),
        ));
    }
    Ok(())
}
