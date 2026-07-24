use std::collections::HashSet;

use serde::{Deserialize, Serialize};

use crate::{KernelError, OFFICE_KERNEL_PROTOCOL_VERSION};

const MAX_PRESENTATION_ELEMENTS: usize = 10_000;
const MAX_PRESENTATION_EXTENT: f64 = 1_000_000.0;
const MAX_PRESENTATION_SNAP_THRESHOLD: f64 = 10.0;
const PRESENTATION_SLIDE_EXTENT: f64 = 100.0;
const PRESENTATION_SNAP_EPSILON: f64 = 0.000_001;

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

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum PresentationTransformMode {
    Move,
    Resize,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type")]
pub enum PresentationGeometryOperation {
    #[serde(rename = "alignToSlide")]
    AlignToSlide { alignment: PresentationAlignment },
    #[serde(rename = "snapElement")]
    SnapElement {
        #[serde(rename = "movingElementId")]
        moving_element_id: String,
        mode: PresentationTransformMode,
        #[serde(rename = "thresholdX")]
        threshold_x: f64,
        #[serde(rename = "thresholdY")]
        threshold_y: f64,
    },
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

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum PresentationSnapGuideAxis {
    X,
    Y,
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum PresentationSnapGuideSource {
    Element,
    Slide,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationSnapGuide {
    pub axis: PresentationSnapGuideAxis,
    pub position: f64,
    pub source: PresentationSnapGuideSource,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub target_id: Option<String>,
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
    pub guides: Vec<PresentationSnapGuide>,
}

pub fn resolve_presentation_geometry(
    request: &PresentationGeometryRequest,
) -> Result<PresentationGeometryResult, KernelError> {
    validate_request(request)?;
    let (elements, guides) = match &request.operation {
        PresentationGeometryOperation::AlignToSlide { alignment } => (
            request
                .elements
                .iter()
                .map(|element| align_element_to_slide(element, *alignment))
                .collect(),
            Vec::new(),
        ),
        PresentationGeometryOperation::SnapElement {
            moving_element_id,
            mode,
            threshold_x,
            threshold_y,
        } => snap_presentation_element(
            &request.elements,
            moving_element_id,
            *mode,
            *threshold_x,
            *threshold_y,
        )?,
    };
    Ok(PresentationGeometryResult {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: "presentationGeometryResult".into(),
        request_id: request.request_id,
        revision: request.revision,
        document_revision: request.document_revision,
        engine: "wasm".into(),
        elements,
        guides,
    })
}

fn align_element_to_slide(
    element: &PresentationGeometryElement,
    alignment: PresentationAlignment,
) -> PresentationGeometryElement {
    let mut aligned = element.clone();
    let maximum_x = (PRESENTATION_SLIDE_EXTENT - element.width).max(0.0);
    let maximum_y = (PRESENTATION_SLIDE_EXTENT - element.height).max(0.0);
    match alignment {
        PresentationAlignment::Left => aligned.x = 0.0,
        PresentationAlignment::Center => aligned.x = maximum_x / 2.0,
        PresentationAlignment::Right => aligned.x = maximum_x,
        PresentationAlignment::Top => aligned.y = 0.0,
        PresentationAlignment::Middle => aligned.y = maximum_y / 2.0,
        PresentationAlignment::Bottom => aligned.y = maximum_y,
    }
    aligned
}

fn snap_presentation_element(
    elements: &[PresentationGeometryElement],
    moving_element_id: &str,
    mode: PresentationTransformMode,
    threshold_x: f64,
    threshold_y: f64,
) -> Result<(Vec<PresentationGeometryElement>, Vec<PresentationSnapGuide>), KernelError> {
    let moving = elements
        .iter()
        .find(|element| element.id == moving_element_id)
        .ok_or_else(|| {
            KernelError::invalid(
                "office.kernel.moving_element_missing",
                format!("Presentation element '{moving_element_id}' does not exist."),
            )
        })?;
    let x_snap = snap_for_axis(
        moving,
        elements,
        PresentationSnapGuideAxis::X,
        mode,
        threshold_x,
    );
    let y_snap = snap_for_axis(
        moving,
        elements,
        PresentationSnapGuideAxis::Y,
        mode,
        threshold_y,
    );
    let snapped = apply_snap(moving, mode, x_snap.as_ref(), y_snap.as_ref());
    let guides = [x_snap, y_snap]
        .into_iter()
        .flatten()
        .filter(|snap| snap_applied(moving, &snapped, mode, snap))
        .map(|snap| snap.guide)
        .collect();
    Ok((
        elements
            .iter()
            .map(|element| {
                if element.id == moving_element_id {
                    snapped.clone()
                } else {
                    element.clone()
                }
            })
            .collect(),
        guides,
    ))
}

#[derive(Debug, Clone)]
struct PresentationSnap {
    axis: PresentationSnapGuideAxis,
    delta: f64,
    guide: PresentationSnapGuide,
}

#[derive(Debug, Clone)]
struct PresentationSnapCandidate {
    snap: PresentationSnap,
    distance: f64,
}

fn snap_for_axis(
    moving: &PresentationGeometryElement,
    elements: &[PresentationGeometryElement],
    axis: PresentationSnapGuideAxis,
    mode: PresentationTransformMode,
    threshold: f64,
) -> Option<PresentationSnap> {
    let mut targets = [
        0.0,
        PRESENTATION_SLIDE_EXTENT / 2.0,
        PRESENTATION_SLIDE_EXTENT,
    ]
    .into_iter()
    .map(|position| PresentationSnapGuide {
        axis,
        position,
        source: PresentationSnapGuideSource::Slide,
        target_id: None,
    })
    .collect::<Vec<_>>();
    for element in elements {
        if element.id == moving.id {
            continue;
        }
        targets.extend(element_anchors(element, axis).into_iter().map(|position| {
            PresentationSnapGuide {
                axis,
                position,
                source: PresentationSnapGuideSource::Element,
                target_id: Some(element.id.clone()),
            }
        }));
    }

    let mut best: Option<PresentationSnapCandidate> = None;
    for moving_position in moving_anchors(moving, axis, mode) {
        for target in &targets {
            let delta = target.position - moving_position;
            let distance = delta.abs();
            if distance > threshold
                || resize_snap_collapses_element(moving, axis, mode, delta)
                || best.as_ref().is_some_and(|candidate| {
                    distance >= candidate.distance - PRESENTATION_SNAP_EPSILON
                })
            {
                continue;
            }
            best = Some(PresentationSnapCandidate {
                snap: PresentationSnap {
                    axis,
                    delta,
                    guide: target.clone(),
                },
                distance,
            });
        }
    }
    best.map(|candidate| candidate.snap)
}

fn resize_snap_collapses_element(
    element: &PresentationGeometryElement,
    axis: PresentationSnapGuideAxis,
    mode: PresentationTransformMode,
    delta: f64,
) -> bool {
    if mode != PresentationTransformMode::Resize {
        return false;
    }
    let (_, size) = axis_extent(element, axis);
    size + delta <= 0.0
}

fn moving_anchors(
    element: &PresentationGeometryElement,
    axis: PresentationSnapGuideAxis,
    mode: PresentationTransformMode,
) -> Vec<f64> {
    let (start, size) = axis_extent(element, axis);
    match mode {
        PresentationTransformMode::Move => vec![start, start + size / 2.0, start + size],
        PresentationTransformMode::Resize => vec![start + size],
    }
}

fn element_anchors(
    element: &PresentationGeometryElement,
    axis: PresentationSnapGuideAxis,
) -> [f64; 3] {
    let (start, size) = axis_extent(element, axis);
    [start, start + size / 2.0, start + size]
}

fn axis_extent(
    element: &PresentationGeometryElement,
    axis: PresentationSnapGuideAxis,
) -> (f64, f64) {
    match axis {
        PresentationSnapGuideAxis::X => (element.x, element.width),
        PresentationSnapGuideAxis::Y => (element.y, element.height),
    }
}

fn apply_snap(
    element: &PresentationGeometryElement,
    mode: PresentationTransformMode,
    x_snap: Option<&PresentationSnap>,
    y_snap: Option<&PresentationSnap>,
) -> PresentationGeometryElement {
    let mut snapped = element.clone();
    match mode {
        PresentationTransformMode::Move => {
            snapped.x = clamp_extent(
                element.x + x_snap.map_or(0.0, |snap| snap.delta),
                PRESENTATION_SLIDE_EXTENT - element.width,
            );
            snapped.y = clamp_extent(
                element.y + y_snap.map_or(0.0, |snap| snap.delta),
                PRESENTATION_SLIDE_EXTENT - element.height,
            );
        }
        PresentationTransformMode::Resize => {
            snapped.width = clamp_extent(
                element.width + x_snap.map_or(0.0, |snap| snap.delta),
                PRESENTATION_SLIDE_EXTENT - element.x,
            );
            snapped.height = clamp_extent(
                element.height + y_snap.map_or(0.0, |snap| snap.delta),
                PRESENTATION_SLIDE_EXTENT - element.y,
            );
        }
    }
    snapped
}

fn snap_applied(
    before: &PresentationGeometryElement,
    after: &PresentationGeometryElement,
    mode: PresentationTransformMode,
    snap: &PresentationSnap,
) -> bool {
    let applied = match (mode, snap.axis) {
        (PresentationTransformMode::Move, PresentationSnapGuideAxis::X) => after.x - before.x,
        (PresentationTransformMode::Move, PresentationSnapGuideAxis::Y) => after.y - before.y,
        (PresentationTransformMode::Resize, PresentationSnapGuideAxis::X) => {
            after.width - before.width
        }
        (PresentationTransformMode::Resize, PresentationSnapGuideAxis::Y) => {
            after.height - before.height
        }
    };
    (applied - snap.delta).abs() <= PRESENTATION_SNAP_EPSILON
}

fn clamp_extent(value: f64, maximum: f64) -> f64 {
    value.max(0.0).min(maximum.max(0.0))
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
    if request.kind != "presentationGeometry" {
        return Err(KernelError::invalid(
            "office.kernel.request_kind_invalid",
            "The presentation geometry kernel received an invalid request kind.",
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
    if let PresentationGeometryOperation::SnapElement {
        moving_element_id,
        threshold_x,
        threshold_y,
        ..
    } = &request.operation
    {
        if moving_element_id.trim().is_empty() {
            return Err(KernelError::invalid(
                "office.kernel.moving_element_invalid",
                "A presentation snap request requires a moving element ID.",
            ));
        }
        for (name, value) in [("thresholdX", *threshold_x), ("thresholdY", *threshold_y)] {
            if !value.is_finite() || !(0.0..=MAX_PRESENTATION_SNAP_THRESHOLD).contains(&value) {
                return Err(KernelError::invalid(
                    "office.kernel.snap_threshold_invalid",
                    format!("{name} must be between 0 and {MAX_PRESENTATION_SNAP_THRESHOLD}."),
                ));
            }
        }
        if !ids.contains(moving_element_id.as_str()) {
            return Err(KernelError::invalid(
                "office.kernel.moving_element_missing",
                format!("Presentation element '{moving_element_id}' does not exist."),
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
