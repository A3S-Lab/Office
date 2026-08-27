use a3s_use_core::UseResult;
use pdfium_render::prelude::{
    PdfPage, PdfPageAnnotationCommon, PdfPageObjectCommon, PdfPageObjectType, PdfPageObjectsCommon,
    PdfRect,
};
use serde::{Deserialize, Serialize};

use super::super::engine::signed_points_to_millipoints;
use super::super::NativeOfficePdfPageBox;
use super::{invalid_text_layer, text_unsupported};

/// Hard bound for typed non-text page objects retained with one PDF text
/// layer. Pages above this limit keep the exact summary but mark geometry as
/// truncated so consumers must use a complete-page visual fallback.
pub const MAX_NATIVE_OFFICE_PDF_VISUAL_OBJECTS: usize = 4_096;

/// Closed inventory of renderable objects on the exact PDF page whose native
/// text layer was extracted.
///
/// Text-layer completeness alone cannot prove that a page image, vector path,
/// form object, annotation, or unsupported object contains no additional
/// visible text. Consumers can therefore route visual inspection from typed
/// PDF object evidence instead of content-length or page-coverage guesses.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfPageObjectSummary {
    pub total_objects: u32,
    pub text_objects: u32,
    pub path_objects: u32,
    pub image_objects: u32,
    pub shading_objects: u32,
    pub form_objects: u32,
    pub unsupported_objects: u32,
    pub annotations: u32,
}

impl NativeOfficePdfPageObjectSummary {
    /// Whether native text cannot by itself account for every renderable page
    /// object. This is deliberately conservative and contains no text/content
    /// heuristics.
    pub fn has_unverified_visual_content(self) -> bool {
        self.total_objects != self.text_objects || self.annotations != 0
    }

    pub(crate) fn validate(self) -> UseResult<()> {
        let classified = [
            self.text_objects,
            self.path_objects,
            self.image_objects,
            self.shading_objects,
            self.form_objects,
            self.unsupported_objects,
        ]
        .into_iter()
        .try_fold(0_u32, u32::checked_add)
        .ok_or_else(invalid_text_layer)?;
        if classified != self.total_objects {
            return Err(invalid_text_layer());
        }
        Ok(())
    }

    fn visual_object_count(self) -> UseResult<usize> {
        let count = self
            .total_objects
            .checked_sub(self.text_objects)
            .and_then(|count| count.checked_add(self.annotations))
            .ok_or_else(invalid_text_layer)?;
        usize::try_from(count).map_err(|_| invalid_text_layer())
    }
}

/// PDFium's closed non-text object families plus page annotations.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficePdfVisualObjectKind {
    Path,
    Image,
    Shading,
    Form,
    Unsupported,
    Annotation,
}

/// One non-text render object in stable page-object then annotation order.
/// Bounds are the conservative axis-aligned envelope of PDFium's exact object
/// quadrilateral. Missing or degenerate bounds force complete-page fallback.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfVisualObject {
    pub index: u32,
    pub kind: NativeOfficePdfVisualObjectKind,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bounds: Option<NativeOfficePdfPageBox>,
}

/// Bounded geometry for every potentially text-bearing non-text page object.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfVisualObjectInventory {
    pub max_objects: usize,
    pub truncated: bool,
    pub objects: Vec<NativeOfficePdfVisualObject>,
}

impl NativeOfficePdfVisualObjectInventory {
    /// True only when every non-text object and annotation has a finite,
    /// positive-area PDF-space envelope.
    pub fn has_complete_geometry(&self) -> bool {
        !self.truncated && self.objects.iter().all(|object| object.bounds.is_some())
    }

    pub(crate) fn validate(&self, summary: NativeOfficePdfPageObjectSummary) -> UseResult<()> {
        let expected = summary.visual_object_count()?;
        let cardinality_valid = if expected > self.max_objects {
            self.truncated && self.objects.is_empty()
        } else {
            !self.truncated && self.objects.len() == expected
        };
        let mut kind_counts = NativeOfficePdfPageObjectSummary::default();
        for object in &self.objects {
            let count = match object.kind {
                NativeOfficePdfVisualObjectKind::Path => &mut kind_counts.path_objects,
                NativeOfficePdfVisualObjectKind::Image => &mut kind_counts.image_objects,
                NativeOfficePdfVisualObjectKind::Shading => &mut kind_counts.shading_objects,
                NativeOfficePdfVisualObjectKind::Form => &mut kind_counts.form_objects,
                NativeOfficePdfVisualObjectKind::Unsupported => {
                    &mut kind_counts.unsupported_objects
                }
                NativeOfficePdfVisualObjectKind::Annotation => &mut kind_counts.annotations,
            };
            *count = count.checked_add(1).ok_or_else(invalid_text_layer)?;
        }
        let kinds_valid = self.truncated
            || (kind_counts.path_objects == summary.path_objects
                && kind_counts.image_objects == summary.image_objects
                && kind_counts.shading_objects == summary.shading_objects
                && kind_counts.form_objects == summary.form_objects
                && kind_counts.unsupported_objects == summary.unsupported_objects
                && kind_counts.annotations == summary.annotations);
        if self.max_objects != MAX_NATIVE_OFFICE_PDF_VISUAL_OBJECTS
            || !cardinality_valid
            || !kinds_valid
            || self.objects.iter().enumerate().any(|(offset, object)| {
                object.index != u32::try_from(offset.saturating_add(1)).unwrap_or(u32::MAX)
                    || object
                        .bounds
                        .is_some_and(|bounds| bounds.validate().is_err())
            })
        {
            return Err(invalid_text_layer());
        }
        Ok(())
    }
}

pub(super) fn page_object_inventory(
    page: &PdfPage<'_>,
) -> UseResult<(
    NativeOfficePdfPageObjectSummary,
    NativeOfficePdfVisualObjectInventory,
)> {
    let total_objects = u32::try_from(page.objects().len()).map_err(|_| text_unsupported())?;
    let annotations = u32::try_from(page.annotations().len()).map_err(|_| text_unsupported())?;
    let mut summary = NativeOfficePdfPageObjectSummary {
        total_objects,
        annotations,
        ..NativeOfficePdfPageObjectSummary::default()
    };
    for object in page.objects().iter() {
        let count = match object.object_type() {
            PdfPageObjectType::Text => &mut summary.text_objects,
            PdfPageObjectType::Path => &mut summary.path_objects,
            PdfPageObjectType::Image => &mut summary.image_objects,
            PdfPageObjectType::Shading => &mut summary.shading_objects,
            PdfPageObjectType::XObjectForm => &mut summary.form_objects,
            PdfPageObjectType::Unsupported => &mut summary.unsupported_objects,
        };
        *count = count.checked_add(1).ok_or_else(text_unsupported)?;
    }
    summary.validate()?;

    let expected = summary.visual_object_count()?;
    if expected > MAX_NATIVE_OFFICE_PDF_VISUAL_OBJECTS {
        return Ok((
            summary,
            NativeOfficePdfVisualObjectInventory {
                max_objects: MAX_NATIVE_OFFICE_PDF_VISUAL_OBJECTS,
                truncated: true,
                objects: Vec::new(),
            },
        ));
    }

    let mut objects = Vec::with_capacity(expected);
    for object in page.objects().iter() {
        let kind = match object.object_type() {
            PdfPageObjectType::Text => continue,
            PdfPageObjectType::Path => NativeOfficePdfVisualObjectKind::Path,
            PdfPageObjectType::Image => NativeOfficePdfVisualObjectKind::Image,
            PdfPageObjectType::Shading => NativeOfficePdfVisualObjectKind::Shading,
            PdfPageObjectType::XObjectForm => NativeOfficePdfVisualObjectKind::Form,
            PdfPageObjectType::Unsupported => NativeOfficePdfVisualObjectKind::Unsupported,
        };
        objects.push(NativeOfficePdfVisualObject {
            index: next_index(&objects)?,
            kind,
            bounds: object
                .bounds()
                .ok()
                .and_then(|bounds| page_box(bounds.to_rect())),
        });
    }
    let annotations = page.annotations();
    for annotation_index in 0..annotations.len() {
        objects.push(NativeOfficePdfVisualObject {
            index: next_index(&objects)?,
            kind: NativeOfficePdfVisualObjectKind::Annotation,
            bounds: annotations
                .get(annotation_index)
                .ok()
                .and_then(|annotation| annotation.bounds().ok())
                .and_then(page_box),
        });
    }
    let inventory = NativeOfficePdfVisualObjectInventory {
        max_objects: MAX_NATIVE_OFFICE_PDF_VISUAL_OBJECTS,
        truncated: false,
        objects,
    };
    inventory.validate(summary)?;
    Ok((summary, inventory))
}

fn next_index(objects: &[NativeOfficePdfVisualObject]) -> UseResult<u32> {
    u32::try_from(objects.len().saturating_add(1)).map_err(|_| text_unsupported())
}

fn page_box(rect: PdfRect) -> Option<NativeOfficePdfPageBox> {
    let bounds = NativeOfficePdfPageBox {
        left_millipoints: signed_points_to_millipoints(rect.left().value).ok()?,
        bottom_millipoints: signed_points_to_millipoints(rect.bottom().value).ok()?,
        right_millipoints: signed_points_to_millipoints(rect.right().value).ok()?,
        top_millipoints: signed_points_to_millipoints(rect.top().value).ok()?,
    };
    bounds.validate().ok().map(|()| bounds)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn visual_inventory_requires_exact_bounded_cardinality_and_geometry() {
        let summary = NativeOfficePdfPageObjectSummary {
            total_objects: 3,
            text_objects: 1,
            path_objects: 1,
            image_objects: 1,
            ..NativeOfficePdfPageObjectSummary::default()
        };
        let bounds = NativeOfficePdfPageBox {
            left_millipoints: 0,
            bottom_millipoints: 0,
            right_millipoints: 1_000,
            top_millipoints: 1_000,
        };
        let inventory = NativeOfficePdfVisualObjectInventory {
            max_objects: MAX_NATIVE_OFFICE_PDF_VISUAL_OBJECTS,
            truncated: false,
            objects: vec![
                NativeOfficePdfVisualObject {
                    index: 1,
                    kind: NativeOfficePdfVisualObjectKind::Path,
                    bounds: Some(bounds),
                },
                NativeOfficePdfVisualObject {
                    index: 2,
                    kind: NativeOfficePdfVisualObjectKind::Image,
                    bounds: Some(bounds),
                },
            ],
        };
        inventory.validate(summary).unwrap();
        assert!(inventory.has_complete_geometry());

        let mut missing = inventory.clone();
        missing.objects.pop();
        assert_eq!(
            missing.validate(summary).unwrap_err().code,
            "use.office.pdf_text_layer_invalid"
        );

        let mut unbounded = inventory;
        unbounded.objects[0].bounds = None;
        unbounded.validate(summary).unwrap();
        assert!(!unbounded.has_complete_geometry());

        let mut wrong_kind = unbounded;
        wrong_kind.objects[0].kind = NativeOfficePdfVisualObjectKind::Image;
        assert_eq!(
            wrong_kind.validate(summary).unwrap_err().code,
            "use.office.pdf_text_layer_invalid"
        );
    }
}
