import type {
  WorkDocumentImageAlignment,
  WorkDocumentImageHorizontalReference,
  WorkDocumentImageLayout,
  WorkDocumentImageProperties,
  WorkDocumentImageWrapSide,
  WorkDocumentImageVerticalReference,
} from '../work-document-image-layout';
import {
  MAX_DOCUMENT_IMAGE_RELATIVE_HEIGHT,
  normalizeDocumentImageLayer,
} from '../work-document-image-layout';
import { normalizeDocumentImageWrapSide } from '../work-document-image-wrap-contour';

export const IMAGE_PIXELS_PER_CENTIMETER = 96 / 2.54;

export interface DocumentPicturePropertiesSource {
  properties: WorkDocumentImageProperties;
  renderedWidth?: number;
  renderedHeight?: number;
}

export interface DocumentPicturePropertiesDraft {
  width: string;
  height: string;
  lockAspectRatio: boolean;
  aspectRatio: number;
  layout: WorkDocumentImageLayout;
  alignment: WorkDocumentImageAlignment;
  wrapDistance: string;
  wrapSide: WorkDocumentImageWrapSide;
  alternativeText: string;
  precisePosition: boolean;
  horizontalOffset: string;
  verticalOffset: string;
  horizontalReference: WorkDocumentImageHorizontalReference;
  verticalReference: WorkDocumentImageVerticalReference;
  cropTop: string;
  cropRight: string;
  cropBottom: string;
  cropLeft: string;
  relativeHeight: string;
  behindDocument: boolean;
  allowOverlap: boolean;
  layoutInCell: boolean;
  lockAnchor: boolean;
}

export interface DocumentPicturePropertiesErrors {
  width: string | null;
  height: string | null;
  wrapDistance: string | null;
  horizontalOffset: string | null;
  verticalOffset: string | null;
  crop: string | null;
  relativeHeight: string | null;
}

const FALLBACK_IMAGE_WIDTH_PIXELS = 320;
const FALLBACK_IMAGE_HEIGHT_PIXELS = 180;
const MINIMUM_IMAGE_SIZE_CENTIMETERS = 0.01;
const MAXIMUM_IMAGE_SIZE_CENTIMETERS = 55.87;
const MAXIMUM_WRAP_DISTANCE_MILLIMETERS = 25;
const MAXIMUM_IMAGE_OFFSET_MILLIMETERS = 558.7;

export function createDocumentPicturePropertiesDraft(
  source: DocumentPicturePropertiesSource,
): DocumentPicturePropertiesDraft {
  const width = effectiveDimension(
    source.properties.width,
    source.renderedWidth,
    FALLBACK_IMAGE_WIDTH_PIXELS,
  );
  const height = effectiveDimension(
    source.properties.height,
    source.renderedHeight,
    FALLBACK_IMAGE_HEIGHT_PIXELS,
  );
  const layer = normalizeDocumentImageLayer(source.properties.layer ?? {});
  return {
    width: centimeters(width),
    height: centimeters(height),
    lockAspectRatio: source.properties.lockAspectRatio,
    aspectRatio: validAspectRatio(width, height),
    layout: source.properties.layout,
    alignment: source.properties.alignment,
    wrapDistance: formatNumber(source.properties.wrapDistance),
    wrapSide: normalizeDocumentImageWrapSide(source.properties.wrapSide),
    alternativeText: source.properties.alternativeText,
    precisePosition: Boolean(source.properties.position),
    horizontalOffset: formatNumber(
      source.properties.position?.horizontalOffset ?? 0,
    ),
    verticalOffset: formatNumber(
      source.properties.position?.verticalOffset ?? 0,
    ),
    horizontalReference:
      source.properties.position?.horizontalReference ?? 'column',
    verticalReference:
      source.properties.position?.verticalReference ?? 'paragraph',
    cropTop: formatNumber(source.properties.crop?.top ?? 0),
    cropRight: formatNumber(source.properties.crop?.right ?? 0),
    cropBottom: formatNumber(source.properties.crop?.bottom ?? 0),
    cropLeft: formatNumber(source.properties.crop?.left ?? 0),
    relativeHeight: String(layer.relativeHeight),
    behindDocument: layer.behindDocument,
    allowOverlap: layer.allowOverlap,
    layoutInCell: layer.layoutInCell,
    lockAnchor: layer.lockAnchor,
  };
}

export function withDocumentPictureDimension(
  draft: DocumentPicturePropertiesDraft,
  dimension: 'height' | 'width',
  value: string,
): DocumentPicturePropertiesDraft {
  const next = { ...draft, [dimension]: value };
  const numeric = numericDraft(value);
  if (!draft.lockAspectRatio || numeric === null) return next;
  if (dimension === 'width') {
    return {
      ...next,
      height: formatNumber(numeric / draft.aspectRatio),
    };
  }
  return {
    ...next,
    width: formatNumber(numeric * draft.aspectRatio),
  };
}

export function withDocumentPictureAspectRatioLock(
  draft: DocumentPicturePropertiesDraft,
  lockAspectRatio: boolean,
): DocumentPicturePropertiesDraft {
  if (!lockAspectRatio) return { ...draft, lockAspectRatio: false };
  const width = numericDraft(draft.width);
  const height = numericDraft(draft.height);
  return {
    ...draft,
    lockAspectRatio: true,
    aspectRatio:
      width !== null && height !== null
        ? validAspectRatio(width, height)
        : draft.aspectRatio,
  };
}

export function documentPicturePropertiesErrors(
  draft: DocumentPicturePropertiesDraft,
): DocumentPicturePropertiesErrors {
  return {
    width: imageDimensionError(draft.width),
    height: imageDimensionError(draft.height),
    wrapDistance:
      draft.layout === 'inline' ||
      validNumber(draft.wrapDistance, 0, MAXIMUM_WRAP_DISTANCE_MILLIMETERS)
        ? null
        : '请输入 0 到 25 之间的毫米数。',
    horizontalOffset:
      !draft.precisePosition ||
      validSignedNumber(
        draft.horizontalOffset,
        MAXIMUM_IMAGE_OFFSET_MILLIMETERS,
      )
        ? null
        : '请输入 -558.7 到 558.7 之间的毫米数。',
    verticalOffset:
      !draft.precisePosition ||
      validSignedNumber(draft.verticalOffset, MAXIMUM_IMAGE_OFFSET_MILLIMETERS)
        ? null
        : '请输入 -558.7 到 558.7 之间的毫米数。',
    crop: imageCropError(draft),
    relativeHeight:
      draft.layout === 'inline' ||
      validInteger(draft.relativeHeight, 0, MAX_DOCUMENT_IMAGE_RELATIVE_HEIGHT)
        ? null
        : `请输入 0 到 ${MAX_DOCUMENT_IMAGE_RELATIVE_HEIGHT} 之间的整数。`,
  };
}

export function hasDocumentPicturePropertiesErrors(
  errors: DocumentPicturePropertiesErrors,
): boolean {
  return Boolean(
    errors.width ||
      errors.height ||
      errors.wrapDistance ||
      errors.horizontalOffset ||
      errors.verticalOffset ||
      errors.crop ||
      errors.relativeHeight,
  );
}

export function documentPicturePropertyChanges(
  initial: DocumentPicturePropertiesDraft,
  current: DocumentPicturePropertiesDraft,
): Partial<WorkDocumentImageProperties> | null {
  const errors = documentPicturePropertiesErrors(current);
  if (hasDocumentPicturePropertiesErrors(errors)) return null;

  const changes: Partial<WorkDocumentImageProperties> = {};
  if (!sameDraftNumber(initial.width, current.width)) {
    changes.width = rounded(
      Number(current.width) * IMAGE_PIXELS_PER_CENTIMETER,
    );
  }
  if (!sameDraftNumber(initial.height, current.height)) {
    changes.height = rounded(
      Number(current.height) * IMAGE_PIXELS_PER_CENTIMETER,
    );
  }
  if (initial.lockAspectRatio !== current.lockAspectRatio) {
    changes.lockAspectRatio = current.lockAspectRatio;
  }
  if (initial.layout !== current.layout) changes.layout = current.layout;
  if (initial.alignment !== current.alignment) {
    changes.alignment = current.alignment;
  }
  if (!sameDraftNumber(initial.wrapDistance, current.wrapDistance)) {
    changes.wrapDistance = Number(current.wrapDistance);
  }
  if (initial.wrapSide !== current.wrapSide) {
    changes.wrapSide = current.wrapSide;
  }
  const alternativeText = current.alternativeText.trim();
  if (initial.alternativeText.trim() !== alternativeText) {
    changes.alternativeText = alternativeText;
  }
  if (
    initial.precisePosition !== current.precisePosition ||
    (current.precisePosition &&
      (!sameDraftNumber(initial.horizontalOffset, current.horizontalOffset) ||
        !sameDraftNumber(initial.verticalOffset, current.verticalOffset) ||
        initial.horizontalReference !== current.horizontalReference ||
        initial.verticalReference !== current.verticalReference))
  ) {
    changes.position = current.precisePosition
      ? {
          horizontalOffset: Number(current.horizontalOffset),
          verticalOffset: Number(current.verticalOffset),
          horizontalReference: current.horizontalReference,
          verticalReference: current.verticalReference,
        }
      : null;
  }
  if (
    !sameDraftNumber(initial.cropTop, current.cropTop) ||
    !sameDraftNumber(initial.cropRight, current.cropRight) ||
    !sameDraftNumber(initial.cropBottom, current.cropBottom) ||
    !sameDraftNumber(initial.cropLeft, current.cropLeft)
  ) {
    const crop = {
      top: Number(current.cropTop),
      right: Number(current.cropRight),
      bottom: Number(current.cropBottom),
      left: Number(current.cropLeft),
    };
    changes.crop = Object.values(crop).some((edge) => edge > 0) ? crop : null;
  }
  if (
    !sameDraftNumber(initial.relativeHeight, current.relativeHeight) ||
    initial.behindDocument !== current.behindDocument ||
    initial.allowOverlap !== current.allowOverlap ||
    initial.layoutInCell !== current.layoutInCell ||
    initial.lockAnchor !== current.lockAnchor
  ) {
    changes.layer = normalizeDocumentImageLayer({
      relativeHeight: current.relativeHeight,
      behindDocument: current.behindDocument,
      allowOverlap: current.allowOverlap,
      layoutInCell: current.layoutInCell,
      lockAnchor: current.lockAnchor,
    });
  }
  return Object.keys(changes).length ? changes : null;
}

function effectiveDimension(
  explicit: number | null,
  rendered: number | undefined,
  fallback: number,
): number {
  if (explicit !== null && Number.isFinite(explicit) && explicit > 0) {
    return explicit;
  }
  if (rendered !== undefined && Number.isFinite(rendered) && rendered > 0) {
    return rendered;
  }
  return fallback;
}

function imageDimensionError(value: string): string | null {
  return validNumber(
    value,
    MINIMUM_IMAGE_SIZE_CENTIMETERS,
    MAXIMUM_IMAGE_SIZE_CENTIMETERS,
  )
    ? null
    : '请输入 0.01 到 55.87 之间的厘米数。';
}

function validNumber(value: string, minimum: number, maximum: number): boolean {
  const number = numericDraft(value);
  return number !== null && number >= minimum && number <= maximum;
}

function validSignedNumber(value: string, maximumMagnitude: number): boolean {
  if (!value.trim()) return false;
  const number = Number(value);
  return Number.isFinite(number) && Math.abs(number) <= maximumMagnitude;
}

function validInteger(
  value: string,
  minimum: number,
  maximum: number,
): boolean {
  if (!value.trim()) return false;
  const number = Number(value);
  return Number.isInteger(number) && number >= minimum && number <= maximum;
}

function imageCropError(draft: DocumentPicturePropertiesDraft): string | null {
  const edges = [
    draft.cropTop,
    draft.cropRight,
    draft.cropBottom,
    draft.cropLeft,
  ].map(numericDraft);
  if (edges.some((edge) => edge === null || edge > 99.99)) {
    return '请输入 0 到 99.99 之间的裁剪百分比。';
  }
  const [top = 0, right = 0, bottom = 0, left = 0] = edges as number[];
  return left + right < 100 && top + bottom < 100
    ? null
    : '相对两边的裁剪量之和必须小于 100%。';
}

function numericDraft(value: string): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number) && number >= 0 ? number : null;
}

function sameDraftNumber(left: string, right: string): boolean {
  if (!left.trim() || !right.trim()) return left === right;
  return Number(left) === Number(right);
}

function validAspectRatio(width: number, height: number): number {
  const ratio = width / height;
  return Number.isFinite(ratio) && ratio > 0 ? ratio : 1;
}

function centimeters(pixels: number): string {
  return formatNumber(pixels / IMAGE_PIXELS_PER_CENTIMETER);
}

function formatNumber(value: number): string {
  return Number.isFinite(value) ? value.toFixed(2).replace(/\.?0+$/, '') : '';
}

function rounded(value: number): number {
  return Math.round(value * 100) / 100;
}
