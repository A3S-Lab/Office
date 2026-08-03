import type {
  WorkDocumentImageAlignment,
  WorkDocumentImageLayout,
  WorkDocumentImageProperties,
} from '../work-document-image-layout';

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
  alternativeText: string;
}

export interface DocumentPicturePropertiesErrors {
  width: string | null;
  height: string | null;
  wrapDistance: string | null;
}

const FALLBACK_IMAGE_WIDTH_PIXELS = 320;
const FALLBACK_IMAGE_HEIGHT_PIXELS = 180;
const MINIMUM_IMAGE_SIZE_CENTIMETERS = 0.01;
const MAXIMUM_IMAGE_SIZE_CENTIMETERS = 55.87;
const MAXIMUM_WRAP_DISTANCE_MILLIMETERS = 25;

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
  return {
    width: centimeters(width),
    height: centimeters(height),
    lockAspectRatio: source.properties.lockAspectRatio,
    aspectRatio: validAspectRatio(width, height),
    layout: source.properties.layout,
    alignment: source.properties.alignment,
    wrapDistance: formatNumber(source.properties.wrapDistance),
    alternativeText: source.properties.alternativeText,
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
  };
}

export function hasDocumentPicturePropertiesErrors(
  errors: DocumentPicturePropertiesErrors,
): boolean {
  return Boolean(errors.width || errors.height || errors.wrapDistance);
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
  const alternativeText = current.alternativeText.trim();
  if (initial.alternativeText.trim() !== alternativeText) {
    changes.alternativeText = alternativeText;
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
