import { attribute } from './work-ooxml-package';

/**
 * Bounded Traditional Office / WPS worksheet-picture transforms.
 *
 * Admit only quadrant rotations (0/90/180/270) and boolean flips so XLSX
 * round trips stay editable instead of silently dropping geometry.
 */

export interface WorkSpreadsheetImageTransform {
  /** Degrees clockwise; only 0, 90, 180, and 270 are editable. */
  rotation: number;
  flipHorizontal: boolean;
  flipVertical: boolean;
}

const XLSX_ROTATION_UNITS_PER_DEGREE = 60_000;
const XLSX_QUADRANT_ROTATION_UNITS = 5_400_000;

export function normalizeSpreadsheetImageTransform(
  transform: WorkSpreadsheetImageTransform,
): WorkSpreadsheetImageTransform {
  const rotation = ((Math.round(transform.rotation / 90) % 4) + 4) % 4;
  return {
    rotation: rotation * 90,
    flipHorizontal: Boolean(transform.flipHorizontal),
    flipVertical: Boolean(transform.flipVertical),
  };
}

export function isIdentitySpreadsheetImageTransform(
  transform: WorkSpreadsheetImageTransform | null | undefined,
): boolean {
  if (!transform) return true;
  const normalized = normalizeSpreadsheetImageTransform(transform);
  return (
    normalized.rotation === 0 &&
    !normalized.flipHorizontal &&
    !normalized.flipVertical
  );
}

export interface XlsxImageTransformReadResult {
  transform: WorkSpreadsheetImageTransform | null;
  /** False when rotation is present but not a supported quadrant. */
  supported: boolean;
}

export function readXlsxImageTransform(
  xfrm: Element | null | undefined,
): XlsxImageTransformReadResult {
  if (!xfrm) return { transform: null, supported: true };
  const rawRotation = attribute(xfrm, 'rot');
  const rawFlipHorizontal = attribute(xfrm, 'flipH');
  const rawFlipVertical = attribute(xfrm, 'flipV');
  const rotationUnits = parseRotationUnits(rawRotation ?? undefined);
  const flipHorizontal = parseFlip(rawFlipHorizontal ?? undefined);
  const flipVertical = parseFlip(rawFlipVertical ?? undefined);
  const supported =
    rotationUnits !== null &&
    flipHorizontal !== null &&
    flipVertical !== null &&
    (rotationUnits === 0 || rotationUnits % XLSX_QUADRANT_ROTATION_UNITS === 0);
  const normalized = normalizeSpreadsheetImageTransform({
    rotation:
      rotationUnits === null
        ? 0
        : rotationUnits / XLSX_ROTATION_UNITS_PER_DEGREE,
    flipHorizontal: flipHorizontal ?? false,
    flipVertical: flipVertical ?? false,
  });
  return {
    transform: isIdentitySpreadsheetImageTransform(normalized)
      ? null
      : normalized,
    supported,
  };
}

export function xlsxImageTransformAttributes(
  transform: WorkSpreadsheetImageTransform | null | undefined,
): string {
  if (isIdentitySpreadsheetImageTransform(transform)) return '';
  const normalized = normalizeSpreadsheetImageTransform(transform!);
  const attributes: string[] = [];
  if (normalized.rotation !== 0) {
    attributes.push(
      `rot="${normalized.rotation * XLSX_ROTATION_UNITS_PER_DEGREE}"`,
    );
  }
  if (normalized.flipHorizontal) attributes.push('flipH="1"');
  if (normalized.flipVertical) attributes.push('flipV="1"');
  return attributes.length ? ` ${attributes.join(' ')}` : '';
}

function parseRotationUnits(source: string | undefined): number | null {
  if (source == null || source.trim() === '') return 0;
  const value = Number(source);
  return Number.isFinite(value) ? Math.trunc(value) : null;
}

function parseFlip(source: string | undefined): boolean | null {
  if (source == null || source.trim() === '') return false;
  if (source === '1' || source.toLowerCase() === 'true') return true;
  if (source === '0' || source.toLowerCase() === 'false') return false;
  return null;
}
