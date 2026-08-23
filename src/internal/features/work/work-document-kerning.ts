export const DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE =
  'data-office-kerning-threshold-half-points';
export const DOCUMENT_KERNING_THRESHOLD_MIN_HALF_POINTS = 0;
export const DOCUMENT_KERNING_THRESHOLD_MAX_HALF_POINTS = 3_277;
export const DOCUMENT_KERNING_DEFAULT_THRESHOLD_HALF_POINTS = 24;

const DOCUMENT_DEFAULT_FONT_SIZE_POINTS = 10.5;

export function normalizeDocumentKerningThresholdHalfPoints(
  value: unknown,
): number | null {
  const threshold =
    typeof value === 'number'
      ? value
      : typeof value === 'string' && /^\+?\d+$/u.test(value.trim())
        ? Number(value.trim())
        : Number.NaN;
  if (
    !Number.isSafeInteger(threshold) ||
    threshold < DOCUMENT_KERNING_THRESHOLD_MIN_HALF_POINTS ||
    threshold > DOCUMENT_KERNING_THRESHOLD_MAX_HALF_POINTS
  ) {
    return null;
  }
  return Object.is(threshold, -0) ? 0 : threshold;
}

export function documentKerningThresholdPoints(value: unknown): number | null {
  const threshold = normalizeDocumentKerningThresholdHalfPoints(value);
  return threshold === null ? null : threshold / 2;
}

export function documentKerningIsEffective(
  thresholdValue: unknown,
  fontSizeValue: unknown,
): boolean {
  const threshold = normalizeDocumentKerningThresholdHalfPoints(thresholdValue);
  if (threshold === null) return false;
  if (threshold === 0) return true;
  return documentKerningFontSizePoints(fontSizeValue) * 2 >= threshold;
}

export function documentKerningDomAttributes(
  thresholdValue: unknown,
  fontSizeValue?: unknown,
): Record<string, string> {
  const threshold = normalizeDocumentKerningThresholdHalfPoints(thresholdValue);
  if (threshold === null) return {};
  return {
    [DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE]: String(threshold),
    style: `font-kerning: ${documentKerningIsEffective(threshold, fontSizeValue) ? 'normal' : 'none'}`,
  };
}

export function documentKerningThresholdHalfPointsFromElement(
  element: HTMLElement,
): number | null {
  return element.hasAttribute(DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE)
    ? normalizeDocumentKerningThresholdHalfPoints(
        element.getAttribute(DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE),
      )
    : null;
}

export function documentKerningFontSizePoints(value: unknown): number {
  if (typeof value === 'number') {
    return Number.isFinite(value) && value > 0
      ? value
      : DOCUMENT_DEFAULT_FONT_SIZE_POINTS;
  }
  if (typeof value !== 'string') return DOCUMENT_DEFAULT_FONT_SIZE_POINTS;
  const match = /^(\d+(?:\.\d*)?|\.\d+)(pt|px)?$/iu.exec(value.trim());
  if (!match) return DOCUMENT_DEFAULT_FONT_SIZE_POINTS;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount) || amount <= 0) {
    return DOCUMENT_DEFAULT_FONT_SIZE_POINTS;
  }
  return match[2]?.toLowerCase() === 'px' ? amount * 0.75 : amount;
}
