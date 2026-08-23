export const DOCUMENT_CHARACTER_SCALE_ATTRIBUTE =
  'data-office-character-scale-percent';
export const DOCUMENT_CHARACTER_SCALE_DEFAULT_PERCENT = 100;
export const DOCUMENT_CHARACTER_SCALE_MIN_PERCENT = 1;
export const DOCUMENT_CHARACTER_SCALE_MAX_PERCENT = 600;

export function normalizeDocumentCharacterScalePercent(
  value: unknown,
): number | null {
  if (value === null || value === undefined || value === '') return null;
  const scale = Number(value);
  return Number.isSafeInteger(scale) &&
    scale >= DOCUMENT_CHARACTER_SCALE_MIN_PERCENT &&
    scale <= DOCUMENT_CHARACTER_SCALE_MAX_PERCENT
    ? scale
    : null;
}

export function documentCharacterScaleDomAttributes(
  value: unknown,
): Record<string, string> {
  const scale = normalizeDocumentCharacterScalePercent(value);
  if (scale === null) return {};
  return {
    [DOCUMENT_CHARACTER_SCALE_ATTRIBUTE]: String(scale),
    style: `font-stretch: ${scale}%`,
  };
}

export function documentCharacterScalePercentFromElement(
  element: HTMLElement,
): number | null {
  if (element.hasAttribute(DOCUMENT_CHARACTER_SCALE_ATTRIBUTE)) {
    return normalizeDocumentCharacterScalePercent(
      element.getAttribute(DOCUMENT_CHARACTER_SCALE_ATTRIBUTE),
    );
  }
  return documentCharacterScalePercentFromCss(element.style.fontStretch);
}

export function documentCharacterScalePercentFromCss(
  value: unknown,
): number | null {
  if (typeof value !== 'string') return null;
  const source = value.trim().toLowerCase();
  if (!source || source === 'normal') return null;
  const match = /^(\d+(?:\.\d*)?|\.\d+)%$/u.exec(source);
  if (!match) return null;
  return normalizeDocumentCharacterScalePercent(Number(match[1]));
}
