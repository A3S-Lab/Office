export const DOCUMENT_CHARACTER_SPACING_ATTRIBUTE =
  'data-office-character-spacing-twips';
export const DOCUMENT_CHARACTER_SPACING_MAX_TWIPS = 31_680;
export const DOCUMENT_CHARACTER_SPACING_MIN_TWIPS =
  -DOCUMENT_CHARACTER_SPACING_MAX_TWIPS;

export function normalizeDocumentCharacterSpacingTwips(
  value: unknown,
): number | null {
  if (value === null || value === undefined || value === '') return null;
  const spacing = Number(value);
  if (
    !Number.isSafeInteger(spacing) ||
    spacing < DOCUMENT_CHARACTER_SPACING_MIN_TWIPS ||
    spacing > DOCUMENT_CHARACTER_SPACING_MAX_TWIPS
  ) {
    return null;
  }
  return Object.is(spacing, -0) ? 0 : spacing;
}

export function documentCharacterSpacingDomAttributes(
  value: unknown,
): Record<string, string> {
  const spacing = normalizeDocumentCharacterSpacingTwips(value);
  if (spacing === null) return {};
  return {
    [DOCUMENT_CHARACTER_SPACING_ATTRIBUTE]: String(spacing),
    style: `letter-spacing: ${formatCharacterSpacingPoints(spacing)}pt`,
  };
}

export function documentCharacterSpacingTwipsFromElement(
  element: HTMLElement,
): number | null {
  if (element.hasAttribute(DOCUMENT_CHARACTER_SPACING_ATTRIBUTE)) {
    return normalizeDocumentCharacterSpacingTwips(
      element.getAttribute(DOCUMENT_CHARACTER_SPACING_ATTRIBUTE),
    );
  }
  return documentCharacterSpacingTwipsFromCss(element.style.letterSpacing);
}

export function documentCharacterSpacingTwipsFromCss(
  value: unknown,
): number | null {
  if (typeof value !== 'string') return null;
  const source = value.trim().toLowerCase();
  if (!source || source === 'normal') return null;
  const match = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))(pt|px)?$/.exec(source);
  if (!match) return null;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount)) return null;
  const unit = match[2];
  if (!unit && amount !== 0) return null;
  const twips = Math.round(amount * (unit === 'px' ? 15 : 20));
  return normalizeDocumentCharacterSpacingTwips(twips);
}

export function documentCharacterSpacingPoints(value: unknown): number | null {
  const spacing = normalizeDocumentCharacterSpacingTwips(value);
  return spacing === null ? null : spacing / 20;
}

function formatCharacterSpacingPoints(value: number): string {
  return String(value / 20);
}
