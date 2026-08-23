export const DOCUMENT_CHARACTER_POSITION_ATTRIBUTE =
  'data-office-character-position-half-points';
export const DOCUMENT_CHARACTER_POSITION_CSS_PROPERTY =
  '--work-document-character-position';
export const DOCUMENT_CHARACTER_POSITION_MAX_HALF_POINTS = 3_168;
export const DOCUMENT_CHARACTER_POSITION_MIN_HALF_POINTS =
  -DOCUMENT_CHARACTER_POSITION_MAX_HALF_POINTS;

export function normalizeDocumentCharacterPositionHalfPoints(
  value: unknown,
): number | null {
  if (value === null || value === undefined || value === '') return null;
  const position = Number(value);
  if (
    !Number.isSafeInteger(position) ||
    position < DOCUMENT_CHARACTER_POSITION_MIN_HALF_POINTS ||
    position > DOCUMENT_CHARACTER_POSITION_MAX_HALF_POINTS
  ) {
    return null;
  }
  return Object.is(position, -0) ? 0 : position;
}

export function documentCharacterPositionDomAttributes(
  value: unknown,
): Record<string, string> {
  const position = normalizeDocumentCharacterPositionHalfPoints(value);
  if (position === null) return {};
  return {
    [DOCUMENT_CHARACTER_POSITION_ATTRIBUTE]: String(position),
    style: `${DOCUMENT_CHARACTER_POSITION_CSS_PROPERTY}: ${formatCharacterPositionPoints(position)}pt`,
  };
}

export function documentCharacterPositionHalfPointsFromElement(
  element: HTMLElement,
): number | null {
  if (element.hasAttribute(DOCUMENT_CHARACTER_POSITION_ATTRIBUTE)) {
    return normalizeDocumentCharacterPositionHalfPoints(
      element.getAttribute(DOCUMENT_CHARACTER_POSITION_ATTRIBUTE),
    );
  }
  return (
    documentCharacterPositionHalfPointsFromCss(
      element.style.getPropertyValue(DOCUMENT_CHARACTER_POSITION_CSS_PROPERTY),
    ) ?? documentCharacterPositionHalfPointsFromCss(element.style.verticalAlign)
  );
}

export function documentCharacterPositionHalfPointsFromCss(
  value: unknown,
): number | null {
  if (typeof value !== 'string') return null;
  const source = value.trim().toLowerCase();
  if (!source || source === 'baseline') return null;
  const match = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))(pt|px)?$/.exec(source);
  if (!match) return null;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount)) return null;
  const unit = match[2];
  if (!unit && amount !== 0) return null;
  const halfPoints = Math.round(amount * (unit === 'px' ? 1.5 : 2));
  return normalizeDocumentCharacterPositionHalfPoints(halfPoints);
}

export function documentCharacterPositionPoints(value: unknown): number | null {
  const position = normalizeDocumentCharacterPositionHalfPoints(value);
  return position === null ? null : position / 2;
}

function formatCharacterPositionPoints(value: number): string {
  return String(value / 2);
}
