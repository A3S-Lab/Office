export const DOCUMENT_EMPHASIS_MARK_ATTRIBUTE = 'data-office-emphasis-mark';

export const DOCUMENT_EMPHASIS_MARKS = [
  'none',
  'dot',
  'comma',
  'circle',
  'underDot',
] as const;

export type WorkDocumentEmphasisMark = (typeof DOCUMENT_EMPHASIS_MARKS)[number];

const DOCUMENT_EMPHASIS_MARK_SET = new Set<WorkDocumentEmphasisMark>(
  DOCUMENT_EMPHASIS_MARKS,
);

export function normalizeDocumentEmphasisMark(
  value: unknown,
): WorkDocumentEmphasisMark | null {
  return typeof value === 'string' &&
    DOCUMENT_EMPHASIS_MARK_SET.has(value as WorkDocumentEmphasisMark)
    ? (value as WorkDocumentEmphasisMark)
    : null;
}

export function documentEmphasisMarkCss(value: unknown): string {
  const emphasisMark = normalizeDocumentEmphasisMark(value);
  if (!emphasisMark) return '';
  if (emphasisMark === 'none') {
    return 'text-emphasis-style:none;-webkit-text-emphasis-style:none';
  }
  const style =
    emphasisMark === 'comma'
      ? '","'
      : emphasisMark === 'circle'
        ? 'open circle'
        : 'filled dot';
  const position = emphasisMark === 'underDot' ? 'under right' : 'over right';
  return `text-emphasis-style:${style};text-emphasis-position:${position};-webkit-text-emphasis-style:${style};-webkit-text-emphasis-position:${position}`;
}

export function documentEmphasisMarkDomAttributes(
  value: unknown,
): Record<string, string> {
  const emphasisMark = normalizeDocumentEmphasisMark(value);
  if (!emphasisMark) return {};
  return {
    [DOCUMENT_EMPHASIS_MARK_ATTRIBUTE]: emphasisMark,
    style: documentEmphasisMarkCss(emphasisMark),
  };
}

export function documentEmphasisMarkFromElement(
  element: HTMLElement,
): WorkDocumentEmphasisMark | null {
  return element.hasAttribute(DOCUMENT_EMPHASIS_MARK_ATTRIBUTE)
    ? normalizeDocumentEmphasisMark(
        element.getAttribute(DOCUMENT_EMPHASIS_MARK_ATTRIBUTE),
      )
    : null;
}
