import { normalizeCssColor } from './work-css-color';

export const DOCUMENT_HIGHLIGHT_ATTRIBUTE = 'data-office-highlight';

export type WorkDocumentHighlight =
  | 'black'
  | 'blue'
  | 'cyan'
  | 'darkBlue'
  | 'darkCyan'
  | 'darkGray'
  | 'darkGreen'
  | 'darkMagenta'
  | 'darkRed'
  | 'darkYellow'
  | 'green'
  | 'lightGray'
  | 'magenta'
  | 'none'
  | 'red'
  | 'white'
  | 'yellow';

export const DOCUMENT_HIGHLIGHT_VALUES = new Set<WorkDocumentHighlight>([
  'black',
  'blue',
  'cyan',
  'darkBlue',
  'darkCyan',
  'darkGray',
  'darkGreen',
  'darkMagenta',
  'darkRed',
  'darkYellow',
  'green',
  'lightGray',
  'magenta',
  'none',
  'red',
  'white',
  'yellow',
]);

const HIGHLIGHT_COLORS: Readonly<Record<WorkDocumentHighlight, string>> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00ffff',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGray: '#808080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  green: '#00ff00',
  lightGray: '#c0c0c0',
  magenta: '#ff00ff',
  none: 'transparent',
  red: '#ff0000',
  white: '#ffffff',
  yellow: '#ffff00',
};

export function normalizeDocumentHighlight(
  source: unknown,
): WorkDocumentHighlight | null {
  return typeof source === 'string' &&
    DOCUMENT_HIGHLIGHT_VALUES.has(source as WorkDocumentHighlight)
    ? (source as WorkDocumentHighlight)
    : null;
}

export function documentHighlightFromDocxValue(
  source: unknown,
): WorkDocumentHighlight | null {
  if (typeof source !== 'string') return null;
  const normalized = source.trim().toLowerCase();
  for (const value of DOCUMENT_HIGHLIGHT_VALUES) {
    if (value.toLowerCase() === normalized) return value;
  }
  return null;
}

export function documentHighlightCssColor(source: unknown): string | null {
  const value = normalizeDocumentHighlight(source);
  return value ? HIGHLIGHT_COLORS[value] : null;
}

export function documentHighlightForCssColor(
  source: unknown,
): WorkDocumentHighlight | null {
  const color = normalizeCssColor(typeof source === 'string' ? source : null);
  if (!color) return null;
  for (const [value, candidate] of Object.entries(HIGHLIGHT_COLORS)) {
    if (candidate === color) return value as WorkDocumentHighlight;
  }
  return null;
}

export function documentHighlightFromElement(
  element: HTMLElement,
): WorkDocumentHighlight | null {
  return (
    normalizeDocumentHighlight(
      element.getAttribute(DOCUMENT_HIGHLIGHT_ATTRIBUTE),
    ) ?? documentHighlightForCssColor(element.style.backgroundColor)
  );
}

export function documentHighlightDomAttributes(
  source: unknown,
): Record<string, string> {
  const value = normalizeDocumentHighlight(source);
  const color = documentHighlightCssColor(value);
  return value && color
    ? {
        [DOCUMENT_HIGHLIGHT_ATTRIBUTE]: value,
        style: `background-color: ${color}`,
      }
    : {};
}
