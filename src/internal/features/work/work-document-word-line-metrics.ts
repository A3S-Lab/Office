import FontFamily from '@tiptap/extension-text-style/font-family';
import { TextStyle } from '@tiptap/extension-text-style';
import Highlight from '@tiptap/extension-highlight';
import {
  documentCharacterScaleDomAttributes,
  documentCharacterScalePercentFromElement,
  normalizeDocumentCharacterScalePercent,
} from './work-document-character-scale';
import {
  documentCharacterSpacingDomAttributes,
  documentCharacterSpacingTwipsFromElement,
  normalizeDocumentCharacterSpacingTwips,
} from './work-document-character-spacing';
import {
  documentCharacterPositionDomAttributes,
  documentCharacterPositionHalfPointsFromElement,
  normalizeDocumentCharacterPositionHalfPoints,
} from './work-document-character-position';
import {
  parseDocxThemeReference,
  serializeDocxThemeReference,
} from './work-docx-theme-reference';
import {
  DOCUMENT_TEXT_CASE_ATTRIBUTE,
  documentTextCaseCss,
  normalizeDocumentTextCase,
  type WorkDocumentTextCase,
} from './work-document-text-case';

export const DOCUMENT_WORD_DEFAULT_SINGLE_LINE_HEIGHT = 1.15;

// These factors are WPS 12.1 single-line advances divided by the requested
// font size. The Latin values agree with the fonts' OpenType horizontal
// metrics. WPS applies larger compatibility advances to legacy Chinese fonts,
// so those values are calibrated from the checked-in WPS layout matrix.
const DOCUMENT_WORD_LINE_HEIGHT_FACTORS = new Map<string, number>([
  ['arial', 1.15],
  ['times new roman', 1.15],
  ['calibri', 1.2207],
  ['segoe ui', 1.3301],
  ['microsoft yahei', 1.7143],
  ['microsoft yahei ui', 1.7143],
  ['simsun', 1.2976],
  ['simhei', 1.2976],
  ['fangsong', 1.2976],
  ['kaiti', 1.2976],
  ['dengxian', 1.3548],
]);

declare module '@tiptap/extension-text-style' {
  interface TextStyleAttributes {
    characterScalePercent?: number | null;
    characterPositionHalfPoints?: number | null;
    characterSpacingTwips?: number | null;
    wordLineHeightFactor?: number | null;
    wordSnapToGrid?: boolean | null;
    themeColor?: string | null;
    textCase?: WorkDocumentTextCase | null;
  }
}

export function documentWordLineHeightFactor(fontFamily: unknown): number {
  if (typeof fontFamily !== 'string') {
    return DOCUMENT_WORD_DEFAULT_SINGLE_LINE_HEIGHT;
  }
  const family = firstCssFontFamily(fontFamily).toLocaleLowerCase();
  return (
    DOCUMENT_WORD_LINE_HEIGHT_FACTORS.get(family) ??
    DOCUMENT_WORD_DEFAULT_SINGLE_LINE_HEIGHT
  );
}

export function normalizedDocumentWordLineHeightFactor(
  value: unknown,
): number | null {
  if (value === null || value === undefined || value === '') return null;
  const factor = Number(value);
  if (!Number.isFinite(factor) || factor < 0.5 || factor > 4) return null;
  return Number(factor.toFixed(4));
}

export const DocumentTextStyle = TextStyle.extend({
  addAttributes() {
    return {
      characterScalePercent: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          documentCharacterScalePercentFromElement(element),
        renderHTML: (attributes: Record<string, unknown>) =>
          documentCharacterScaleDomAttributes(
            normalizeDocumentCharacterScalePercent(
              attributes.characterScalePercent,
            ),
          ),
      },
      characterPositionHalfPoints: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          documentCharacterPositionHalfPointsFromElement(element),
        renderHTML: (attributes: Record<string, unknown>) =>
          documentCharacterPositionDomAttributes(
            normalizeDocumentCharacterPositionHalfPoints(
              attributes.characterPositionHalfPoints,
            ),
          ),
      },
      characterSpacingTwips: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          documentCharacterSpacingTwipsFromElement(element),
        renderHTML: (attributes: Record<string, unknown>) =>
          documentCharacterSpacingDomAttributes(
            normalizeDocumentCharacterSpacingTwips(
              attributes.characterSpacingTwips,
            ),
          ),
      },
      wordLineHeightFactor: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          normalizedDocumentWordLineHeightFactor(
            element.dataset.officeWordLineHeightFactor ??
              element.style.getPropertyValue(
                '--work-document-word-line-height-factor',
              ),
          ),
        renderHTML: (attributes: Record<string, unknown>) => {
          const factor = normalizedDocumentWordLineHeightFactor(
            attributes.wordLineHeightFactor,
          );
          if (factor === null) return {};
          return {
            'data-office-word-line-height-factor': String(factor),
            style: `--work-document-word-line-height-factor: ${factor}`,
          };
        },
      },
      wordSnapToGrid: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          normalizedBoolean(element.dataset.officeWordSnapToGrid),
        renderHTML: (attributes: Record<string, unknown>) => {
          const value = normalizedBoolean(attributes.wordSnapToGrid);
          return value === null
            ? {}
            : { 'data-office-word-snap-to-grid': String(value) };
        },
      },
      themeColor: themeReferenceAttribute('officeThemeColor'),
      textCase: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          normalizeDocumentTextCase(element.dataset.officeTextCase),
        renderHTML: (attributes: Record<string, unknown>) => {
          const textCase = normalizeDocumentTextCase(attributes.textCase);
          return textCase
            ? {
                [DOCUMENT_TEXT_CASE_ATTRIBUTE]: textCase,
                style: documentTextCaseCss(textCase),
              }
            : {};
        },
      },
    };
  },
});

export const DocumentHighlight = Highlight.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      themeFill: themeReferenceAttribute('officeThemeFill'),
    };
  },
  parseHTML() {
    return [
      ...(this.parent?.() ?? []),
      { tag: 'span[data-office-theme-fill]' },
      { tag: 'span[style*="background-color"]' },
    ];
  },
});

export const DocumentFontFamily = FontFamily.extend({
  addCommands() {
    return {
      setFontFamily:
        (fontFamily: string) =>
        ({ chain }) =>
          chain()
            .setMark('textStyle', {
              fontFamily,
              wordLineHeightFactor: documentWordLineHeightFactor(fontFamily),
            })
            .run(),
      unsetFontFamily:
        () =>
        ({ chain }) =>
          chain()
            .setMark('textStyle', {
              fontFamily: null,
              wordLineHeightFactor: null,
            })
            .removeEmptyTextStyle()
            .run(),
    };
  },
});

function firstCssFontFamily(value: string): string {
  const source = value.trim();
  if (!source) return '';
  const quote = source[0];
  if (quote === '"' || quote === "'") {
    const end = source.indexOf(quote, 1);
    if (end > 0) return source.slice(1, end).trim();
  }
  return (source.split(',')[0] ?? source)
    .trim()
    .replace(/^(['"])(.*)\1$/, '$2');
}

function normalizedBoolean(value: unknown): boolean | null {
  if (value === true || value === 'true' || value === '1') return true;
  if (value === false || value === 'false' || value === '0') return false;
  return null;
}

function themeReferenceAttribute(datasetKey: string) {
  return {
    default: null,
    parseHTML: (element: HTMLElement) =>
      serializeDocxThemeReference(
        parseDocxThemeReference(element.dataset[datasetKey]),
      ) ?? null,
    renderHTML: (attributes: Record<string, unknown>) => {
      const value = serializeDocxThemeReference(
        parseDocxThemeReference(
          typeof attributes[
            datasetKey === 'officeThemeColor' ? 'themeColor' : 'themeFill'
          ] === 'string'
            ? String(
                attributes[
                  datasetKey === 'officeThemeColor' ? 'themeColor' : 'themeFill'
                ],
              )
            : null,
        ),
      );
      return value ? { [`data-${toKebabCase(datasetKey)}`]: value } : {};
    },
  };
}

function toKebabCase(value: string): string {
  return value.replace(/[A-Z]/g, (character) => `-${character.toLowerCase()}`);
}
