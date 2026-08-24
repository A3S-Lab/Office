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
  documentEmphasisMarkDomAttributes,
  documentEmphasisMarkFromElement,
  normalizeDocumentEmphasisMark,
  type WorkDocumentEmphasisMark,
} from './work-document-emphasis';
import {
  documentKerningDomAttributes,
  documentKerningThresholdHalfPointsFromElement,
  normalizeDocumentKerningThresholdHalfPoints,
} from './work-document-kerning';
import {
  DOCUMENT_HIDDEN_TEXT_ATTRIBUTE,
  documentHiddenTextDomAttributes,
  documentHiddenTextFromElement,
} from './work-document-hidden-text';
import {
  DOCUMENT_LEGACY_TEXT_EMBOSS_ATTRIBUTE,
  DOCUMENT_LEGACY_TEXT_IMPRINT_ATTRIBUTE,
  DOCUMENT_LEGACY_TEXT_OUTLINE_ATTRIBUTE,
  DOCUMENT_LEGACY_TEXT_SHADOW_ATTRIBUTE,
  documentLegacyTextEffectFromElement,
  documentLegacyTextEffectsDomAttributes,
  documentLegacyTextStyleAttributeName,
  normalizeDocumentLegacyTextEffect,
  type WorkDocumentLegacyTextEffectName,
} from './work-document-legacy-text-effects';
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
import {
  DOCUMENT_RUN_BORDER_ATTRIBUTE,
  documentRunBorderDomAttributes,
  parseDocumentRunBorder,
  parseDocumentRunBorderElement,
  serializeDocumentRunBorder,
} from './work-document-run-border';
import {
  DOCUMENT_RUN_SHADING_ATTRIBUTE,
  documentRunShadingDomAttributes,
  parseDocumentRunShading,
  parseDocumentRunShadingElement,
  serializeDocumentRunShading,
} from './work-document-run-shading';
import {
  DOCUMENT_HIGHLIGHT_ATTRIBUTE,
  documentHighlightForCssColor,
  normalizeDocumentHighlight,
} from './work-document-highlight';
import {
  documentScriptFontsDomAttributes,
  documentScriptFontsForAllText,
  normalizeDocumentScriptFontSlot,
  parseDocumentScriptFonts,
  serializeDocumentScriptFonts,
  type WorkDocumentScriptFontSlot,
} from './work-document-script-fonts';

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
    emphasisMark?: WorkDocumentEmphasisMark | null;
    hiddenText?: boolean | null;
    legacyTextOutline?: boolean | null;
    legacyTextShadow?: boolean | null;
    legacyTextEmboss?: boolean | null;
    legacyTextImprint?: boolean | null;
    runBorder?: string | null;
    runShading?: string | null;
    kerningThresholdHalfPoints?: number | null;
    scriptFonts?: string | null;
    scriptFontSlot?: WorkDocumentScriptFontSlot | null;
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
  parseHTML() {
    return [
      ...(this.parent?.() ?? []),
      {
        tag: `span[${DOCUMENT_HIDDEN_TEXT_ATTRIBUTE}]`,
        consuming: false,
      },
      ...[
        DOCUMENT_LEGACY_TEXT_OUTLINE_ATTRIBUTE,
        DOCUMENT_LEGACY_TEXT_SHADOW_ATTRIBUTE,
        DOCUMENT_LEGACY_TEXT_EMBOSS_ATTRIBUTE,
        DOCUMENT_LEGACY_TEXT_IMPRINT_ATTRIBUTE,
      ].map((attribute) => ({
        tag: `span[${attribute}]`,
        consuming: false,
      })),
      {
        tag: `span[${DOCUMENT_RUN_BORDER_ATTRIBUTE}]`,
        consuming: false,
      },
      {
        tag: `span[${DOCUMENT_RUN_SHADING_ATTRIBUTE}]`,
        consuming: false,
      },
    ];
  },
  addCommands() {
    return {
      ...(this.parent?.() ?? {}),
      removeEmptyTextStyle:
        () =>
        ({ tr }) => {
          const { from, to } = tr.selection;
          tr.doc.nodesBetween(from, to, (node, position) => {
            if (!node.isText) return true;
            const mark = node.marks.find(
              (candidate) =>
                candidate.type === this.type &&
                documentTextStyleAttributesAreEmpty(candidate.attrs),
            );
            if (mark) {
              tr.removeMark(
                Math.max(position, from),
                Math.min(position + node.nodeSize, to),
                mark,
              );
            }
            return false;
          });
          return true;
        },
    };
  },
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
      emphasisMark: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          documentEmphasisMarkFromElement(element),
        renderHTML: (attributes: Record<string, unknown>) =>
          documentEmphasisMarkDomAttributes(
            normalizeDocumentEmphasisMark(attributes.emphasisMark),
          ),
      },
      hiddenText: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          documentHiddenTextFromElement(element),
        renderHTML: (attributes: Record<string, unknown>) =>
          documentHiddenTextDomAttributes(attributes.hiddenText),
      },
      legacyTextOutline: legacyTextEffectAttribute('outline'),
      legacyTextShadow: legacyTextEffectAttribute('shadow'),
      legacyTextEmboss: legacyTextEffectAttribute('emboss'),
      legacyTextImprint: legacyTextEffectAttribute('imprint'),
      runBorder: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          serializeDocumentRunBorder(parseDocumentRunBorderElement(element)) ??
          null,
        renderHTML: (attributes: Record<string, unknown>) =>
          documentRunBorderDomAttributes(
            parseDocumentRunBorder(attributes.runBorder),
          ),
      },
      runShading: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          serializeDocumentRunShading(
            parseDocumentRunShadingElement(element),
          ) ?? null,
        renderHTML: (attributes: Record<string, unknown>) =>
          documentRunShadingDomAttributes(
            parseDocumentRunShading(attributes.runShading),
          ),
      },
      kerningThresholdHalfPoints: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          documentKerningThresholdHalfPointsFromElement(element),
        renderHTML: (attributes: Record<string, unknown>) =>
          documentKerningDomAttributes(
            normalizeDocumentKerningThresholdHalfPoints(
              attributes.kerningThresholdHalfPoints,
            ),
            attributes.fontSize,
          ),
      },
      scriptFonts: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          serializeDocumentScriptFonts(
            parseDocumentScriptFonts(element.dataset.officeScriptFonts),
          ),
        renderHTML: (attributes: Record<string, unknown>) => {
          const domAttributes = documentScriptFontsDomAttributes(
            typeof attributes.scriptFonts === 'string'
              ? parseDocumentScriptFonts(attributes.scriptFonts)
              : attributes.scriptFonts,
            attributes.scriptFontSlot,
          );
          delete domAttributes.style;
          return domAttributes;
        },
      },
      scriptFontSlot: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          normalizeDocumentScriptFontSlot(element.dataset.officeScriptFontSlot),
        renderHTML: () => ({}),
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

function legacyTextEffectAttribute(effect: WorkDocumentLegacyTextEffectName) {
  return {
    default: null,
    parseHTML: (element: HTMLElement) =>
      documentLegacyTextEffectFromElement(element, effect),
    renderHTML: (attributes: Record<string, unknown>) => {
      const value = normalizeDocumentLegacyTextEffect(
        attributes[documentLegacyTextStyleAttributeName(effect)],
      );
      return value === null
        ? {}
        : documentLegacyTextEffectsDomAttributes({ [effect]: value });
    },
  };
}

function documentTextStyleAttributesAreEmpty(
  attributes: Record<string, unknown>,
): boolean {
  return Object.values(attributes).every(
    (value) => value === null || value === undefined || value === '',
  );
}

export const DocumentHighlight = Highlight.extend({
  addKeyboardShortcuts() {
    return {};
  },
  addAttributes() {
    const parent = (this.parent?.() ?? {}) as Record<
      string,
      Record<string, unknown>
    >;
    return {
      ...parent,
      color: {
        ...(parent.color ?? {}),
        renderHTML: (attributes: Record<string, unknown>) => {
          const color =
            typeof attributes.color === 'string' ? attributes.color : null;
          if (!color) return {};
          const native =
            normalizeDocumentHighlight(attributes.nativeHighlight) ??
            documentHighlightForCssColor(color);
          return {
            'data-color': color,
            ...(native ? { [DOCUMENT_HIGHLIGHT_ATTRIBUTE]: native } : {}),
            style: `background-color: ${color}; color: inherit`,
          };
        },
      },
      nativeHighlight: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          normalizeDocumentHighlight(
            element.getAttribute(DOCUMENT_HIGHLIGHT_ATTRIBUTE),
          ),
        renderHTML: (attributes: Record<string, unknown>) => {
          const value = normalizeDocumentHighlight(attributes.nativeHighlight);
          return value ? { [DOCUMENT_HIGHLIGHT_ATTRIBUTE]: value } : {};
        },
      },
      themeFill: themeReferenceAttribute('officeThemeFill'),
    };
  },
  parseHTML() {
    const excludesRunShading = (element: HTMLElement) =>
      element.hasAttribute(DOCUMENT_RUN_SHADING_ATTRIBUTE) &&
      !element.hasAttribute(DOCUMENT_HIGHLIGHT_ATTRIBUTE)
        ? false
        : null;
    return [
      ...(this.parent?.() ?? []),
      { tag: `span[${DOCUMENT_HIGHLIGHT_ATTRIBUTE}]` },
      { tag: 'span[data-office-theme-fill]', getAttrs: excludesRunShading },
      {
        tag: 'span[style*="background-color"]',
        getAttrs: excludesRunShading,
      },
    ];
  },
});

export const DocumentFontFamily = FontFamily.extend({
  addCommands() {
    return {
      setFontFamily:
        (fontFamily: string) =>
        ({ chain }) => {
          const scriptFonts = serializeDocumentScriptFonts(
            documentScriptFontsForAllText(fontFamily),
          );
          return chain()
            .setMark('textStyle', {
              fontFamily,
              scriptFonts,
              scriptFontSlot: null,
              wordLineHeightFactor: documentWordLineHeightFactor(fontFamily),
            })
            .run();
        },
      unsetFontFamily:
        () =>
        ({ chain }) =>
          chain()
            .setMark('textStyle', {
              fontFamily: null,
              scriptFonts: null,
              scriptFontSlot: null,
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
