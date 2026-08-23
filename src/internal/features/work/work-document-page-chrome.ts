import {
  createDocumentEquationElement,
  documentEquationFromElement,
} from './work-document-equations';
import {
  DOCUMENT_CHARACTER_SCALE_ATTRIBUTE,
  documentCharacterScaleDomAttributes,
  documentCharacterScalePercentFromElement,
} from './work-document-character-scale';
import {
  DOCUMENT_CHARACTER_POSITION_ATTRIBUTE,
  documentCharacterPositionDomAttributes,
  documentCharacterPositionHalfPointsFromElement,
} from './work-document-character-position';
import {
  DOCUMENT_CHARACTER_SPACING_ATTRIBUTE,
  documentCharacterSpacingDomAttributes,
  documentCharacterSpacingTwipsFromElement,
} from './work-document-character-spacing';
import {
  DOCUMENT_EMPHASIS_MARK_ATTRIBUTE,
  documentEmphasisMarkDomAttributes,
  documentEmphasisMarkFromElement,
} from './work-document-emphasis';
import {
  DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE,
  documentKerningDomAttributes,
  documentKerningThresholdHalfPointsFromElement,
} from './work-document-kerning';
import {
  DOCUMENT_HIDDEN_TEXT_ATTRIBUTE,
  documentHiddenTextFromElement,
} from './work-document-hidden-text';
import { normalizeDocumentImageIdentity } from './work-document-image-identity';
import {
  DOCUMENT_PARAGRAPH_BORDERS_ATTRIBUTE,
  documentParagraphBordersDomAttributes,
  parseDocumentParagraphBordersElement,
} from './work-document-paragraph-borders';
import {
  DOCUMENT_PARAGRAPH_ID_ATTRIBUTE,
  DOCUMENT_PARAGRAPH_TEXT_ID_ATTRIBUTE,
  normalizeDocumentParagraphIdentity,
} from './work-document-paragraph-identity';
import {
  documentParagraphShadingDomAttributes,
  parseDocumentParagraphShadingElement,
} from './work-document-paragraph-shading';
import {
  DOCUMENT_TABLE_ROW_ID_ATTRIBUTE,
  DOCUMENT_TABLE_ROW_TEXT_ID_ATTRIBUTE,
  normalizeDocumentTableRowIdentity,
} from './work-document-table-row-identity';
import {
  DOCUMENT_TEXT_CASE_ATTRIBUTE,
  documentTextCaseCss,
  normalizeDocumentTextCase,
} from './work-document-text-case';
import {
  cssDocumentFontFamily,
  DOCUMENT_SCRIPT_FONTS_ATTRIBUTE,
  DOCUMENT_SCRIPT_FONT_SLOT_ATTRIBUTE,
  documentFontNameFromCssFamily,
  documentScriptFontFamilyForRendering,
  documentScriptFontsFromElement,
  documentScriptFontsDomAttributes,
  documentScriptFontSlotFromElement,
  documentScriptFontSlotFromHint,
} from './work-document-script-fonts';
import {
  DOCUMENT_UNDERLINE_COLOR_ATTRIBUTE,
  DOCUMENT_UNDERLINE_STYLE_ATTRIBUTE,
  DOCUMENT_UNDERLINE_THEME_COLOR_ATTRIBUTE,
  documentUnderlineDomAttributes,
  documentUnderlineFormattingFromElement,
} from './work-document-underline';
import {
  DOCUMENT_STRIKE_STYLE_ATTRIBUTE,
  documentStrikeDomAttributes,
  documentStrikeFormattingFromElement,
} from './work-document-strike';
import type {
  WorkDocumentPageChrome,
  WorkDocumentPageChromeContent,
  WorkDocumentPageChromeVariant,
  WorkDocumentSectionLayout,
} from './work-types';

interface LegacyPageChrome {
  headerText?: string;
  footerText?: string;
  showPageNumbers?: boolean;
}

export interface ResolvedDocumentPageChrome
  extends WorkDocumentPageChromeContent {
  variant: WorkDocumentPageChromeVariant;
}

const EMPTY_CONTENT: WorkDocumentPageChromeContent = {
  headerHtml: '',
  footerHtml: '',
  showPageNumber: false,
};
const IMAGE_IDENTITY_ATTRIBUTES = [
  'data-office-image-object-id',
  'data-office-image-doc-properties-id',
  'data-office-image-anchor-id',
  'data-office-image-edit-id',
] as const;
const PARAGRAPH_IDENTITY_ATTRIBUTES = [
  DOCUMENT_PARAGRAPH_ID_ATTRIBUTE,
  DOCUMENT_PARAGRAPH_TEXT_ID_ATTRIBUTE,
] as const;
const PARAGRAPH_DEFAULT_COLLAPSED_ATTRIBUTE = 'data-office-default-collapsed';
const PARAGRAPH_BORDERS_ATTRIBUTE = DOCUMENT_PARAGRAPH_BORDERS_ATTRIBUTE;
const PARAGRAPH_SHADING_ATTRIBUTE = 'data-office-paragraph-shading';
const TABLE_ROW_IDENTITY_ATTRIBUTES = [
  DOCUMENT_TABLE_ROW_ID_ATTRIBUTE,
  DOCUMENT_TABLE_ROW_TEXT_ID_ATTRIBUTE,
] as const;
const EQUATION_ATTRIBUTES = [
  'aria-label',
  'class',
  'contenteditable',
  'data-document-equation',
  'data-equation-display',
  'data-equation-model',
  'role',
] as const;
const MATHML_ATTRIBUTES = new Set([
  'accent',
  'accentunder',
  'align',
  'bevelled',
  'columnalign',
  'depth',
  'dir',
  'display',
  'fence',
  'height',
  'lang',
  'linethickness',
  'mathbackground',
  'mathcolor',
  'mathsize',
  'mathvariant',
  'notation',
  'rowspacing',
  'scriptlevel',
  'separator',
  'style',
  'width',
  'xmlns',
]);
const MATHML_NAMESPACE = 'http://www.w3.org/1998/Math/MathML';
const MATHML_TAGS = new Set([
  'math',
  'maligngroup',
  'malignmark',
  'menclose',
  'mfrac',
  'mmultiscripts',
  'mo',
  'mover',
  'mpadded',
  'mphantom',
  'mprescripts',
  'mroot',
  'mrow',
  'msqrt',
  'msub',
  'msubsup',
  'msup',
  'mstyle',
  'mtable',
  'mtd',
  'mtext',
  'mtr',
  'munder',
  'munderover',
  'none',
]);

const ALLOWED_TAGS = new Set([
  'a',
  'b',
  'blockquote',
  'br',
  'div',
  'em',
  'i',
  'img',
  'li',
  'math',
  'maligngroup',
  'malignmark',
  'menclose',
  'mfrac',
  'mmultiscripts',
  'mo',
  'mover',
  'mpadded',
  'mphantom',
  'mprescripts',
  'mroot',
  'mrow',
  'msqrt',
  'msub',
  'msubsup',
  'msup',
  'mstyle',
  'mtable',
  'mtd',
  'mtext',
  'mtr',
  'munder',
  'munderover',
  'none',
  'ol',
  'p',
  's',
  'span',
  'strike',
  'strong',
  'sub',
  'sup',
  'table',
  'tbody',
  'td',
  'th',
  'thead',
  'tr',
  'u',
  'ul',
]);

export function normalizeDocumentPageChrome(
  source?: Partial<WorkDocumentPageChrome> | null,
  legacy: LegacyPageChrome = {},
): WorkDocumentPageChrome {
  const defaultContent = normalizePageChromeContent(source?.default, {
    headerHtml: plainTextToPageChromeHtml(legacy.headerText),
    footerHtml: plainTextToPageChromeHtml(legacy.footerText),
    showPageNumber: Boolean(legacy.showPageNumbers),
  });
  return {
    differentFirstPage: Boolean(source?.differentFirstPage),
    differentOddEvenPages: Boolean(source?.differentOddEvenPages),
    default: defaultContent,
    first: normalizePageChromeContent(source?.first, EMPTY_CONTENT),
    even: normalizePageChromeContent(source?.even, EMPTY_CONTENT),
  };
}

export function serializeDocumentPageChrome(
  chrome: WorkDocumentPageChrome,
): string {
  return JSON.stringify(normalizeDocumentPageChrome(chrome));
}

export function parseDocumentPageChrome(
  source: string | undefined,
  legacy: LegacyPageChrome = {},
  fallback?: WorkDocumentPageChrome,
): WorkDocumentPageChrome {
  if (source?.trim()) {
    try {
      const parsed = JSON.parse(source) as Partial<WorkDocumentPageChrome>;
      return normalizeDocumentPageChrome(parsed, legacy);
    } catch {
      // Fall through to the compatible legacy representation.
    }
  }
  if (
    legacy.headerText ||
    legacy.footerText ||
    legacy.showPageNumbers !== undefined
  ) {
    return normalizeDocumentPageChrome(undefined, legacy);
  }
  return normalizeDocumentPageChrome(fallback);
}

export function resolveDocumentPageChrome(
  layout: WorkDocumentSectionLayout,
  sectionPage: number,
  physicalPage: number,
): ResolvedDocumentPageChrome {
  const chrome = normalizeDocumentPageChrome(layout.pageChrome, layout);
  const variant: WorkDocumentPageChromeVariant =
    chrome.differentFirstPage && sectionPage === 1
      ? 'first'
      : chrome.differentOddEvenPages && physicalPage % 2 === 0
        ? 'even'
        : 'default';
  return { variant, ...chrome[variant] };
}

export function documentPageChromeLegacyFields(
  chrome: WorkDocumentPageChrome,
): LegacyPageChrome {
  const normalized = normalizeDocumentPageChrome(chrome);
  return {
    headerText: pageChromePlainText(normalized.default.headerHtml) || undefined,
    footerText: pageChromePlainText(normalized.default.footerHtml) || undefined,
    showPageNumbers: normalized.default.showPageNumber,
  };
}

export function updateDocumentPageChromeVariant(
  chrome: WorkDocumentPageChrome,
  variant: WorkDocumentPageChromeVariant,
  patch: Partial<WorkDocumentPageChromeContent>,
): WorkDocumentPageChrome {
  const normalized = normalizeDocumentPageChrome(chrome);
  return {
    ...normalized,
    [variant]: normalizePageChromeContent(
      { ...normalized[variant], ...patch },
      EMPTY_CONTENT,
    ),
  };
}

export function pageChromePlainText(source: string): string {
  if (!source.trim()) return '';
  const document = new DOMParser().parseFromString(source, 'text/html');
  const blocks = Array.from(document.body.querySelectorAll('p, div, li'))
    .map((element) => element.textContent?.replace(/\s+/g, ' ').trim() ?? '')
    .filter(Boolean);
  return (
    blocks.length ? blocks.join('\n') : (document.body.textContent ?? '')
  ).trim();
}

export function sanitizeDocumentPageChromeHtml(
  source: string | undefined,
): string {
  if (!source?.trim()) return '';
  const document = new DOMParser().parseFromString(source, 'text/html');
  for (const element of Array.from(
    document.body.querySelectorAll('script, iframe, object, embed, link, meta'),
  )) {
    element.remove();
  }
  for (const font of Array.from(
    document.body.querySelectorAll<HTMLElement>('font'),
  )) {
    const span = document.createElement('span');
    const color = font.getAttribute('color')?.trim();
    if (font.getAttribute('style'))
      span.setAttribute('style', font.getAttribute('style') ?? '');
    if (color && !span.style.color) span.style.color = color;
    span.append(...Array.from(font.childNodes));
    font.replaceWith(span);
  }
  for (const element of Array.from(
    document.body.querySelectorAll<HTMLElement>('span[data-document-equation]'),
  )) {
    const equation = documentEquationFromElement(element);
    if (!equation) {
      element.replaceWith(
        document.createTextNode(element.textContent?.trim() || '[Equation]'),
      );
      continue;
    }
    element.replaceWith(createDocumentEquationElement(document, equation));
  }
  for (const element of Array.from(
    document.body.querySelectorAll<HTMLElement>('*'),
  )) {
    const tag = element.tagName.toLowerCase();
    if (!ALLOWED_TAGS.has(tag)) {
      element.replaceWith(...Array.from(element.childNodes));
      continue;
    }
    if (MATHML_TAGS.has(tag)) {
      const equation = element.closest('span[data-document-equation="true"]');
      if (element.namespaceURI !== MATHML_NAMESPACE || !equation) {
        element.replaceWith(...Array.from(element.childNodes));
        continue;
      }
    }
    sanitizeAttributes(element, tag);
  }
  if (
    !(document.body.textContent ?? '').trim() &&
    !document.body.querySelector('img, table')
  )
    return '';
  return document.body.innerHTML;
}

export function plainTextToPageChromeHtml(source: string | undefined): string {
  const lines = source
    ?.split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  if (!lines?.length) return '';
  return lines
    .map((line) => `<p style="text-align: center">${escapeHtml(line)}</p>`)
    .join('');
}

function normalizePageChromeContent(
  source: Partial<WorkDocumentPageChromeContent> | undefined,
  fallback: WorkDocumentPageChromeContent,
): WorkDocumentPageChromeContent {
  return {
    headerHtml: sanitizeDocumentPageChromeHtml(
      source?.headerHtml ?? fallback.headerHtml,
    ),
    footerHtml: sanitizeDocumentPageChromeHtml(
      source?.footerHtml ?? fallback.footerHtml,
    ),
    showPageNumber: source?.showPageNumber ?? fallback.showPageNumber,
  };
}

function sanitizeAttributes(element: Element, tag: string) {
  if (MATHML_TAGS.has(tag)) {
    for (const attribute of Array.from(element.attributes)) {
      if (!MATHML_ATTRIBUTES.has(attribute.name.toLowerCase())) {
        element.removeAttributeNode(attribute);
      }
    }
    return;
  }
  if (!(element instanceof HTMLElement)) return;
  for (const attribute of Array.from(element.attributes)) {
    if (attribute.name.toLowerCase().startsWith('on'))
      element.removeAttribute(attribute.name);
  }
  const textAlign = ['left', 'center', 'right', 'justify'].includes(
    element.style.textAlign,
  )
    ? element.style.textAlign
    : '';
  const color = element.style.color;
  const scriptFonts =
    tag === 'span' ? documentScriptFontsFromElement(element) : null;
  const scriptFontSlot =
    tag === 'span' ? documentScriptFontSlotFromElement(element) : null;
  const scriptFontAttributes = scriptFonts
    ? documentScriptFontsDomAttributes(scriptFonts, scriptFontSlot)
    : {};
  const fontFamily =
    (scriptFonts
      ? documentScriptFontFamilyForRendering(
          scriptFonts,
          scriptFontSlot ?? documentScriptFontSlotFromHint(scriptFonts.hint),
          element.style.fontFamily,
        )
      : null) ??
    cssDocumentFontFamily(
      documentFontNameFromCssFamily(element.style.fontFamily),
    );
  const fontSize =
    tag === 'span'
      ? normalizedPageChromeFontSize(element.style.fontSize)
      : null;
  const textCase =
    tag === 'span'
      ? normalizeDocumentTextCase(
          element.getAttribute(DOCUMENT_TEXT_CASE_ATTRIBUTE),
        )
      : null;
  const characterSpacing =
    tag === 'span' ? documentCharacterSpacingTwipsFromElement(element) : null;
  const characterScale =
    tag === 'span' ? documentCharacterScalePercentFromElement(element) : null;
  const characterScaleAttributes =
    documentCharacterScaleDomAttributes(characterScale);
  const characterPosition =
    tag === 'span'
      ? documentCharacterPositionHalfPointsFromElement(element)
      : null;
  const characterPositionAttributes =
    documentCharacterPositionDomAttributes(characterPosition);
  const characterSpacingAttributes =
    documentCharacterSpacingDomAttributes(characterSpacing);
  const kerningThreshold =
    tag === 'span'
      ? documentKerningThresholdHalfPointsFromElement(element)
      : null;
  const kerningAttributes = documentKerningDomAttributes(
    kerningThreshold,
    fontSize,
  );
  const emphasisMark =
    tag === 'span' ? documentEmphasisMarkFromElement(element) : null;
  const emphasisAttributes = documentEmphasisMarkDomAttributes(emphasisMark);
  const hiddenText =
    tag === 'span' ? documentHiddenTextFromElement(element) : null;
  const underline =
    tag === 'u' ? documentUnderlineFormattingFromElement(element) : null;
  const underlineAttributes = underline
    ? documentUnderlineDomAttributes(underline)
    : {};
  const strike =
    tag === 's' || tag === 'strike'
      ? documentStrikeFormattingFromElement(element)
      : null;
  const strikeAttributes = strike ? documentStrikeDomAttributes(strike) : {};
  const paragraphShading =
    tag === 'p' ? parseDocumentParagraphShadingElement(element) : null;
  const shadingAttributes =
    documentParagraphShadingDomAttributes(paragraphShading);
  const paragraphBorders =
    tag === 'p' ? parseDocumentParagraphBordersElement(element) : null;
  const borderAttributes =
    documentParagraphBordersDomAttributes(paragraphBorders);
  const direction = element.getAttribute('dir')?.trim().toLowerCase();
  element.removeAttribute('style');
  element.removeAttribute(PARAGRAPH_BORDERS_ATTRIBUTE);
  element.removeAttribute(PARAGRAPH_SHADING_ATTRIBUTE);
  element.removeAttribute(DOCUMENT_UNDERLINE_STYLE_ATTRIBUTE);
  element.removeAttribute(DOCUMENT_UNDERLINE_COLOR_ATTRIBUTE);
  element.removeAttribute(DOCUMENT_UNDERLINE_THEME_COLOR_ATTRIBUTE);
  element.removeAttribute(DOCUMENT_STRIKE_STYLE_ATTRIBUTE);
  element.removeAttribute(DOCUMENT_CHARACTER_SCALE_ATTRIBUTE);
  element.removeAttribute(DOCUMENT_CHARACTER_POSITION_ATTRIBUTE);
  element.removeAttribute(DOCUMENT_CHARACTER_SPACING_ATTRIBUTE);
  element.removeAttribute(DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE);
  element.removeAttribute(DOCUMENT_EMPHASIS_MARK_ATTRIBUTE);
  element.removeAttribute(DOCUMENT_HIDDEN_TEXT_ATTRIBUTE);
  element.removeAttribute(DOCUMENT_SCRIPT_FONTS_ATTRIBUTE);
  element.removeAttribute(DOCUMENT_SCRIPT_FONT_SLOT_ATTRIBUTE);
  const styles = [
    textAlign ? `text-align: ${textAlign}` : '',
    color ? `color: ${color}` : '',
    fontFamily ? `font-family: ${fontFamily}` : '',
    fontSize ? `font-size: ${fontSize}` : '',
    textCase ? documentTextCaseCss(textCase) : '',
    characterScaleAttributes.style ?? '',
    characterPositionAttributes.style ?? '',
    characterSpacingAttributes.style ?? '',
    kerningAttributes.style ?? '',
    emphasisAttributes.style ?? '',
    underlineAttributes.style ?? '',
    strikeAttributes.style ?? '',
    borderAttributes.style ?? '',
    shadingAttributes.style ?? '',
  ].filter(Boolean);
  if (styles.length) element.setAttribute('style', styles.join('; '));

  if (tag === 'a') {
    const href = element.getAttribute('href')?.trim() ?? '';
    if (!/^(?:https?:|mailto:|#)/i.test(href)) element.removeAttribute('href');
  } else if (tag === 'img') {
    const source = element.getAttribute('src')?.trim() ?? '';
    if (!/^(?:https?:|blob:|data:image\/)/i.test(source))
      element.removeAttribute('src');
    normalizeImageIdentityAttributes(element);
  } else if (tag === 'ol') {
    const start = Number(element.getAttribute('start'));
    if (!Number.isSafeInteger(start) || start <= 0)
      element.removeAttribute('start');
    const type = element.getAttribute('type');
    if (type && !['1', 'A', 'a', 'I', 'i'].includes(type))
      element.removeAttribute('type');
  }
  if (tag === 'p') {
    const borders = borderAttributes[PARAGRAPH_BORDERS_ATTRIBUTE];
    if (borders) element.setAttribute(PARAGRAPH_BORDERS_ATTRIBUTE, borders);
    const shading = shadingAttributes[PARAGRAPH_SHADING_ATTRIBUTE];
    if (shading) element.setAttribute(PARAGRAPH_SHADING_ATTRIBUTE, shading);
    normalizeParagraphIdentityAttributes(element);
    normalizeParagraphDefaultCollapsedAttribute(element);
  }
  if (tag === 'tr') normalizeTableRowIdentityAttributes(element);
  if (tag === 'u') {
    for (const [name, value] of Object.entries(underlineAttributes)) {
      if (name !== 'style') element.setAttribute(name, value);
    }
  }
  if (tag === 's' || tag === 'strike') {
    for (const [name, value] of Object.entries(strikeAttributes)) {
      if (name !== 'style') element.setAttribute(name, value);
    }
  }
  if (tag === 'span' && textCase) {
    element.setAttribute(DOCUMENT_TEXT_CASE_ATTRIBUTE, textCase);
  } else {
    element.removeAttribute(DOCUMENT_TEXT_CASE_ATTRIBUTE);
  }
  if (tag === 'span' && characterSpacing !== null) {
    element.setAttribute(
      DOCUMENT_CHARACTER_SPACING_ATTRIBUTE,
      String(characterSpacing),
    );
  }
  if (tag === 'span' && characterScale !== null) {
    element.setAttribute(
      DOCUMENT_CHARACTER_SCALE_ATTRIBUTE,
      String(characterScale),
    );
  }
  if (tag === 'span' && characterPosition !== null) {
    element.setAttribute(
      DOCUMENT_CHARACTER_POSITION_ATTRIBUTE,
      String(characterPosition),
    );
  }
  if (tag === 'span' && kerningThreshold !== null) {
    element.setAttribute(
      DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE,
      String(kerningThreshold),
    );
  }
  if (tag === 'span' && emphasisMark !== null) {
    element.setAttribute(DOCUMENT_EMPHASIS_MARK_ATTRIBUTE, emphasisMark);
  }
  if (tag === 'span' && hiddenText !== null) {
    element.setAttribute(DOCUMENT_HIDDEN_TEXT_ATTRIBUTE, String(hiddenText));
  }
  if (tag === 'span' && scriptFonts) {
    for (const [name, value] of Object.entries(scriptFontAttributes)) {
      if (name !== 'style') element.setAttribute(name, value);
    }
  }
  if (direction === 'ltr' || direction === 'rtl')
    element.setAttribute('dir', direction);
  else element.removeAttribute('dir');
  const allowed =
    tag === 'span' && element.dataset.documentEquation === 'true'
      ? new Set(EQUATION_ATTRIBUTES)
      : tag === 'a'
        ? new Set(['dir', 'href', 'title', 'style'])
        : tag === 'img'
          ? new Set([
              'dir',
              'src',
              'alt',
              'title',
              'width',
              'height',
              'style',
              ...IMAGE_IDENTITY_ATTRIBUTES,
            ])
          : tag === 'ol'
            ? new Set(['dir', 'start', 'style', 'type'])
            : tag === 'p'
              ? new Set([
                  'dir',
                  'style',
                  PARAGRAPH_DEFAULT_COLLAPSED_ATTRIBUTE,
                  PARAGRAPH_BORDERS_ATTRIBUTE,
                  PARAGRAPH_SHADING_ATTRIBUTE,
                  ...PARAGRAPH_IDENTITY_ATTRIBUTES,
                ])
              : tag === 'tr'
                ? new Set(['dir', 'style', ...TABLE_ROW_IDENTITY_ATTRIBUTES])
                : new Set([
                    'colspan',
                    'dir',
                    'rowspan',
                    'style',
                    ...(tag === 'span'
                      ? [
                          DOCUMENT_TEXT_CASE_ATTRIBUTE,
                          DOCUMENT_CHARACTER_SCALE_ATTRIBUTE,
                          DOCUMENT_CHARACTER_POSITION_ATTRIBUTE,
                          DOCUMENT_CHARACTER_SPACING_ATTRIBUTE,
                          DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE,
                          DOCUMENT_EMPHASIS_MARK_ATTRIBUTE,
                          DOCUMENT_HIDDEN_TEXT_ATTRIBUTE,
                          DOCUMENT_SCRIPT_FONTS_ATTRIBUTE,
                          DOCUMENT_SCRIPT_FONT_SLOT_ATTRIBUTE,
                        ]
                      : []),
                    ...(tag === 'u'
                      ? [
                          DOCUMENT_UNDERLINE_STYLE_ATTRIBUTE,
                          DOCUMENT_UNDERLINE_COLOR_ATTRIBUTE,
                          DOCUMENT_UNDERLINE_THEME_COLOR_ATTRIBUTE,
                        ]
                      : []),
                    ...(tag === 's' || tag === 'strike'
                      ? [DOCUMENT_STRIKE_STYLE_ATTRIBUTE]
                      : []),
                  ]);
  for (const attribute of Array.from(element.attributes)) {
    if (!allowed.has(attribute.name.toLowerCase()))
      element.removeAttribute(attribute.name);
  }
}

function normalizedPageChromeFontSize(value: string): string | null {
  const match = /^(\d+(?:\.\d*)?|\.\d+)(pt|px)$/iu.exec(value.trim());
  if (!match) return null;
  const amount = Number(match[1]);
  const points = match[2]?.toLowerCase() === 'px' ? amount * 0.75 : amount;
  if (
    !Number.isFinite(points) ||
    points <= 0 ||
    points > 512 ||
    !Number.isSafeInteger(points * 2)
  ) {
    return null;
  }
  return `${Number(points.toFixed(1))}pt`;
}

function normalizeTableRowIdentityAttributes(element: HTMLElement): void {
  const identity = normalizeDocumentTableRowIdentity({
    rowId: element.getAttribute(TABLE_ROW_IDENTITY_ATTRIBUTES[0]),
    rowTextId: element.getAttribute(TABLE_ROW_IDENTITY_ATTRIBUTES[1]),
  });
  if (!identity) {
    for (const name of TABLE_ROW_IDENTITY_ATTRIBUTES) {
      element.removeAttribute(name);
    }
    return;
  }
  element.setAttribute(TABLE_ROW_IDENTITY_ATTRIBUTES[0], identity.rowId);
  element.setAttribute(TABLE_ROW_IDENTITY_ATTRIBUTES[1], identity.rowTextId);
}

function normalizeParagraphIdentityAttributes(element: HTMLElement): void {
  const identity = normalizeDocumentParagraphIdentity({
    paragraphId: element.getAttribute(PARAGRAPH_IDENTITY_ATTRIBUTES[0]),
    textId: element.getAttribute(PARAGRAPH_IDENTITY_ATTRIBUTES[1]),
  });
  if (!identity) {
    for (const name of PARAGRAPH_IDENTITY_ATTRIBUTES) {
      element.removeAttribute(name);
    }
    return;
  }
  element.setAttribute(PARAGRAPH_IDENTITY_ATTRIBUTES[0], identity.paragraphId);
  element.setAttribute(PARAGRAPH_IDENTITY_ATTRIBUTES[1], identity.textId);
}

function normalizeParagraphDefaultCollapsedAttribute(
  element: HTMLElement,
): void {
  const value = element.getAttribute(PARAGRAPH_DEFAULT_COLLAPSED_ATTRIBUTE);
  if (value === 'true' || value === '1') {
    element.setAttribute(PARAGRAPH_DEFAULT_COLLAPSED_ATTRIBUTE, 'true');
  } else if (value === 'false' || value === '0') {
    element.setAttribute(PARAGRAPH_DEFAULT_COLLAPSED_ATTRIBUTE, 'false');
  } else {
    element.removeAttribute(PARAGRAPH_DEFAULT_COLLAPSED_ATTRIBUTE);
  }
}

function normalizeImageIdentityAttributes(element: HTMLElement): void {
  const identity = normalizeDocumentImageIdentity({
    objectId: element.getAttribute(IMAGE_IDENTITY_ATTRIBUTES[0]),
    docPropertiesId: element.getAttribute(IMAGE_IDENTITY_ATTRIBUTES[1]),
    anchorId: element.getAttribute(IMAGE_IDENTITY_ATTRIBUTES[2]),
    editId: element.getAttribute(IMAGE_IDENTITY_ATTRIBUTES[3]),
  });
  if (!identity) {
    for (const name of IMAGE_IDENTITY_ATTRIBUTES) element.removeAttribute(name);
    return;
  }
  element.setAttribute(IMAGE_IDENTITY_ATTRIBUTES[0], identity.objectId);
  element.setAttribute(
    IMAGE_IDENTITY_ATTRIBUTES[1],
    String(identity.docPropertiesId),
  );
  element.setAttribute(IMAGE_IDENTITY_ATTRIBUTES[2], identity.anchorId);
  element.setAttribute(IMAGE_IDENTITY_ATTRIBUTES[3], identity.editId);
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}
