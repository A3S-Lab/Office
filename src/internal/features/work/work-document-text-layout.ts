import type { Editor } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type {
  OfficeKernelTextLayoutParagraph,
  OfficeKernelTextLayoutRun,
} from '../../kernel/office-kernel-protocol';
import type { WorkDocumentLayoutFont } from './work-document-fonts';
import {
  documentBlockId,
  documentListChildId,
  elementForNode,
  finitePixels,
  isDocumentListNode,
  nonNegativePixels,
  positivePixels,
  reusableDocumentLayoutBlocks,
} from './work-document-pagination-dom';
import type {
  DocumentPaginationSnapshot,
  DocumentTextLayoutCollection,
} from './work-document-pagination-types';
import {
  DEFAULT_DOCUMENT_TAB_INTERVAL_PX,
  normalizeDocumentTabStops,
} from './work-document-tab-stops';

const MAX_DOCUMENT_TEXT_LAYOUT_PARAGRAPHS = 1_024;
const MAX_DOCUMENT_TEXT_LAYOUT_RUNS = 16_384;
const MAX_DOCUMENT_TEXT_LAYOUT_BYTES = 1_048_576;
const MAX_DOCUMENT_TEXT_LAYOUT_RUNS_PER_PARAGRAPH = 4_096;
const MAX_DOCUMENT_TEXT_LAYOUT_FONTS_PER_RUN = 8;

export function collectDocumentTextLayoutParagraphs(
  editor: Editor,
  fonts: readonly WorkDocumentLayoutFont[],
  loadedFontIds: ReadonlySet<string>,
  previous: DocumentPaginationSnapshot | null = null,
  dirtyFrom = 0,
): DocumentTextLayoutCollection {
  const paragraphs: OfficeKernelTextLayoutParagraph[] = [];
  if (!fonts.length || loadedFontIds.size === 0) return { paragraphs };
  let runCount = 0;
  let textBytes = 0;
  const collectParagraph = (
    node: ProseMirrorNode,
    element: HTMLElement,
    id: string,
    position: number,
  ): void => {
    if (paragraphs.length >= MAX_DOCUMENT_TEXT_LAYOUT_PARAGRAPHS) return;
    if (
      reusableDocumentLayoutBlocks(
        previous?.blocks ?? [],
        id,
        element,
        position + node.nodeSize,
        dirtyFrom,
      ).length
    ) {
      return;
    }
    const paragraph = documentTextLayoutParagraph(
      node,
      element,
      id,
      fonts,
      loadedFontIds,
    );
    if (!paragraph) return;
    const paragraphBytes = utf8ByteLength(paragraph.text);
    if (
      runCount + paragraph.runs.length > MAX_DOCUMENT_TEXT_LAYOUT_RUNS ||
      textBytes + paragraphBytes > MAX_DOCUMENT_TEXT_LAYOUT_BYTES
    ) {
      return;
    }
    runCount += paragraph.runs.length;
    textBytes += paragraphBytes;
    paragraphs.push(paragraph);
  };
  editor.state.doc.forEach((section, sectionPosition) => {
    if (section.type.name !== 'documentSection') return;
    let requiresBrowserLineMeasurement = false;
    section.forEach((node, offset, index) => {
      if (paragraphs.length >= MAX_DOCUMENT_TEXT_LAYOUT_PARAGRAPHS) return;
      if (isFloatingSquareDocumentImage(node)) {
        requiresBrowserLineMeasurement = true;
        return;
      }
      if (requiresBrowserLineMeasurement) return;
      const position = sectionPosition + offset + 1;
      const element = elementForNode(editor, position);
      if (!element) return;
      const id = documentBlockId(sectionPosition, index, position);
      if (isDocumentListNode(node)) {
        collectDocumentListTextLayoutParagraphs(
          editor,
          node,
          position,
          id,
          collectParagraph,
        );
        return;
      }
      collectParagraph(node, element, id, position);
    });
  });
  return { paragraphs };
}

function isFloatingSquareDocumentImage(node: ProseMirrorNode): boolean {
  return (
    node.type.name === 'image' &&
    node.attrs.layout === 'square' &&
    (node.attrs.alignment === 'left' || node.attrs.alignment === 'right')
  );
}

function collectDocumentListTextLayoutParagraphs(
  editor: Editor,
  list: ProseMirrorNode,
  listPosition: number,
  listId: string,
  collectParagraph: (
    node: ProseMirrorNode,
    element: HTMLElement,
    id: string,
    position: number,
  ) => void,
): void {
  list.forEach((item, itemOffset, itemIndex) => {
    if (item.type.name !== 'listItem') return;
    const itemPosition = listPosition + itemOffset + 1;
    item.forEach((node, offset, index) => {
      const position = itemPosition + offset + 1;
      const nestedList = isDocumentListNode(node);
      const id = documentListChildId(
        listId,
        itemIndex,
        node,
        index,
        nestedList,
      );
      if (nestedList) {
        collectDocumentListTextLayoutParagraphs(
          editor,
          node,
          position,
          id,
          collectParagraph,
        );
        return;
      }
      const element = elementForNode(editor, position);
      if (element) collectParagraph(node, element, id, position);
    });
  });
}

function utf8ByteLength(value: string): number {
  let bytes = 0;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code <= 0x7f) {
      bytes += 1;
    } else if (code <= 0x7ff) {
      bytes += 2;
    } else if (
      code >= 0xd800 &&
      code <= 0xdbff &&
      index + 1 < value.length &&
      value.charCodeAt(index + 1) >= 0xdc00 &&
      value.charCodeAt(index + 1) <= 0xdfff
    ) {
      bytes += 4;
      index += 1;
    } else {
      bytes += 3;
    }
  }
  return bytes;
}

function documentTextLayoutParagraph(
  node: ProseMirrorNode,
  element: HTMLElement,
  id: string,
  fonts: readonly WorkDocumentLayoutFont[],
  loadedFontIds: ReadonlySet<string>,
): OfficeKernelTextLayoutParagraph | null {
  const text = documentTextLayoutContent(node);
  if (
    node.type.name !== 'paragraph' ||
    text === null ||
    text.length < 2 ||
    text.includes('\u00ad') ||
    element.querySelector(
      [
        'img',
        'br',
        'table',
        '[data-document-note]',
        [
          '[contenteditable="false"]',
          ':not(.work-document-auto-page-break)',
          ':not([data-document-tab])',
        ].join(''),
      ].join(','),
    )
  ) {
    return null;
  }
  const style = getComputedStyle(element);
  const paragraphStyle = documentTextParagraphStyle(style);
  const hasTabs = text.includes('\t');
  if (!paragraphStyle || (hasTabs && paragraphStyle.direction !== 'ltr')) {
    return null;
  }
  const runs = collectDocumentTextLayoutRuns(
    element,
    text,
    fonts,
    loadedFontIds,
    paragraphStyle,
  );
  if (!runs) return null;
  const width =
    element.clientWidth || element.offsetWidth || positivePixels(style.width);
  const maxWidth =
    width -
    nonNegativePixels(style.paddingLeft) -
    nonNegativePixels(style.paddingRight);
  if (maxWidth <= 1) return null;
  const textIndent = finitePixels(style.textIndent);
  return {
    id,
    text,
    runs,
    maxWidth,
    firstLineMaxWidth: Math.max(1, maxWidth - textIndent),
    direction: paragraphStyle.direction,
    whiteSpace: paragraphStyle.whiteSpace,
    ...(hasTabs
      ? {
          tabLayout: {
            origin: nonNegativePixels(style.marginLeft),
            firstLineIndent: textIndent,
            defaultInterval: DEFAULT_DOCUMENT_TAB_INTERVAL_PX,
            stops: normalizeDocumentTabStops(node.attrs.tabStops).map(
              ({ position, alignment }) => ({ position, alignment }),
            ),
          },
        }
      : {}),
  };
}

export function documentTextLayoutContent(
  node: ProseMirrorNode,
): string | null {
  const content: string[] = [];
  let supported = true;
  node.forEach((child) => {
    if (child.isText) {
      content.push(child.text ?? '');
    } else if (child.type.name === 'documentTab') {
      content.push('\t');
    } else {
      supported = false;
    }
  });
  return supported ? content.join('') : null;
}

interface DocumentTextParagraphStyle {
  direction: 'ltr' | 'rtl';
  whiteSpace: 'breakSpaces' | 'normal';
}

interface DocumentTextLayoutRunStyle {
  fontId: string;
  fallbackFontIds?: string[];
  fontSize: number;
  lineHeight: number;
  letterSpacing: number;
  ligatures: boolean;
  kerning: boolean;
}

export function collectDocumentTextLayoutRuns(
  element: HTMLElement,
  expectedText: string,
  fonts: readonly WorkDocumentLayoutFont[],
  loadedFontIds: ReadonlySet<string>,
  paragraphStyle = documentTextParagraphStyle(getComputedStyle(element)),
): OfficeKernelTextLayoutRun[] | null {
  if (!paragraphStyle) return null;
  const tokens = documentTextLayoutDomTokens(element);
  if (!tokens) return null;
  const text: string[] = [];
  const runs: OfficeKernelTextLayoutRun[] = [];
  let utf16Offset = 0;
  for (const token of tokens) {
    const runStyle = documentTextLayoutRunStyle(
      getComputedStyle(token.styleElement),
      fonts,
      loadedFontIds,
      paragraphStyle,
      token.styleElement === element,
    );
    if (!runStyle) return null;
    const startUtf16 = utf16Offset;
    utf16Offset += token.text.length;
    text.push(token.text);
    const previous = runs.at(-1);
    if (
      previous &&
      previous.endUtf16 === startUtf16 &&
      documentTextLayoutRunStyleKey(previous) ===
        documentTextLayoutRunStyleKey(runStyle)
    ) {
      previous.endUtf16 = utf16Offset;
    } else {
      runs.push({
        startUtf16,
        endUtf16: utf16Offset,
        ...runStyle,
      });
    }
    if (runs.length > MAX_DOCUMENT_TEXT_LAYOUT_RUNS_PER_PARAGRAPH) {
      return null;
    }
  }
  return text.join('') === expectedText && runs.length ? runs : null;
}

interface DocumentTextLayoutDomToken {
  text: string;
  styleElement: HTMLElement;
}

function documentTextLayoutDomTokens(
  element: HTMLElement,
): DocumentTextLayoutDomToken[] | null {
  const tokens: DocumentTextLayoutDomToken[] = [];
  const visit = (node: Node): boolean => {
    if (node.nodeType === Node.TEXT_NODE) {
      const textNode = node as Text;
      if (!textNode.data.length) return true;
      const parent = textNode.parentElement;
      if (!parent) return false;
      tokens.push({ text: textNode.data, styleElement: parent });
      return true;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return true;
    const child = node as HTMLElement;
    if (child.classList.contains('work-document-auto-page-break')) return true;
    if (child.matches('[data-document-tab]')) {
      tokens.push({
        text: '\t',
        styleElement: child.parentElement ?? element,
      });
      return true;
    }
    if (child.contentEditable === 'false') return true;
    return Array.from(child.childNodes).every(visit);
  };
  return Array.from(element.childNodes).every(visit) ? tokens : null;
}

function documentTextParagraphStyle(
  style: CSSStyleDeclaration,
): DocumentTextParagraphStyle | null {
  if (
    !cssValueIs(style.whiteSpace, 'normal', 'break-spaces') ||
    !cssValueIs(style.overflowWrap, 'normal', 'anywhere', 'break-word') ||
    !cssValueIs(style.wordBreak, 'normal') ||
    !cssValueIs(style.hyphens, 'manual', 'none') ||
    !cssValueIs(style.textTransform, 'none') ||
    !cssValueIs(style.unicodeBidi, 'normal', 'isolate')
  ) {
    return null;
  }
  return {
    direction: style.direction === 'rtl' ? 'rtl' : 'ltr',
    whiteSpace: style.whiteSpace === 'break-spaces' ? 'breakSpaces' : 'normal',
  };
}

function documentTextLayoutRunStyle(
  style: CSSStyleDeclaration,
  fonts: readonly WorkDocumentLayoutFont[],
  loadedFontIds: ReadonlySet<string>,
  paragraphStyle: DocumentTextParagraphStyle,
  paragraphRoot: boolean,
): DocumentTextLayoutRunStyle | null {
  const direction = style.direction === 'rtl' ? 'rtl' : 'ltr';
  const whiteSpace =
    style.whiteSpace === 'break-spaces' ? 'breakSpaces' : 'normal';
  if (
    direction !== paragraphStyle.direction ||
    whiteSpace !== paragraphStyle.whiteSpace ||
    !cssValueIs(style.whiteSpace, 'normal', 'break-spaces') ||
    !cssValueIs(style.overflowWrap, 'normal', 'anywhere', 'break-word') ||
    !cssValueIs(style.wordBreak, 'normal') ||
    !cssValueIs(style.hyphens, 'manual', 'none') ||
    !cssValueIs(style.wordSpacing, 'normal', '0', '0px') ||
    !cssValueIs(style.textTransform, 'none') ||
    !cssValueIs(style.fontVariationSettings, 'normal') ||
    !cssValueIs(style.fontVariantCaps, 'normal') ||
    !cssValueIs(style.fontVariantEastAsian, 'normal') ||
    !cssValueIs(style.fontVariantNumeric, 'normal') ||
    !cssValueIs(style.fontVariantPosition, 'normal') ||
    !cssValueIs(style.fontStretch, 'normal', '100%') ||
    !cssValueIs(style.verticalAlign, 'baseline') ||
    (!cssValueIs(style.unicodeBidi, 'normal') &&
      !(paragraphRoot && style.unicodeBidi === 'isolate'))
  ) {
    return null;
  }
  const fontStyle =
    !style.fontStyle || style.fontStyle === 'normal'
      ? 'normal'
      : style.fontStyle === 'italic'
        ? 'italic'
        : null;
  if (!fontStyle) return null;
  const families = cssFontFamilies(style.fontFamily);
  if (!families?.length) return null;
  const weight = cssFontWeight(style.fontWeight);
  const exactFonts: WorkDocumentLayoutFont[] = [];
  for (const family of families) {
    if (isGenericCssFontFamily(family)) break;
    const font = fonts.find(
      (candidate) =>
        loadedFontIds.has(candidate.id) &&
        candidate.family === family &&
        (candidate.weight ?? 400) === weight &&
        (candidate.style ?? 'normal') === fontStyle,
    );
    if (!font) return null;
    if (!exactFonts.some((candidate) => candidate.id === font.id)) {
      exactFonts.push(font);
    }
    if (exactFonts.length > MAX_DOCUMENT_TEXT_LAYOUT_FONTS_PER_RUN) {
      return null;
    }
  }
  const [font, ...fallbackFonts] = exactFonts;
  if (!font) return null;
  const fontSize = positivePixels(style.fontSize);
  const lineHeight =
    positivePixels(style.lineHeight) || Math.max(1, fontSize * 1.2);
  if (
    fontSize <= 0 ||
    fontSize > 512 ||
    lineHeight <= 0 ||
    lineHeight > 2_048
  ) {
    return null;
  }
  const ligatures = cssLigatures(style);
  const kerning = cssKerning(style.fontKerning);
  if (ligatures === null || kerning === null) return null;
  const letterSpacing =
    !style.letterSpacing || style.letterSpacing === 'normal'
      ? 0
      : finitePixels(style.letterSpacing);
  if (letterSpacing < -100 || letterSpacing > 100) return null;
  return {
    fontId: font.id,
    ...(fallbackFonts.length
      ? { fallbackFontIds: fallbackFonts.map((candidate) => candidate.id) }
      : {}),
    fontSize,
    lineHeight,
    letterSpacing,
    ligatures,
    kerning,
  };
}

function documentTextLayoutRunStyleKey(
  style: Pick<OfficeKernelTextLayoutRun, 'fontId' | 'fontSize' | 'lineHeight'> &
    Partial<
      Pick<
        OfficeKernelTextLayoutRun,
        'fallbackFontIds' | 'kerning' | 'letterSpacing' | 'ligatures'
      >
    >,
): string {
  return [
    style.fontId,
    (style.fallbackFontIds ?? []).join('\u0002'),
    style.fontSize,
    style.lineHeight,
    style.letterSpacing ?? 0,
    style.ligatures ?? true,
    style.kerning ?? true,
  ].join('\u0000');
}

function cssValueIs(value: string, ...supported: string[]): boolean {
  return !value || supported.includes(value);
}

function cssLigatures(style: CSSStyleDeclaration): boolean | null {
  const variant = style.fontVariantLigatures;
  if (variant && variant !== 'normal' && variant !== 'none') return null;
  const features = style.fontFeatureSettings;
  if (!features || features === 'normal') return variant !== 'none';
  if (/^["']?liga["']?\s+(?:0|off)$/i.test(features.trim())) return false;
  return null;
}

function cssKerning(value: string): boolean | null {
  if (!value || value === 'auto' || value === 'normal') return true;
  if (value === 'none') return false;
  return null;
}

function cssFontFamilies(value: string): string[] | null {
  const families: string[] = [];
  let current = '';
  let quote: '"' | "'" | null = null;
  let escaped = false;
  const pushCurrent = (): boolean => {
    const family = current.trim();
    current = '';
    if (!family) return false;
    families.push(family);
    return true;
  };
  for (const character of value) {
    if (escaped) {
      current += character;
      escaped = false;
      continue;
    }
    if (character === '\\') {
      escaped = true;
      continue;
    }
    if (quote) {
      if (character === quote) quote = null;
      else current += character;
      continue;
    }
    if (character === '"' || character === "'") {
      if (current.trim()) return null;
      quote = character;
      continue;
    }
    if (character === ',') {
      if (!pushCurrent()) return null;
      continue;
    }
    current += character;
  }
  if (escaped || quote || !pushCurrent()) return null;
  return families;
}

function isGenericCssFontFamily(family: string): boolean {
  return CSS_GENERIC_FONT_FAMILIES.has(family.toLowerCase());
}

const CSS_GENERIC_FONT_FAMILIES = new Set([
  'cursive',
  'emoji',
  'fangsong',
  'fantasy',
  'math',
  'monospace',
  'sans-serif',
  'serif',
  'system-ui',
  'ui-monospace',
  'ui-rounded',
  'ui-sans-serif',
  'ui-serif',
]);

function cssFontWeight(value: string): number {
  if (value === 'normal') return 400;
  if (value === 'bold') return 700;
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : 400;
}
