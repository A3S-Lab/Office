import { documentAutoLineHeight } from './work-document-paragraph-formatting';
import { serializeDocumentParagraphFormatting } from './work-document-paragraph-format-changes';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { resolveDocxParagraphAlignment } from './work-docx-paragraph-alignment-import';
import { resolveDocxParagraphBordersFromSources } from './work-docx-paragraph-borders-import';
import { resolveDocxParagraphDirection } from './work-docx-paragraph-direction-import';
import { resolveDocxParagraphIndent } from './work-docx-paragraph-indent-import';
import { resolveDocxParagraphPagination } from './work-docx-paragraph-pagination-import';
import { resolveDocxParagraphShadingFromSources } from './work-docx-paragraph-shading-import';
import { resolveDocxParagraphSpacing } from './work-docx-paragraph-spacing-import';
import {
  type DocxParagraphStyleSource,
  docxParagraphPropertySources,
  resolveDocxParagraphStyleResolver,
} from './work-docx-paragraph-styles';
import {
  type DocxTableStyleSource,
  docxTableParagraphPropertySources,
  resolveDocxTableStyleResolver,
} from './work-docx-table-styles';
import { resolveDocxParagraphTabStops } from './work-docx-tab-stop-import';
import {
  type DocxThemeSource,
  resolveDocxThemeResolver,
} from './work-docx-theme';
import { attribute, descendants, directChildren } from './work-ooxml-package';

export interface ImportedDocxParagraphFormattingChangeMarker {
  marker: string;
  id: string;
  author: string;
  date: string;
  before: string;
}

export interface ImportedDocxParagraphFormattingChangeMarkers {
  paragraphs: ImportedDocxParagraphFormattingChangeMarker[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WORD_2012_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2012/wordml';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const MAX_PARAGRAPH_FORMATTING_CHANGES = 65_536;
const PARAGRAPH_FORMATTING_CHANGE_MARKER_PATTERN =
  /__A3S_WORK_PARAGRAPH_FORMAT_CHANGE_\d+__/g;
const SUPPORTED_PARAGRAPH_PROPERTY_CHANGE_CHILDREN = new Set([
  'jc',
  'bidi',
  'ind',
  'spacing',
  'keepLines',
  'keepNext',
  'pageBreakBefore',
  'widowControl',
  'contextualSpacing',
  'outlineLvl',
  'tabs',
  'pBdr',
  'shd',
]);

export function markDocxParagraphFormattingChanges(
  document: Document,
  styleSource?: DocxParagraphStyleSource,
  themeSource?: DocxThemeSource,
  tableStyleSource?: DocxTableStyleSource,
): ImportedDocxParagraphFormattingChangeMarkers {
  const paragraphs: ImportedDocxParagraphFormattingChangeMarker[] = [];
  const styles = resolveDocxParagraphStyleResolver(styleSource);
  const theme = resolveDocxThemeResolver(themeSource);
  const tableStyles = resolveDocxTableStyleResolver(tableStyleSource);
  for (const paragraph of descendants(document, 'p')) {
    if (!isWordElement(paragraph)) continue;
    const propertyNodes = directChildren(paragraph, 'pPr').filter(
      (element) => element.namespaceURI === paragraph.namespaceURI,
    );
    if (propertyNodes.length !== 1) continue;
    const properties = propertyNodes[0];
    if (!properties) continue;
    const change = supportedParagraphFormattingChange(properties);
    if (!change) continue;
    if (paragraphs.length >= MAX_PARAGRAPH_FORMATTING_CHANGES) continue;
    const sources = docxParagraphPropertySources(
      change.properties,
      styles,
      docxTableParagraphPropertySources(paragraph, tableStyles),
    );
    const borders = resolveDocxParagraphBordersFromSources(sources, theme);
    if (borders.invalidCount || borders.spoofedCount) continue;
    const indent = resolveDocxParagraphIndent(sources);
    const spacing = resolveDocxParagraphSpacing(sources);
    const pagination = resolveDocxParagraphPagination(sources);
    const tabStops = resolveDocxParagraphTabStops(sources);
    const before = serializeDocumentParagraphFormatting({
      textAlign: resolveDocxParagraphAlignment(sources),
      paragraphDirection: resolveDocxParagraphDirection(sources),
      indentLevel: indent.left / 24,
      rightIndent: indent.right,
      firstLineIndent: indent.firstLine,
      spaceBefore: spacing.before ?? null,
      spaceAfter: spacing.after ?? null,
      lineHeight: spacing.lineHeight ?? null,
      lineRule: spacing.lineRule ?? null,
      autoLineHeight:
        spacing.lineRule === 'auto'
          ? documentAutoLineHeight(spacing.lineHeight)
          : null,
      keepLines: pagination.keepLines ?? null,
      keepWithNext: pagination.keepWithNext ?? null,
      pageBreakBefore: pagination.pageBreakBefore ?? null,
      widowControl: pagination.widowControl ?? null,
      contextualSpacing: pagination.contextualSpacing ?? null,
      outlineLevel: pagination.outlineLevel ?? null,
      tabStops: tabStops.length ? tabStops : null,
      paragraphBorders: borders.borders,
      paragraphShading: resolveDocxParagraphShadingFromSources(sources, theme),
      defaultCollapsed: pagination.defaultCollapsed ?? null,
    });
    const marker = `__A3S_WORK_PARAGRAPH_FORMAT_CHANGE_${paragraphs.length + 1}__`;
    insertParagraphMarker(document, paragraph, properties, marker);
    paragraphs.push({
      marker,
      id: `docx-paragraph-format-change-${change.id}`,
      author: change.author,
      date: normalizeRevisionDate(change.date),
      before,
    });
  }
  return { paragraphs };
}

export function applyImportedDocxParagraphFormattingChangeMarkers(
  document: Document,
  markers: ImportedDocxParagraphFormattingChangeMarkers,
): void {
  const changes = new Map(
    markers.paragraphs.map((change) => [change.marker, change]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_PARAGRAPH_FORMAT_CHANGE_')) continue;
    node.data = node.data.replace(
      PARAGRAPH_FORMATTING_CHANGE_MARKER_PATTERN,
      (marker) => {
        const change = changes.get(marker);
        const block = change
          ? closestParagraphBlock(node.parentElement, node)
          : null;
        if (block && change) applyParagraphChange(block, change);
        return '';
      },
    );
  }
  document.body.normalize();
}

export function hasImportedDocxParagraphFormattingChangeMarkers(
  markers: ImportedDocxParagraphFormattingChangeMarkers,
): boolean {
  return markers.paragraphs.length > 0;
}

export function isSupportedDocxParagraphFormattingChange(
  change: Element,
): boolean {
  const parent = change.parentElement;
  const paragraph = parent?.parentElement;
  return Boolean(
    parent?.localName === 'pPr' &&
      DOCX_WORDPROCESSING_NAMESPACES.has(parent.namespaceURI ?? '') &&
      paragraph?.localName === 'p' &&
      DOCX_WORDPROCESSING_NAMESPACES.has(paragraph.namespaceURI ?? '') &&
      Array.from(parent.children).filter(
        (child) => child.localName === 'pPrChange',
      ).length === 1 &&
      supportedParagraphFormattingChangeElement(change),
  );
}

function supportedParagraphFormattingChange(properties: Element): {
  id: string;
  author: string;
  date: string | null;
  properties: Element;
} | null {
  const changes = Array.from(properties.children).filter(
    (child) => child.localName === 'pPrChange',
  );
  return changes.length === 1
    ? supportedParagraphFormattingChangeElement(changes[0])
    : null;
}

function supportedParagraphFormattingChangeElement(change: Element): {
  id: string;
  author: string;
  date: string | null;
  properties: Element;
} | null {
  if (
    change.localName !== 'pPrChange' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(change.namespaceURI ?? '')
  ) {
    return null;
  }
  const propertyNodes = Array.from(change.children).filter(
    (child) =>
      child.localName === 'pPr' &&
      DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? ''),
  );
  if (propertyNodes.length !== 1 || change.children.length !== 1) return null;
  const properties = propertyNodes[0];
  if (!properties) return null;
  const names = new Set<string>();
  for (const child of Array.from(properties.children)) {
    const supportedWordChild =
      DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '') &&
      SUPPORTED_PARAGRAPH_PROPERTY_CHANGE_CHILDREN.has(child.localName);
    const supportedCollapsed =
      child.localName === 'collapsed' &&
      child.namespaceURI === WORD_2012_NAMESPACE;
    if (
      (!supportedWordChild && !supportedCollapsed) ||
      names.has(child.localName)
    ) {
      return null;
    }
    names.add(child.localName);
  }
  const id = attribute(change, 'id')?.trim() ?? '';
  const author = attribute(change, 'author')?.trim() ?? '';
  const date = attribute(change, 'date');
  if (!/^\+?\d{1,10}$/.test(id) || !author || author.length > 255) return null;
  if (date && !Number.isFinite(Date.parse(date))) return null;
  return { id, author, date, properties };
}

function applyParagraphChange(
  element: HTMLElement,
  change: ImportedDocxParagraphFormattingChangeMarker,
): void {
  element.dataset.documentChange = 'true';
  element.dataset.changeKind = 'paragraph-formatting';
  element.dataset.changeId = change.id;
  element.dataset.changeAuthor = change.author;
  element.dataset.changeDate = change.date;
  element.dataset.changeBefore = change.before;
}

function insertParagraphMarker(
  document: Document,
  paragraph: Element,
  properties: Element,
  marker: string,
): void {
  const namespace = paragraph.namespaceURI ?? WORD_NAMESPACE;
  const prefix = paragraph.prefix || 'w';
  const run = document.createElementNS(namespace, `${prefix}:r`);
  const text = document.createElementNS(namespace, `${prefix}:t`);
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  paragraph.insertBefore(run, properties.nextSibling);
}

function closestParagraphBlock(
  element: Element | null,
  markerNode: Node,
): HTMLElement | null {
  const block = element?.closest('p, h1, h2, h3, h4, h5, h6, blockquote, li');
  if (!(block instanceof HTMLElement)) return null;
  return block.tagName.toLowerCase() === 'li'
    ? wrapListItemParagraph(block, markerNode)
    : block;
}

function wrapListItemParagraph(
  item: HTMLElement,
  markerNode: Node,
): HTMLElement {
  let anchor: Node = markerNode;
  while (anchor.parentNode && anchor.parentNode !== item) {
    anchor = anchor.parentNode;
  }
  const children = Array.from(item.childNodes);
  const anchorIndex = children.indexOf(anchor as ChildNode);
  if (anchorIndex < 0) return item;
  let start = anchorIndex;
  let end = anchorIndex;
  while (start > 0 && !isListItemBlock(children[start - 1])) start -= 1;
  while (end + 1 < children.length && !isListItemBlock(children[end + 1])) {
    end += 1;
  }
  const paragraph = item.ownerDocument.createElement('p');
  const grouped = children.slice(start, end + 1);
  item.insertBefore(paragraph, grouped[0] ?? null);
  paragraph.append(...grouped);
  return paragraph;
}

function isListItemBlock(node: Node | undefined): boolean {
  return (
    node instanceof HTMLElement &&
    [
      'blockquote',
      'div',
      'h1',
      'h2',
      'h3',
      'h4',
      'h5',
      'h6',
      'hr',
      'ol',
      'p',
      'pre',
      'table',
      'ul',
    ].includes(node.tagName.toLowerCase())
  );
}

function normalizeRevisionDate(value: string | null): string {
  if (!value) return '';
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time).toISOString() : '';
}

function isWordElement(element: Element): boolean {
  return DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '');
}

function textNodes(root: ParentNode): Text[] {
  const document = root.ownerDocument;
  const walker = document?.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}
