import type { DocumentParagraphPagination } from './work-document-paragraph-formatting';
import {
  type DocxParagraphStyleSource,
  docxParagraphPropertySources,
  resolveDocxParagraphStyleResolver,
} from './work-docx-paragraph-styles';
import { attribute, descendants, directChild } from './work-ooxml-package';

export interface ImportedDocxParagraphPaginationMarker {
  marker: string;
  pagination: Partial<DocumentParagraphPagination>;
}

export interface ImportedDocxParagraphPaginationMarkers {
  paragraphs: ImportedDocxParagraphPaginationMarker[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const PARAGRAPH_PAGINATION_MARKER_PATTERN =
  /__A3S_WORK_PARAGRAPH_PAGINATION_\d+__/g;

export function markDocxParagraphPagination(
  document: Document,
  styleSource?: DocxParagraphStyleSource,
): ImportedDocxParagraphPaginationMarkers {
  const paragraphs: ImportedDocxParagraphPaginationMarker[] = [];
  const styles = resolveDocxParagraphStyleResolver(styleSource);
  for (const paragraph of descendants(document, 'p')) {
    let properties = directChild(paragraph, 'pPr');
    if (properties && directChild(properties, 'numPr')) continue;
    const pagination = paragraphPagination(
      docxParagraphPropertySources(properties, styles),
    );
    if (!Object.keys(pagination).length) continue;
    properties ??= insertParagraphProperties(document, paragraph);
    const marker = `__A3S_WORK_PARAGRAPH_PAGINATION_${paragraphs.length + 1}__`;
    insertParagraphMarker(document, paragraph, properties, marker);
    paragraphs.push({ marker, pagination });
  }
  return { paragraphs };
}

export function applyImportedDocxParagraphPaginationMarkers(
  document: Document,
  markers: ImportedDocxParagraphPaginationMarkers,
): void {
  const paginationByMarker = new Map(
    markers.paragraphs.map(({ marker, pagination }) => [marker, pagination]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_PARAGRAPH_PAGINATION_')) continue;
    const block = closestParagraphBlock(node.parentElement);
    node.data = node.data.replace(
      PARAGRAPH_PAGINATION_MARKER_PATTERN,
      (marker) => {
        const pagination = paginationByMarker.get(marker);
        if (block && pagination) applyParagraphPagination(block, pagination);
        return '';
      },
    );
  }
  document.body.normalize();
}

export function hasImportedDocxParagraphPaginationMarkers(
  markers: ImportedDocxParagraphPaginationMarkers,
): boolean {
  return markers.paragraphs.length > 0;
}

function paragraphPagination(
  sources: readonly Element[],
): Partial<DocumentParagraphPagination> {
  const pagination: Partial<DocumentParagraphPagination> = {};
  for (const properties of sources) {
    const keepLines = directChild(properties, 'keepLines');
    const keepWithNext = directChild(properties, 'keepNext');
    const pageBreakBefore = directChild(properties, 'pageBreakBefore');
    const widowControl = directChild(properties, 'widowControl');
    if (keepLines) pagination.keepLines = onOffValue(keepLines);
    if (keepWithNext) pagination.keepWithNext = onOffValue(keepWithNext);
    if (pageBreakBefore)
      pagination.pageBreakBefore = onOffValue(pageBreakBefore);
    if (widowControl) pagination.widowControl = onOffValue(widowControl);
  }
  return pagination;
}

function onOffValue(element: Element): boolean {
  const value = (attribute(element, 'val') ?? attribute(element, 'w:val'))
    ?.trim()
    .toLowerCase();
  return value !== '0' && value !== 'false' && value !== 'off';
}

function insertParagraphMarker(
  document: Document,
  paragraph: Element,
  properties: Element,
  marker: string,
): void {
  const run = document.createElementNS(WORD_NAMESPACE, 'w:r');
  const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  paragraph.insertBefore(run, properties.nextSibling);
}

function insertParagraphProperties(
  document: Document,
  paragraph: Element,
): Element {
  const properties = document.createElementNS(WORD_NAMESPACE, 'w:pPr');
  paragraph.insertBefore(properties, paragraph.firstChild);
  return properties;
}

function closestParagraphBlock(element: Element | null): HTMLElement | null {
  const block = element?.closest('p, h1, h2, h3, h4, h5, h6, blockquote');
  return block instanceof HTMLElement ? block : null;
}

function applyParagraphPagination(
  element: HTMLElement,
  pagination: Partial<DocumentParagraphPagination>,
): void {
  setBooleanAttribute(element, 'data-office-keep-lines', pagination.keepLines);
  setBooleanAttribute(
    element,
    'data-office-keep-with-next',
    pagination.keepWithNext,
  );
  setBooleanAttribute(
    element,
    'data-office-page-break-before',
    pagination.pageBreakBefore,
  );
  setBooleanAttribute(
    element,
    'data-office-widow-control',
    pagination.widowControl,
  );
}

function setBooleanAttribute(
  element: HTMLElement,
  name: string,
  value: boolean | undefined,
): void {
  if (value !== undefined) element.setAttribute(name, String(value));
}

function textNodes(root: ParentNode): Text[] {
  const document = root.ownerDocument;
  const walker = document?.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}
