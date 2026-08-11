import {
  applyDocumentTableRowIdentityToElement,
  normalizeDocumentTableRowIdentity,
  type WorkDocumentTableRowIdentity,
} from './work-document-table-row-identity';
import { normalizeDocumentParagraphId } from './work-document-paragraph-identity';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { attribute, descendants, directChild } from './work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';

export interface ImportedDocxTableRowMarker {
  marker: string;
  cantSplit?: boolean;
  repeatHeader?: boolean;
  rowId?: string;
  rowHeight?: number;
  rowHeightRule?: 'atLeast' | 'exact';
  rowTextId?: string;
}

export interface ImportedDocxTableRowMarkers {
  rows: ImportedDocxTableRowMarker[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const TABLE_ROW_MARKER_PATTERN = /__A3S_WORK_TABLE_ROW_\d+__/g;
const PIXELS_PER_TWIP = 96 / 1440;

export function markDocxTableRows(
  document: Document,
): ImportedDocxTableRowMarkers {
  const wordRows = descendants(document, 'tr').filter((row) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(row.namespaceURI ?? ''),
  );
  const identities = new Map<Element, WorkDocumentTableRowIdentity>();
  const identityCounts = wordIdentityIdCounts(document);
  for (const row of wordRows) {
    const identity = wordTableRowIdentity(row);
    if (!identity) continue;
    identities.set(row, identity);
  }
  const rows: ImportedDocxTableRowMarker[] = [];
  for (const row of wordRows) {
    const properties = directChild(row, 'trPr');
    const cantSplit = properties
      ? directChild(properties, 'cantSplit')
      : undefined;
    const repeatHeader = properties
      ? directChild(properties, 'tblHeader')
      : undefined;
    const height = properties ? directChild(properties, 'trHeight') : undefined;
    const rowHeight = height
      ? twipsToPixels(Number(attribute(height, 'val')))
      : null;
    const rowHeightRule = tableRowHeightRule(attribute(height ?? row, 'hRule'));
    const rowIdentity = identities.get(row);
    const uniqueIdentity =
      rowIdentity && identityCounts.get(rowIdentity.rowId) === 1
        ? rowIdentity
        : null;
    if (!cantSplit && !repeatHeader && rowHeight === null && !uniqueIdentity) {
      continue;
    }
    const paragraph = firstTableRowParagraph(document, row);
    if (!paragraph) continue;
    const marker = `__A3S_WORK_TABLE_ROW_${rows.length + 1}__`;
    insertRowMarker(document, paragraph, marker);
    rows.push({
      marker,
      ...(cantSplit ? { cantSplit: onOffValue(cantSplit) } : {}),
      ...(repeatHeader ? { repeatHeader: onOffValue(repeatHeader) } : {}),
      ...(uniqueIdentity ?? {}),
      ...(rowHeight !== null ? { rowHeight } : {}),
      ...(rowHeight !== null && rowHeightRule ? { rowHeightRule } : {}),
    });
  }
  return { rows };
}

function wordIdentityIdCounts(document: Document): Map<string, number> {
  const counts = new Map<string, number>();
  for (const localName of ['p', 'tr']) {
    for (const element of descendants(document, localName)) {
      if (!DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')) {
        continue;
      }
      const paragraphId = normalizeDocumentParagraphId(
        word2010Attribute(element, 'paraId'),
      );
      if (!paragraphId) continue;
      counts.set(paragraphId, (counts.get(paragraphId) ?? 0) + 1);
    }
  }
  return counts;
}

export function applyImportedDocxTableRowMarkers(
  document: Document,
  markers: ImportedDocxTableRowMarkers,
): void {
  const rowByMarker = new Map(markers.rows.map((row) => [row.marker, row]));
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_TABLE_ROW_')) continue;
    const row = node.parentElement?.closest('tr');
    node.data = node.data.replace(TABLE_ROW_MARKER_PATTERN, (marker) => {
      const properties = rowByMarker.get(marker);
      if (row instanceof HTMLTableRowElement && properties) {
        applyDocumentTableRowIdentityToElement(row, properties);
        setBooleanAttribute(
          row,
          'data-office-cant-split',
          properties.cantSplit,
        );
        setBooleanAttribute(
          row,
          'data-office-repeat-header',
          properties.repeatHeader,
        );
        if (properties.rowHeight !== undefined) {
          row.dataset.officeRowHeight = String(properties.rowHeight);
          row.style.height = `${properties.rowHeight}px`;
          row.dataset.officeRowHeightRule =
            properties.rowHeightRule ?? 'atLeast';
        }
      }
      return '';
    });
  }
  document.body.normalize();
}

export function hasImportedDocxTableRowMarkers(
  markers: ImportedDocxTableRowMarkers,
): boolean {
  return markers.rows.length > 0;
}

function firstTableRowParagraph(
  document: Document,
  row: Element,
): Element | null {
  const cell = directChild(row, 'tc');
  if (!cell) return null;
  const existing =
    directChild(cell, 'p') ??
    descendants(cell, 'p').find(
      (paragraph) => closestAncestor(paragraph, 'tr') === row,
    );
  if (existing) return existing;
  const namespace = cell.namespaceURI ?? WORD_NAMESPACE;
  const prefix = cell.prefix ? `${cell.prefix}:` : '';
  const paragraph = document.createElementNS(namespace, `${prefix}p`);
  cell.append(paragraph);
  return paragraph;
}

function insertRowMarker(
  document: Document,
  paragraph: Element,
  marker: string,
): void {
  const namespace = paragraph.namespaceURI ?? WORD_NAMESPACE;
  const prefix = paragraph.prefix ? `${paragraph.prefix}:` : '';
  const run = document.createElementNS(namespace, `${prefix}r`);
  const text = document.createElementNS(namespace, `${prefix}t`);
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  const properties = directChild(paragraph, 'pPr');
  paragraph.insertBefore(run, properties?.nextSibling ?? paragraph.firstChild);
}

function wordTableRowIdentity(
  row: Element,
): WorkDocumentTableRowIdentity | null {
  return normalizeDocumentTableRowIdentity({
    rowId: word2010Attribute(row, 'paraId'),
    rowTextId: word2010Attribute(row, 'textId'),
  });
}

function word2010Attribute(element: Element, localName: string): string | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === localName &&
        xmlAttributeNamespace(element, item) === WORD_2010_NAMESPACE,
    )?.value ?? null
  );
}

function onOffValue(element: Element): boolean {
  const value = (attribute(element, 'val') ?? attribute(element, 'w:val'))
    ?.trim()
    .toLowerCase();
  return value !== '0' && value !== 'false' && value !== 'off';
}

function setBooleanAttribute(
  element: HTMLElement,
  name: string,
  value: boolean | undefined,
): void {
  if (value !== undefined) element.setAttribute(name, String(value));
}

function tableRowHeightRule(value: string | null): 'atLeast' | 'exact' | null {
  if (value === 'exact') return 'exact';
  return value === 'atLeast' ? 'atLeast' : null;
}

function twipsToPixels(value: number): number | null {
  if (!Number.isFinite(value) || value <= 0) return null;
  return Math.round(value * PIXELS_PER_TWIP * 100) / 100;
}

function closestAncestor(element: Element, localName: string): Element | null {
  let current: Element | null = element;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
}

function textNodes(root: ParentNode): Text[] {
  const document = root.ownerDocument;
  const walker = document?.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}
