import { attribute, descendants, directChild } from './work-ooxml-package';

export interface ImportedDocxTableRowMarker {
  marker: string;
  cantSplit?: boolean;
  repeatHeader?: boolean;
}

export interface ImportedDocxTableRowMarkers {
  rows: ImportedDocxTableRowMarker[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const TABLE_ROW_MARKER_PATTERN = /__A3S_WORK_TABLE_ROW_\d+__/g;

export function markDocxTableRows(
  document: Document,
): ImportedDocxTableRowMarkers {
  const rows: ImportedDocxTableRowMarker[] = [];
  for (const row of descendants(document, 'tr')) {
    const properties = directChild(row, 'trPr');
    const cantSplit = properties
      ? directChild(properties, 'cantSplit')
      : undefined;
    const repeatHeader = properties
      ? directChild(properties, 'tblHeader')
      : undefined;
    if (!cantSplit && !repeatHeader) continue;
    const paragraph = firstTableRowParagraph(document, row);
    if (!paragraph) continue;
    const marker = `__A3S_WORK_TABLE_ROW_${rows.length + 1}__`;
    insertRowMarker(document, paragraph, marker);
    rows.push({
      marker,
      ...(cantSplit ? { cantSplit: onOffValue(cantSplit) } : {}),
      ...(repeatHeader ? { repeatHeader: onOffValue(repeatHeader) } : {}),
    });
  }
  return { rows };
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
  const paragraph = document.createElementNS(WORD_NAMESPACE, 'w:p');
  cell.append(paragraph);
  return paragraph;
}

function insertRowMarker(
  document: Document,
  paragraph: Element,
  marker: string,
): void {
  const run = document.createElementNS(WORD_NAMESPACE, 'w:r');
  const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  const properties = directChild(paragraph, 'pPr');
  paragraph.insertBefore(run, properties?.nextSibling ?? paragraph.firstChild);
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
