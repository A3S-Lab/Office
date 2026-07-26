import {
  attribute,
  descendants,
  directChild,
  directChildren,
} from './work-ooxml-package';
import type { DocumentTableLayoutMode } from './work-document-table-sizing';

export interface ImportedDocxTableSizingMarker {
  marker: string;
  layoutMode: DocumentTableLayoutMode;
  columnWidths: number[];
}

export interface ImportedDocxTableSizingMarkers {
  tables: ImportedDocxTableSizingMarker[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const TABLE_SIZING_MARKER_PATTERN = /__A3S_WORK_TABLE_SIZING_\d+__/g;
const PIXELS_PER_TWIP = 96 / 1440;

export function markDocxTableSizing(
  document: Document,
): ImportedDocxTableSizingMarkers {
  const tables: ImportedDocxTableSizingMarker[] = [];
  for (const table of descendants(document, 'tbl')) {
    const paragraph = firstTableParagraph(document, table);
    if (!paragraph) continue;
    const properties = directChild(table, 'tblPr');
    const layout = properties
      ? directChild(properties, 'tblLayout')
      : undefined;
    const preferredWidth = properties
      ? directChild(properties, 'tblW')
      : undefined;
    const grid = directChild(table, 'tblGrid');
    const columnWidths = grid
      ? directChildren(grid, 'gridCol')
          .map((column) => twipsToPixels(Number(attribute(column, 'w'))))
          .filter((width): width is number => width !== null)
      : [];
    const marker = `__A3S_WORK_TABLE_SIZING_${tables.length + 1}__`;
    insertMarker(document, paragraph, marker);
    tables.push({
      marker,
      layoutMode: importedTableLayoutMode(layout, preferredWidth),
      columnWidths,
    });
  }
  return { tables };
}

export function applyImportedDocxTableSizingMarkers(
  document: Document,
  markers: ImportedDocxTableSizingMarkers,
): void {
  const tableByMarker = new Map(
    markers.tables.map((table) => [table.marker, table]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_TABLE_SIZING_')) continue;
    const table = node.parentElement?.closest('table');
    node.data = node.data.replace(TABLE_SIZING_MARKER_PATTERN, (marker) => {
      const sizing = tableByMarker.get(marker);
      if (table instanceof HTMLTableElement && sizing) {
        table.dataset.officeTableLayout = sizing.layoutMode;
        if (sizing.layoutMode === 'fixed' && sizing.columnWidths.length) {
          applyColumnWidths(table, sizing.columnWidths);
        }
      }
      return '';
    });
  }
  document.body.normalize();
}

export function hasImportedDocxTableSizingMarkers(
  markers: ImportedDocxTableSizingMarkers,
): boolean {
  return markers.tables.length > 0;
}

function importedTableLayoutMode(
  layout: Element | undefined,
  preferredWidth: Element | undefined,
): DocumentTableLayoutMode {
  const layoutType = layout ? attribute(layout, 'type') : null;
  const widthType = preferredWidth ? attribute(preferredWidth, 'type') : null;
  if (layoutType === 'autofit' || widthType === 'auto') return 'contents';
  if (widthType === 'pct') return 'window';
  return 'fixed';
}

function applyColumnWidths(
  table: HTMLTableElement,
  columnWidths: readonly number[],
): void {
  const occupiedUntilRow: number[] = [];
  Array.from(table.rows).forEach((row, rowIndex) => {
    let columnIndex = 0;
    for (const cell of Array.from(row.cells)) {
      while ((occupiedUntilRow[columnIndex] ?? 0) > rowIndex) {
        columnIndex += 1;
      }
      const widths = columnWidths.slice(
        columnIndex,
        columnIndex + cell.colSpan,
      );
      if (widths.length === cell.colSpan) {
        cell.setAttribute('colwidth', widths.join(','));
      }
      if (cell.rowSpan > 1) {
        for (let offset = 0; offset < cell.colSpan; offset += 1) {
          occupiedUntilRow[columnIndex + offset] = rowIndex + cell.rowSpan;
        }
      }
      columnIndex += cell.colSpan;
    }
  });
}

function firstTableParagraph(
  document: Document,
  table: Element,
): Element | null {
  const cell = descendants(table, 'tc').find(
    (candidate) => closestAncestor(candidate, 'tbl') === table,
  );
  if (!cell) return null;
  const existing = descendants(cell, 'p').find(
    (paragraph) => closestAncestor(paragraph, 'tc') === cell,
  );
  if (existing) return existing;
  const paragraph = document.createElementNS(WORD_NAMESPACE, 'w:p');
  cell.append(paragraph);
  return paragraph;
}

function insertMarker(
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
