import { normalizeTableColor } from './work-document-table-borders';
import {
  parseDocxTblLookElement,
  serializeDocumentTableLookDataset,
  type DocumentTableLook,
} from './work-document-table-look';
import {
  normalizeDocumentTableOverlap,
  normalizeDocumentTableStyleId,
  normalizeDocumentTableCellSpacing,
  normalizeDocumentTableColBandSize,
  normalizeDocumentTableRowBandSize,
  normalizeDocumentTableCaption,
  normalizeDocumentTableDescription,
  type DocumentTableOverlap,
} from './work-document-table-format-changes';
import {
  DEFAULT_DOCUMENT_TABLE_CELL_MARGINS,
  DEFAULT_DOCUMENT_TABLE_GEOMETRY,
  applyDocumentTableGeometryToElement,
  type DocumentTableAlignment,
  type DocumentTableCellMargins,
  type DocumentTableGeometry,
  type DocumentTableLayoutAlgorithm,
  type DocumentTablePreferredWidth,
} from './work-document-table-geometry';
import {
  applyDocumentTableFloatOmmlToElement,
  serializePreservableDocxTableFloat,
} from './work-document-table-float';
import {
  applyDocumentTablePropertyRevisionOmmlToElement,
  serializePreservableDocxTablePropertyRevision,
} from './work-document-table-property-revision';
import {
  applyDocumentTableFormattingChangeToElement,
  importedDocxTableFormattingBorders,
  supportedDocxTableFormattingChangeFromProperties,
  type SupportedDocxTableFormattingChange,
} from './work-docx-table-format-change-import';
import {
  serializeDocumentTableFormattingBordersDataset,
  type DocumentTableFormattingBorders,
} from './work-document-table-formatting-borders';
import {
  docxTablePropertySources,
  resolveDocxTableStyleResolver,
  type DocxTableStyleSource,
} from './work-docx-table-styles';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
} from './work-ooxml-package';

export interface ImportedDocxTableSizingMarker {
  marker: string;
  geometry: DocumentTableGeometry;
  columnWidths: number[];
  columnPercentages: number[];
  floatOmml?: string;
  propertyRevisionOmml?: string;
  formattingChange?: SupportedDocxTableFormattingChange;
  bidiVisual?: boolean;
  fill?: string;
  look?: DocumentTableLook;
  overlap?: DocumentTableOverlap;
  styleId?: string;
  cellSpacing?: number;
  colBandSize?: number;
  rowBandSize?: number;
  borders?: DocumentTableFormattingBorders;
  caption?: string;
  description?: string;
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
  tableStyleSource?: DocxTableStyleSource,
): ImportedDocxTableSizingMarkers {
  const tables: ImportedDocxTableSizingMarker[] = [];
  const tableStyles = resolveDocxTableStyleResolver(tableStyleSource);
  for (const table of descendants(document, 'tbl')) {
    const paragraph = firstTableParagraph(document, table);
    if (!paragraph) continue;
    const grid = directChild(table, 'tblGrid');
    const columnWidths = grid
      ? directChildren(grid, 'gridCol')
          .map((column) => twipsToPixels(Number(attribute(column, 'w')), false))
          .filter((width): width is number => width !== null)
      : [];
    const marker = `__A3S_WORK_TABLE_SIZING_${tables.length + 1}__`;
    insertMarker(document, paragraph, marker);
    const floatOmml = serializePreservableDocxTableFloat(
      directChild(table, 'tblPr'),
    );
    const tableProperties = directChild(table, 'tblPr');
    const formattingChange =
      supportedDocxTableFormattingChangeFromProperties(tableProperties);
    const propertyRevisionOmml = formattingChange
      ? undefined
      : (serializePreservableDocxTablePropertyRevision(tableProperties) ??
        undefined);
    tables.push({
      marker,
      geometry: importedTableGeometry(
        docxTablePropertySources(table, tableStyles),
      ),
      columnWidths,
      columnPercentages: importedColumnPercentages(table, columnWidths),
      ...(floatOmml ? { floatOmml } : {}),
      ...(propertyRevisionOmml ? { propertyRevisionOmml } : {}),
      ...(formattingChange ? { formattingChange } : {}),
      ...(importedTableBidiVisual(tableProperties) ? { bidiVisual: true } : {}),
      ...(importedTableFill(tableProperties)
        ? { fill: importedTableFill(tableProperties)! }
        : {}),
      ...(importedTableLook(tableProperties)
        ? { look: importedTableLook(tableProperties)! }
        : {}),
      ...(importedTableOverlap(tableProperties)
        ? { overlap: importedTableOverlap(tableProperties)! }
        : {}),
      ...(importedTableStyleId(tableProperties)
        ? { styleId: importedTableStyleId(tableProperties)! }
        : {}),
      ...(importedTableCellSpacing(tableProperties) !== null
        ? { cellSpacing: importedTableCellSpacing(tableProperties)! }
        : {}),
      ...(importedTableColBandSize(tableProperties) !== null
        ? { colBandSize: importedTableColBandSize(tableProperties)! }
        : {}),
      ...(importedTableRowBandSize(tableProperties) !== null
        ? { rowBandSize: importedTableRowBandSize(tableProperties)! }
        : {}),
      ...(importedTableBorders(tableProperties)
        ? { borders: importedTableBorders(tableProperties)! }
        : {}),
      ...(importedTableCaption(tableProperties)
        ? { caption: importedTableCaption(tableProperties)! }
        : {}),
      ...(importedTableDescription(tableProperties)
        ? { description: importedTableDescription(tableProperties)! }
        : {}),
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
        table.dataset.officeTableImported = 'true';
        applyDocumentTableGeometryToElement(table, sizing.geometry);
        applyDocumentTableFloatOmmlToElement(table, sizing.floatOmml);
        applyDocumentTablePropertyRevisionOmmlToElement(
          table,
          sizing.propertyRevisionOmml,
        );
        if (sizing.formattingChange) {
          applyDocumentTableFormattingChangeToElement(
            table,
            sizing.formattingChange,
          );
        }
        if (sizing.bidiVisual) {
          table.dataset.officeTableBidiVisual = 'true';
        }
        if (sizing.fill) {
          table.dataset.officeTableFill = sizing.fill;
        }
        if (sizing.look) {
          table.dataset.officeTableLook = serializeDocumentTableLookDataset(
            sizing.look,
          );
        }
        if (sizing.overlap) {
          table.dataset.officeTableOverlap = sizing.overlap;
        }
        if (sizing.styleId) {
          table.dataset.officeTableStyleId = sizing.styleId;
        }
        if (sizing.cellSpacing !== undefined) {
          table.dataset.officeTableCellSpacing = String(sizing.cellSpacing);
        }
        if (sizing.colBandSize !== undefined) {
          table.dataset.officeTableColBandSize = String(sizing.colBandSize);
        }
        if (sizing.rowBandSize !== undefined) {
          table.dataset.officeTableRowBandSize = String(sizing.rowBandSize);
        }
        if (sizing.borders) {
          const encoded = serializeDocumentTableFormattingBordersDataset(
            sizing.borders,
          );
          if (encoded) table.dataset.officeTableBorders = encoded;
        }
        if (sizing.caption) {
          table.dataset.officeTableCaption = sizing.caption;
        }
        if (sizing.description) {
          table.dataset.officeTableDescription = sizing.description;
        }
        if (sizing.columnWidths.length) {
          applyColumnWidths(table, sizing.columnWidths);
        }
        if (sizing.columnPercentages.length) {
          applyColumnPercentages(table, sizing.columnPercentages);
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

function importedTableFill(
  properties: Element | null | undefined,
): string | null {
  if (!properties) return null;
  const element = directChild(properties, 'shd');
  if (!element) return null;
  const fill = attribute(element, 'fill')?.trim() ?? '';
  if (!/^[0-9A-Fa-f]{6}$/.test(fill)) return null;
  const val = attribute(element, 'val');
  if (
    val !== null &&
    val !== '' &&
    !['clear', 'nil', 'none'].includes(val.trim().toLowerCase())
  ) {
    return null;
  }
  return normalizeTableColor(`#${fill}`);
}

function importedTableLook(
  properties: Element | null | undefined,
): DocumentTableLook | null {
  if (!properties) return null;
  const element = directChild(properties, 'tblLook');
  if (!element) return null;
  return parseDocxTblLookElement(element);
}

function importedTableOverlap(
  properties: Element | null | undefined,
): DocumentTableOverlap | null {
  if (!properties) return null;
  const element = directChild(properties, 'tblOverlap');
  if (!element || element.children.length > 0) return null;
  return normalizeDocumentTableOverlap(attribute(element, 'val'));
}

function importedTableStyleId(
  properties: Element | null | undefined,
): string | null {
  if (!properties) return null;
  const element = directChild(properties, 'tblStyle');
  if (!element || element.children.length > 0) return null;
  return normalizeDocumentTableStyleId(attribute(element, 'val'));
}

function importedTableCellSpacing(
  properties: Element | null | undefined,
): number | null {
  if (!properties) return null;
  const element = directChild(properties, 'tblCellSpacing');
  if (!element || element.children.length > 0) return null;
  const type = attribute(element, 'type');
  if (type !== null && type !== 'dxa') return null;
  const twips = Number(attribute(element, 'w'));
  if (!Number.isFinite(twips) || twips < 0) return null;
  return normalizeDocumentTableCellSpacing(
    Math.round(twips * PIXELS_PER_TWIP * 100) / 100,
  );
}

function importedTableColBandSize(
  properties: Element | null | undefined,
): number | null {
  if (!properties) return null;
  const element = directChild(properties, 'tblStyleColBandSize');
  if (!element || element.children.length > 0) return null;
  return normalizeDocumentTableColBandSize(attribute(element, 'val'));
}

function importedTableRowBandSize(
  properties: Element | null | undefined,
): number | null {
  if (!properties) return null;
  const element = directChild(properties, 'tblStyleRowBandSize');
  if (!element || element.children.length > 0) return null;
  return normalizeDocumentTableRowBandSize(attribute(element, 'val'));
}

function importedTableBorders(
  properties: Element | null | undefined,
): DocumentTableFormattingBorders | null {
  if (!properties) return null;
  const element = directChild(properties, 'tblBorders');
  if (!element) return null;
  return importedDocxTableFormattingBorders(element);
}

function importedTableCaption(
  properties: Element | null | undefined,
): string | null {
  if (!properties) return null;
  const element = directChild(properties, 'tblCaption');
  if (!element || element.children.length > 0) return null;
  return normalizeDocumentTableCaption(attribute(element, 'val'));
}

function importedTableDescription(
  properties: Element | null | undefined,
): string | null {
  if (!properties) return null;
  const element = directChild(properties, 'tblDescription');
  if (!element || element.children.length > 0) return null;
  return normalizeDocumentTableDescription(attribute(element, 'val'));
}

function importedTableBidiVisual(
  properties: Element | null | undefined,
): boolean {
  if (!properties) return false;
  const element = directChild(properties, 'bidiVisual');
  if (!element) return false;
  const value = attribute(element, 'val');
  if (value === null || value === '') return true;
  const normalized = value.trim().toLowerCase();
  return (
    normalized === '1' ||
    normalized === 'true' ||
    normalized === 'on' ||
    normalized === 'yes'
  );
}

function importedTableGeometry(
  propertySources: readonly Element[],
): DocumentTableGeometry {
  let layout: DocumentTableLayoutAlgorithm =
    DEFAULT_DOCUMENT_TABLE_GEOMETRY.layout;
  let width: DocumentTablePreferredWidth = { type: 'auto', value: null };
  let alignment: DocumentTableAlignment =
    DEFAULT_DOCUMENT_TABLE_GEOMETRY.alignment;
  let indent = DEFAULT_DOCUMENT_TABLE_GEOMETRY.indent;
  let cellMargins: DocumentTableCellMargins = {
    ...DEFAULT_DOCUMENT_TABLE_CELL_MARGINS,
  };
  for (const properties of propertySources) {
    const layoutElement = directChild(properties, 'tblLayout');
    const layoutValue = layoutElement
      ? importedTableLayout(attribute(layoutElement, 'type'))
      : null;
    if (layoutValue) layout = layoutValue;

    const widthElement = directChild(properties, 'tblW');
    const widthValue = widthElement
      ? importedTablePreferredWidth(widthElement)
      : null;
    if (widthValue) width = widthValue;

    const alignmentElement = directChild(properties, 'jc');
    const alignmentValue = alignmentElement
      ? importedTableAlignment(attribute(alignmentElement, 'val'))
      : null;
    if (alignmentValue) alignment = alignmentValue;

    const indentElement = directChild(properties, 'tblInd');
    const indentValue = indentElement
      ? tableWidthPixels(indentElement, true)
      : null;
    if (indentValue !== null) indent = indentValue;

    const marginElement = directChild(properties, 'tblCellMar');
    if (marginElement) {
      cellMargins = {
        ...cellMargins,
        ...importedCellMargins(marginElement),
      };
    }
  }
  return { layout, width, alignment, indent, cellMargins };
}

function importedTableLayout(
  value: string | null,
): DocumentTableLayoutAlgorithm | null {
  if (value === 'fixed') return 'fixed';
  return value === 'autofit' ? 'autofit' : null;
}

function importedTablePreferredWidth(
  width: Element,
): DocumentTablePreferredWidth | null {
  const type = attribute(width, 'type');
  if (type === 'auto' || type === 'nil') return { type: 'auto', value: null };
  if (type === 'pct') {
    const value = percentageValue(attribute(width, 'w'));
    return value === null ? null : { type: 'percent', value };
  }
  if (type === 'dxa') {
    const value = tableWidthPixels(width, false);
    return value === null ? null : { type: 'pixels', value };
  }
  return null;
}

function importedTableAlignment(
  value: string | null,
): DocumentTableAlignment | null {
  if (value === 'center') return 'center';
  if (value === 'right' || value === 'end') return 'right';
  if (value === 'left' || value === 'start') return 'left';
  return null;
}

function importedCellMargins(
  margins: Element,
): Partial<DocumentTableCellMargins> {
  const children = new Map(
    directChildren(margins).map((margin) => [margin.localName, margin]),
  );
  const result: Partial<DocumentTableCellMargins> = {};
  assignMargin(result, 'top', children.get('top'));
  assignMargin(result, 'right', children.get('right') ?? children.get('end'));
  assignMargin(result, 'bottom', children.get('bottom'));
  assignMargin(result, 'left', children.get('left') ?? children.get('start'));
  return result;
}

function assignMargin(
  target: Partial<DocumentTableCellMargins>,
  side: keyof DocumentTableCellMargins,
  element: Element | undefined,
): void {
  if (!element) return;
  const value = tableWidthPixels(element, true);
  if (value !== null) target[side] = value;
}

function tableWidthPixels(element: Element, allowZero: boolean): number | null {
  const type = attribute(element, 'type');
  if (type && type !== 'dxa') return null;
  return twipsToPixels(Number(attribute(element, 'w')), allowZero);
}

function percentageValue(value: string | null): number | null {
  const normalized = value?.trim();
  if (!normalized) return null;
  const percentage = normalized.endsWith('%')
    ? Number(normalized.slice(0, -1))
    : Number(normalized) / 50;
  if (!Number.isFinite(percentage) || percentage <= 0) return null;
  return Math.round(percentage * 100) / 100;
}

function applyColumnWidths(
  table: HTMLTableElement,
  columnWidths: readonly number[],
): void {
  const occupiedUntilRow: number[] = [];
  ownedTableRows(table).forEach((row, rowIndex) => {
    let columnIndex = 0;
    for (const cell of ownedTableRowCells(row)) {
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

function applyColumnPercentages(
  table: HTMLTableElement,
  columnPercentages: readonly number[],
): void {
  const occupiedUntilRow: number[] = [];
  ownedTableRows(table).forEach((row, rowIndex) => {
    let columnIndex = 0;
    for (const cell of ownedTableRowCells(row)) {
      while ((occupiedUntilRow[columnIndex] ?? 0) > rowIndex) columnIndex += 1;
      const percentages = columnPercentages.slice(
        columnIndex,
        columnIndex + cell.colSpan,
      );
      if (percentages.length === cell.colSpan) {
        cell.dataset.officeColumnWidthsPercent = percentages.join(',');
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

function ownedTableRows(table: HTMLTableElement): HTMLTableRowElement[] {
  return Array.from(table.rows).filter((row) => row.closest('table') === table);
}

function ownedTableRowCells(row: HTMLTableRowElement): HTMLTableCellElement[] {
  return Array.from(row.cells).filter((cell) => cell.closest('tr') === row);
}

function importedColumnPercentages(
  table: Element,
  columnWidths: readonly number[],
): number[] {
  const columnCount = columnWidths.length || logicalTableColumnCount(table);
  const percentages: Array<number | null> = Array.from(
    { length: columnCount },
    () => null,
  );
  const rows = directChildren(table, 'tr');
  for (const row of rows) {
    let column = rowGridBefore(row);
    for (const cell of directChildren(row, 'tc')) {
      const properties = directChild(cell, 'tcPr');
      const gridSpan = directChild(properties ?? cell, 'gridSpan');
      const span = Math.max(
        1,
        Number(gridSpan ? attribute(gridSpan, 'val') : null) || 1,
      );
      const width = directChild(properties ?? cell, 'tcW');
      const total =
        width && attribute(width, 'type') === 'pct'
          ? percentageValue(attribute(width, 'w'))
          : null;
      if (total !== null) {
        const physical = columnWidths.slice(column, column + span);
        const physicalTotal = physical.reduce((sum, value) => sum + value, 0);
        for (let offset = 0; offset < span; offset += 1) {
          if (percentages[column + offset] !== null) continue;
          const share =
            physicalTotal > 0
              ? total * ((physical[offset] ?? 0) / physicalTotal)
              : total / span;
          percentages[column + offset] = Math.round(share * 100) / 100;
        }
      }
      column += span;
    }
  }
  return percentages.every((value) => value !== null && value > 0)
    ? (percentages as number[])
    : [];
}

function logicalTableColumnCount(table: Element): number {
  return directChildren(table, 'tr').reduce((maximum, row) => {
    const count = directChildren(row, 'tc').reduce((sum, cell) => {
      const properties = directChild(cell, 'tcPr');
      const gridSpan = directChild(properties ?? cell, 'gridSpan');
      const span = Number(gridSpan ? attribute(gridSpan, 'val') : null);
      return sum + (Number.isSafeInteger(span) && span > 0 ? span : 1);
    }, rowGridBefore(row));
    return Math.max(maximum, count);
  }, 0);
}

function rowGridBefore(row: Element): number {
  const properties = directChild(row, 'trPr');
  const gridBefore = directChild(properties ?? row, 'gridBefore');
  const value = Number(gridBefore ? attribute(gridBefore, 'val') : null);
  return Number.isSafeInteger(value) && value > 0 ? value : 0;
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

function twipsToPixels(value: number, allowZero: boolean): number | null {
  if (!Number.isFinite(value) || value < 0 || (!allowZero && value === 0)) {
    return null;
  }
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
