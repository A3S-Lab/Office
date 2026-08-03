import {
  normalizeDocumentTableVerticalAlign,
  type DocumentTableVerticalAlign,
} from './work-document-table-cell-formatting';
import {
  normalizeDocumentTableCellMarginOverrides,
  normalizeDocumentTableProperties,
  type DocumentTableCellMarginOverrides,
  type DocumentTableProperties,
} from './work-document-table-geometry';
import type { DocumentTableRowHeightRule } from './work-document-table-row';

export const MIN_DOCUMENT_TABLE_COLUMN_WIDTH = 25;
export const MIN_DOCUMENT_TABLE_ROW_HEIGHT = 12;
export const MAX_DOCUMENT_TABLE_DIMENSION = 4_000;

export interface DocumentTablePropertyChanges {
  table?: DocumentTableProperties;
  row?: {
    height: number | null;
    heightRule: DocumentTableRowHeightRule;
    cantSplit: boolean;
    repeatHeader?: boolean;
  };
  column?: {
    width: number;
    renderedColumnWidths?: readonly number[];
  };
  cell?: {
    verticalAlign: DocumentTableVerticalAlign;
    margins: DocumentTableCellMarginOverrides | null;
  };
}

export function normalizeDocumentTablePropertyChanges(
  value: unknown,
): DocumentTablePropertyChanges | null {
  if (!isRecord(value)) return null;
  const changes: DocumentTablePropertyChanges = {};

  if (value.table !== undefined) {
    const table = normalizeDocumentTableProperties(value.table);
    if (!table) return null;
    changes.table = table;
  }

  if (value.row !== undefined) {
    const row = value.row;
    if (!isRecord(row)) return null;
    if (row.height !== null && typeof row.height !== 'number') return null;
    const height =
      row.height === null
        ? null
        : normalizeDocumentTableDimension(
            row.height,
            MIN_DOCUMENT_TABLE_ROW_HEIGHT,
          );
    if (row.height !== null && height === null) return null;
    if (row.heightRule !== 'atLeast' && row.heightRule !== 'exact') {
      return null;
    }
    if (typeof row.cantSplit !== 'boolean') return null;
    if (
      row.repeatHeader !== undefined &&
      typeof row.repeatHeader !== 'boolean'
    ) {
      return null;
    }
    changes.row = {
      height,
      heightRule: row.heightRule,
      cantSplit: row.cantSplit,
      ...(row.repeatHeader === undefined
        ? {}
        : { repeatHeader: row.repeatHeader }),
    };
  }

  if (value.column !== undefined) {
    const column = value.column;
    if (!isRecord(column) || typeof column.width !== 'number') return null;
    const width = normalizeDocumentTableDimension(
      column.width,
      MIN_DOCUMENT_TABLE_COLUMN_WIDTH,
    );
    if (width === null) return null;
    changes.column = {
      width,
      ...(Array.isArray(column.renderedColumnWidths)
        ? { renderedColumnWidths: column.renderedColumnWidths }
        : {}),
    };
  }

  if (value.cell !== undefined) {
    const cell = value.cell;
    if (!isRecord(cell) || typeof cell.verticalAlign !== 'string') return null;
    const verticalAlign = normalizeDocumentTableVerticalAlign(
      cell.verticalAlign,
    );
    const margins =
      cell.margins === null
        ? null
        : normalizeDocumentTableCellMarginOverrides(cell.margins);
    if (!verticalAlign || (cell.margins !== null && !margins)) return null;
    changes.cell = { verticalAlign, margins };
  }

  return Object.keys(changes).length ? changes : null;
}

export function normalizeDocumentTableDimension(
  value: number,
  minimum: number,
): number | null {
  if (
    !Number.isFinite(value) ||
    value < minimum ||
    value > MAX_DOCUMENT_TABLE_DIMENSION
  ) {
    return null;
  }
  return Math.round(value * 100) / 100;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}
