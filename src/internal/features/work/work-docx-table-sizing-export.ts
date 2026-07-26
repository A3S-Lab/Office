import type { ITableCellOptions, ITableOptions, ITableRowOptions } from 'docx';
import { normalizeDocumentTableRowHeightRule } from './work-document-table-row';
import { normalizeDocumentTableLayoutMode } from './work-document-table-sizing';

type TableSizingDocxOptions = Pick<
  ITableOptions,
  'columnWidths' | 'layout' | 'width'
>;
type TableCellSizingDocxOptions = Pick<ITableCellOptions, 'width'>;
type TableRowSizingDocxOptions = Pick<ITableRowOptions, 'height'>;

const TWIPS_PER_PIXEL = 1440 / 96;

export function documentTableSizingDocxOptions(
  table: HTMLTableElement,
  docx: typeof import('docx'),
): TableSizingDocxOptions {
  const mode = normalizeDocumentTableLayoutMode(
    table.dataset.officeTableLayout,
  );
  const pixelWidths = documentTableColumnWidths(table);
  const columnWidths = pixelWidths?.map(pixelsToTwips);
  if (mode === 'contents') {
    return {
      layout: docx.TableLayoutType.AUTOFIT,
      width: { size: 0, type: docx.WidthType.AUTO },
    };
  }
  if (mode === 'fixed' && columnWidths?.length) {
    return {
      columnWidths,
      layout: docx.TableLayoutType.FIXED,
      width: {
        size: columnWidths.reduce((sum, width) => sum + width, 0),
        type: docx.WidthType.DXA,
      },
    };
  }
  return {
    layout: docx.TableLayoutType.FIXED,
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
  };
}

export function documentTableCellSizingDocxOptions(
  cell: HTMLTableCellElement,
  docx: typeof import('docx'),
): TableCellSizingDocxOptions {
  const widths = cellColumnWidths(cell);
  if (!widths?.length) return {};
  return {
    width: {
      size: pixelsToTwips(widths.reduce((sum, width) => sum + width, 0)),
      type: docx.WidthType.DXA,
    },
  };
}

export function documentTableRowSizingDocxOptions(
  row: HTMLTableRowElement,
  docx: typeof import('docx'),
): TableRowSizingDocxOptions {
  const pixels = rowHeightPixels(row);
  if (pixels === null) return {};
  const rule = normalizeDocumentTableRowHeightRule(
    row.dataset.officeRowHeightRule,
  );
  return {
    height: {
      value: pixelsToTwips(pixels),
      rule: rule === 'exact' ? docx.HeightRule.EXACT : docx.HeightRule.ATLEAST,
    },
  };
}

function documentTableColumnWidths(table: HTMLTableElement): number[] | null {
  const row = table.rows[0];
  if (!row) return null;
  const widths: number[] = [];
  for (const cell of Array.from(row.cells)) {
    const cellWidths = cellColumnWidths(cell);
    if (!cellWidths || cellWidths.length !== cell.colSpan) return null;
    widths.push(...cellWidths);
  }
  return widths.length ? widths : null;
}

function cellColumnWidths(cell: HTMLTableCellElement): number[] | null {
  const raw = cell.getAttribute('colwidth') ?? cell.dataset.colwidth;
  if (!raw) return null;
  const widths = raw
    .split(',')
    .map((width) => Number(width))
    .filter((width) => Number.isFinite(width) && width > 0);
  return widths.length === cell.colSpan ? widths : null;
}

function rowHeightPixels(row: HTMLTableRowElement): number | null {
  const value = Number.parseFloat(
    row.dataset.officeRowHeight || row.style.height,
  );
  return Number.isFinite(value) && value > 0 ? value : null;
}

function pixelsToTwips(pixels: number): number {
  return Math.max(1, Math.round(pixels * TWIPS_PER_PIXEL));
}
