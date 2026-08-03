import type { ITableCellOptions, ITableOptions, ITableRowOptions } from 'docx';
import { normalizeDocumentTableRowHeightRule } from './work-document-table-row';
import { documentTableGeometryFromElement } from './work-document-table-geometry';

type TableSizingDocxOptions = Pick<
  ITableOptions,
  'alignment' | 'columnWidths' | 'indent' | 'layout' | 'margins' | 'width'
>;
type TableCellSizingDocxOptions = Pick<ITableCellOptions, 'width'>;
type TableRowSizingDocxOptions = Pick<ITableRowOptions, 'height'>;

const TWIPS_PER_PIXEL = 1440 / 96;

export function documentTableSizingDocxOptions(
  table: HTMLTableElement,
  docx: typeof import('docx'),
): TableSizingDocxOptions {
  const geometry = documentTableGeometryFromElement(table);
  const pixelWidths = documentTableColumnWidths(table);
  const columnWidths = pixelWidths?.map(pixelsToTwips);
  return {
    ...(columnWidths?.length ? { columnWidths } : {}),
    layout:
      geometry.layout === 'fixed'
        ? docx.TableLayoutType.FIXED
        : docx.TableLayoutType.AUTOFIT,
    width:
      geometry.width.type === 'percent' && geometry.width.value !== null
        ? {
            size: geometry.width.value,
            type: docx.WidthType.PERCENTAGE,
          }
        : geometry.width.type === 'pixels' && geometry.width.value !== null
          ? {
              size: pixelsToTwips(geometry.width.value),
              type: docx.WidthType.DXA,
            }
          : { size: 0, type: docx.WidthType.AUTO },
    alignment:
      geometry.alignment === 'center'
        ? docx.AlignmentType.CENTER
        : geometry.alignment === 'right'
          ? docx.AlignmentType.RIGHT
          : docx.AlignmentType.LEFT,
    ...(geometry.indent > 0
      ? {
          indent: {
            size: pixelsToTwips(geometry.indent),
            type: docx.WidthType.DXA,
          },
        }
      : {}),
    margins: {
      top: pixelsToNonNegativeTwips(geometry.cellMargins.top),
      right: pixelsToNonNegativeTwips(geometry.cellMargins.right),
      bottom: pixelsToNonNegativeTwips(geometry.cellMargins.bottom),
      left: pixelsToNonNegativeTwips(geometry.cellMargins.left),
    },
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

function pixelsToNonNegativeTwips(pixels: number): number {
  return Math.max(0, Math.round(pixels * TWIPS_PER_PIXEL));
}
