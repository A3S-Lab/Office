import type { Cell, CellMatrix } from '@fortune-sheet/core';
import { sparseArrayIndexes } from '../spreadsheet-sparse';
import type { SpreadsheetVirtualGridViewport } from './spreadsheet-virtual-grid-viewport';
import {
  SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH,
  SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT,
} from './spreadsheet-virtual-grid-viewport';

export interface SpreadsheetVirtualGridPaintCell {
  row: number;
  column: number;
  text: string;
  background?: string;
  color?: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * Collect populated cells inside a viewport window for canvas paint.
 * Empty sparse holes are skipped so max-dimension sheets stay allocation-bounded.
 */
export function spreadsheetVirtualGridPaintCells(options: {
  data: CellMatrix | undefined;
  viewport: SpreadsheetVirtualGridViewport;
  rowHeight?: number;
  columnWidth?: number;
  merges?: readonly SpreadsheetVirtualGridMerge[];
}): SpreadsheetVirtualGridPaintCell[] {
  const data = options.data;
  if (!data) return [];
  const { viewport } = options;
  if (viewport.rowEnd <= viewport.rowStart || viewport.columnEnd <= viewport.columnStart) {
    return [];
  }

  const rowHeight =
    options.rowHeight && options.rowHeight > 0
      ? options.rowHeight
      : SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT;
  const columnWidth =
    options.columnWidth && options.columnWidth > 0
      ? options.columnWidth
      : SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH;

  const cells: SpreadsheetVirtualGridPaintCell[] = [];
  for (const rowIndex of sparseArrayIndexes(data)) {
    if (rowIndex < viewport.rowStart || rowIndex >= viewport.rowEnd) continue;
    const row = data[rowIndex];
    if (!row) continue;
    for (const columnIndex of sparseArrayIndexes(row)) {
      if (
        columnIndex < viewport.columnStart ||
        columnIndex >= viewport.columnEnd
      ) {
        continue;
      }
      const cell = row[columnIndex];
      const text = spreadsheetVirtualGridCellText(cell);
      const background =
        typeof cell?.bg === 'string' && cell.bg.trim() ? cell.bg : undefined;
      const color =
        typeof cell?.fc === 'string' && cell.fc.trim() ? cell.fc : undefined;
      if (!text && !background) continue;
      cells.push({
        row: rowIndex,
        column: columnIndex,
        text,
        background,
        color,
        x:
          (columnIndex - viewport.columnStart) * columnWidth -
          viewport.offsetLeft,
        y: (rowIndex - viewport.rowStart) * rowHeight - viewport.offsetTop,
        width: columnWidth,
        height: rowHeight,
      });
    }
  }
  return applySpreadsheetVirtualGridMerges(
    cells,
    options.merges ?? [],
    viewport,
    rowHeight,
    columnWidth,
  );
}

export interface SpreadsheetVirtualGridMerge {
  r: number;
  c: number;
  rs: number;
  cs: number;
}

function applySpreadsheetVirtualGridMerges(
  cells: SpreadsheetVirtualGridPaintCell[],
  merges: readonly SpreadsheetVirtualGridMerge[],
  viewport: SpreadsheetVirtualGridViewport,
  rowHeight: number,
  columnWidth: number,
): SpreadsheetVirtualGridPaintCell[] {
  if (merges.length === 0) return cells;
  const covered = new Set<string>();
  const spans = new Map<string, SpreadsheetVirtualGridMerge>();
  for (const merge of merges) {
    if (
      !Number.isInteger(merge.r) ||
      !Number.isInteger(merge.c) ||
      merge.rs < 1 ||
      merge.cs < 1
    ) {
      continue;
    }
    const rowEnd = merge.r + merge.rs;
    const columnEnd = merge.c + merge.cs;
    if (
      rowEnd <= viewport.rowStart ||
      merge.r >= viewport.rowEnd ||
      columnEnd <= viewport.columnStart ||
      merge.c >= viewport.columnEnd
    ) {
      continue;
    }
    spans.set(`${merge.r}:${merge.c}`, merge);
    for (let row = merge.r; row < rowEnd; row += 1) {
      for (let column = merge.c; column < columnEnd; column += 1) {
        if (row === merge.r && column === merge.c) continue;
        covered.add(`${row}:${column}`);
      }
    }
  }
  if (spans.size === 0) return cells;

  return cells.flatMap((cell) => {
    const key = `${cell.row}:${cell.column}`;
    if (covered.has(key)) return [];
    const merge = spans.get(key);
    if (!merge) return [cell];
    const visibleRowStart = Math.max(merge.r, viewport.rowStart);
    const visibleColumnStart = Math.max(merge.c, viewport.columnStart);
    const visibleRowEnd = Math.min(merge.r + merge.rs, viewport.rowEnd);
    const visibleColumnEnd = Math.min(merge.c + merge.cs, viewport.columnEnd);
    return [
      {
        ...cell,
        x:
          (visibleColumnStart - viewport.columnStart) * columnWidth -
          viewport.offsetLeft,
        y:
          (visibleRowStart - viewport.rowStart) * rowHeight -
          viewport.offsetTop,
        width: (visibleColumnEnd - visibleColumnStart) * columnWidth,
        height: (visibleRowEnd - visibleRowStart) * rowHeight,
      },
    ];
  });
}

export function spreadsheetVirtualGridCellText(
  cell: Cell | null | undefined,
): string {
  if (!cell) return '';
  if (typeof cell.m === 'string' && cell.m.length > 0) return cell.m;
  if (typeof cell.v === 'string') return cell.v;
  if (typeof cell.v === 'number' && Number.isFinite(cell.v)) {
    return String(cell.v);
  }
  if (typeof cell.f === 'string' && cell.f.trim().length > 0) return cell.f;
  return '';
}
