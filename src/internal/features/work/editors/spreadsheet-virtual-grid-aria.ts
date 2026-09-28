/**
 * ARIA contract for the A3S-owned spreadsheet virtual grid.
 * Fortune canvas has no grid semantics; this module owns the accessible name
 * and activedescendant identity the virtual grid surface must expose.
 */

export interface SpreadsheetVirtualGridAria {
  role: 'grid';
  'aria-label': string;
  'aria-rowcount': number;
  'aria-colcount': number;
  'aria-activedescendant': string;
}

export function spreadsheetVirtualGridActiveCellId(
  row: number,
  column: number,
): string {
  return `a3s-ss-cell-${row}-${column}`;
}

export function spreadsheetVirtualGridAria(options: {
  sheetName: string;
  rowCount: number;
  columnCount: number;
  activeRow: number;
  activeColumn: number;
}): SpreadsheetVirtualGridAria {
  const rowCount = Math.max(0, Math.floor(options.rowCount));
  const columnCount = Math.max(0, Math.floor(options.columnCount));
  const activeRow = clampIndex(options.activeRow, rowCount);
  const activeColumn = clampIndex(options.activeColumn, columnCount);

  return {
    role: 'grid',
    'aria-label': options.sheetName.trim() || 'Sheet',
    'aria-rowcount': rowCount,
    'aria-colcount': columnCount,
    'aria-activedescendant': spreadsheetVirtualGridActiveCellId(
      activeRow,
      activeColumn,
    ),
  };
}

function clampIndex(value: number, count: number): number {
  if (count <= 0) return 0;
  if (!Number.isFinite(value)) return 0;
  return Math.min(count - 1, Math.max(0, Math.floor(value)));
}
