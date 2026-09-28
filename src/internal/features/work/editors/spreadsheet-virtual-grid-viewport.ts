/**
 * Viewport window for the A3S-owned spreadsheet virtual grid.
 * Computes the sparse row/column range that must paint for the current scroll
 * without materializing empty sheet cells.
 */

export const SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT = 24;
export const SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH = 96;

export interface SpreadsheetVirtualGridViewport {
  /** Inclusive first row in the paint window (may include overscan). */
  rowStart: number;
  /** Exclusive last row in the paint window. */
  rowEnd: number;
  /** Inclusive first column in the paint window (may include overscan). */
  columnStart: number;
  /** Exclusive last column in the paint window. */
  columnEnd: number;
  /**
   * Pixels of `rowStart` clipped above the scroll origin. Paint row `r` at
   * `(r - rowStart) * rowHeight - offsetTop`.
   */
  offsetTop: number;
  /**
   * Pixels of `columnStart` clipped left of the scroll origin. Paint column
   * `c` at `(c - columnStart) * columnWidth - offsetLeft`.
   */
  offsetLeft: number;
}

export interface SpreadsheetVirtualGridViewportOptions {
  scrollTop: number;
  scrollLeft: number;
  clientHeight: number;
  clientWidth: number;
  rowCount: number;
  columnCount: number;
  rowHeight?: number;
  columnWidth?: number;
  /** Extra rows/columns beyond the visible window. Defaults to 0. */
  overscan?: number;
}

export function spreadsheetVirtualGridViewport(
  options: SpreadsheetVirtualGridViewportOptions,
): SpreadsheetVirtualGridViewport {
  const rowCount = Math.max(0, Math.floor(options.rowCount));
  const columnCount = Math.max(0, Math.floor(options.columnCount));
  if (rowCount === 0 || columnCount === 0) {
    return {
      rowStart: 0,
      rowEnd: 0,
      columnStart: 0,
      columnEnd: 0,
      offsetTop: 0,
      offsetLeft: 0,
    };
  }

  const rowHeight = positiveSize(
    options.rowHeight,
    SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT,
  );
  const columnWidth = positiveSize(
    options.columnWidth,
    SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH,
  );
  const overscan = Math.max(0, Math.floor(options.overscan ?? 0));
  const scrollTop = Math.max(0, options.scrollTop);
  const scrollLeft = Math.max(0, options.scrollLeft);
  const clientHeight = Math.max(0, options.clientHeight);
  const clientWidth = Math.max(0, options.clientWidth);

  const visibleRowStart = Math.min(
    rowCount - 1,
    Math.floor(scrollTop / rowHeight),
  );
  const visibleColumnStart = Math.min(
    columnCount - 1,
    Math.floor(scrollLeft / columnWidth),
  );
  const visibleRowEnd = Math.min(
    rowCount,
    Math.ceil((scrollTop + clientHeight) / rowHeight),
  );
  const visibleColumnEnd = Math.min(
    columnCount,
    Math.ceil((scrollLeft + clientWidth) / columnWidth),
  );

  const rowStart = Math.max(0, visibleRowStart - overscan);
  const columnStart = Math.max(0, visibleColumnStart - overscan);
  const rowEnd = Math.min(rowCount, visibleRowEnd + overscan);
  const columnEnd = Math.min(columnCount, visibleColumnEnd + overscan);

  return {
    rowStart,
    rowEnd,
    columnStart,
    columnEnd,
    offsetTop: scrollTop - rowStart * rowHeight,
    offsetLeft: scrollLeft - columnStart * columnWidth,
  };
}

function positiveSize(value: number | undefined, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value) && value > 0
    ? value
    : fallback;
}
