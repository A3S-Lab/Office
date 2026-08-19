import type { Cell, CellMatrix } from '@fortune-sheet/core';
import { sameSpreadsheetHistoryValue } from './spreadsheet-history-value';
import { sparseArrayIndexes } from './spreadsheet-sparse';

export const SPREADSHEET_SHOWN_COMMENT_CELLS_PROPERTY =
  '__a3sShownCommentCells';

export interface SpreadsheetMatrixProfile {
  columnCount: number;
  formulaCells: readonly Readonly<{ column: number; row: number }>[];
  fortuneReady: boolean;
  historyRoot: object;
  historyState: object;
  populatedCellCount: number;
  protectionCellKey: string;
  rowCount: number;
  shownCommentCells: readonly Readonly<{ c: number; r: number }>[];
}

export interface SpreadsheetMatrixCellChange {
  column: number;
  current: Cell | null | undefined;
  previous: Cell | null | undefined;
  row: number;
}

interface ImportedSpreadsheetMatrixOptions {
  columnCount: number;
  formulaCells: readonly Readonly<{ column: number; row: number }>[];
  fortuneReady: boolean;
  populatedCellCount: number;
  protectionCellKey: string;
  rowCount: number;
  shownCommentCells: readonly Readonly<{ c: number; r: number }>[];
}

const spreadsheetMatrixProfiles = new WeakMap<
  CellMatrix,
  SpreadsheetMatrixProfile
>();

/**
 * Freezes a package-owned import cell before it becomes controlled content.
 * Fortune can then share the cell as an Immer base value without recursively
 * freezing or cloning every populated cell during editor initialization.
 */
export function freezeImportedSpreadsheetCell(cell: Cell): Cell {
  freezeObject(cell.mc);
  freezeObject(cell.ct?.s);
  freezeObject(cell.ct);
  freezeObject(cell.ps);
  freezeObject(cell.hl);
  freezeObject(cell.spl);
  return Object.freeze(cell);
}

/**
 * Authenticates a normalized, package-owned matrix for clone-free projection.
 * The registry is identity-based, so an edited or host-created matrix cannot
 * accidentally inherit the optimization.
 */
export function registerImportedSpreadsheetMatrix(
  data: CellMatrix,
  options: ImportedSpreadsheetMatrixOptions,
): void {
  data.length = Math.max(data.length, options.rowCount);
  data[0] ??= [];
  for (const rowIndex of sparseArrayIndexes(data)) {
    const row = data[rowIndex];
    if (!row) continue;
    row.length = Math.max(row.length, options.columnCount);
    Object.freeze(row);
  }
  registerSpreadsheetMatrixProfile(data, options);
}

/**
 * Carries import-time summaries across a bounded cell edit without rescanning
 * the complete matrix. The source identity must have been authenticated by
 * this package and must already be directly consumable by Fortune.
 */
export function registerDerivedSpreadsheetMatrix(
  data: CellMatrix,
  source: CellMatrix,
  changes: readonly SpreadsheetMatrixCellChange[],
): boolean {
  const sourceProfile = spreadsheetMatrixProfiles.get(source);
  if (!sourceProfile?.fortuneReady) return false;

  const formulaCells = new Map(
    sourceProfile.formulaCells.map(({ column, row }) => [
      spreadsheetCoordinateKey(row, column),
      { column, row },
    ]),
  );
  const shownCommentCells = new Map(
    sourceProfile.shownCommentCells.map(({ c, r }) => [
      spreadsheetCoordinateKey(r, c),
      { c, r },
    ]),
  );
  let populatedCellCount = sourceProfile.populatedCellCount;
  let protectionCellKey = sourceProfile.protectionCellKey;
  let protectionCells: Map<string, string> | undefined;

  for (const { column, current, previous, row } of changes) {
    const coordinateKey = spreadsheetCoordinateKey(row, column);
    if (previous == null && current != null) populatedCellCount += 1;
    if (previous != null && current == null) populatedCellCount -= 1;

    if (current?.f) formulaCells.set(coordinateKey, { column, row });
    else formulaCells.delete(coordinateKey);

    if (current?.ps?.isShow) {
      shownCommentCells.set(coordinateKey, { c: column, r: row });
    } else {
      shownCommentCells.delete(coordinateKey);
    }

    if (
      spreadsheetProtectionSignature(previous) !==
      spreadsheetProtectionSignature(current)
    ) {
      protectionCells ??= spreadsheetProtectionCells(protectionCellKey);
      const signature = spreadsheetProtectionSignature(current);
      if (signature === null) protectionCells.delete(coordinateKey);
      else protectionCells.set(coordinateKey, signature);
    }

    const changedRow = data[row];
    if (changedRow && !Object.isFrozen(changedRow)) Object.freeze(changedRow);
  }

  if (protectionCells) {
    protectionCellKey = Array.from(protectionCells, ([coordinate, value]) => ({
      coordinate,
      value,
    }))
      .sort((left, right) =>
        compareSpreadsheetCoordinates(left.coordinate, right.coordinate),
      )
      .map(({ coordinate, value }) => `${coordinate}:${value}`)
      .join(',');
  }

  registerSpreadsheetMatrixProfile(
    data,
    {
      columnCount: Math.max(
        sourceProfile.columnCount,
        ...changes.map(({ column, row }) =>
          Math.max(column + 1, data[row]?.length ?? 0),
        ),
      ),
      formulaCells: Array.from(formulaCells.values()).sort(
        compareSpreadsheetCells,
      ),
      fortuneReady: true,
      populatedCellCount: Math.max(0, populatedCellCount),
      protectionCellKey,
      rowCount: Math.max(
        sourceProfile.rowCount,
        data.length,
        ...changes.map(({ row }) => row + 1),
      ),
      shownCommentCells: Array.from(shownCommentCells.values()).sort(
        (left, right) => left.r - right.r || left.c - right.c,
      ),
    },
    {
      historyRoot: sourceProfile.historyRoot,
      historyState: changes.some(
        ({ current, previous }) =>
          !sameSpreadsheetHistoryValue(current, previous),
      )
        ? Object.freeze({})
        : sourceProfile.historyState,
    },
  );
  return true;
}

export function spreadsheetMatrixProfile(
  data: CellMatrix | undefined,
): SpreadsheetMatrixProfile | undefined {
  return data ? spreadsheetMatrixProfiles.get(data) : undefined;
}

export function attachSpreadsheetShownCommentCells(
  data: CellMatrix,
  cells: readonly Readonly<{ c: number; r: number }>[],
): void {
  Object.defineProperty(data, SPREADSHEET_SHOWN_COMMENT_CELLS_PROPERTY, {
    configurable: false,
    enumerable: false,
    value: cells,
    writable: false,
  });
}

function registerSpreadsheetMatrixProfile(
  data: CellMatrix,
  options: Pick<
    ImportedSpreadsheetMatrixOptions,
    | 'columnCount'
    | 'formulaCells'
    | 'fortuneReady'
    | 'populatedCellCount'
    | 'protectionCellKey'
    | 'rowCount'
    | 'shownCommentCells'
  >,
  history?: Pick<SpreadsheetMatrixProfile, 'historyRoot' | 'historyState'>,
): void {
  const initialHistoryState = Object.freeze({});
  const profile = Object.freeze({
    columnCount: options.columnCount,
    formulaCells: Object.freeze(
      options.formulaCells.map((cell) => Object.freeze({ ...cell })),
    ),
    fortuneReady: options.fortuneReady,
    historyRoot: history?.historyRoot ?? initialHistoryState,
    historyState: history?.historyState ?? initialHistoryState,
    populatedCellCount: options.populatedCellCount,
    protectionCellKey: options.protectionCellKey,
    rowCount: options.rowCount,
    shownCommentCells: Object.freeze(
      options.shownCommentCells.map((cell) => Object.freeze({ ...cell })),
    ),
  });
  attachSpreadsheetShownCommentCells(data, profile.shownCommentCells);
  Object.freeze(data);
  spreadsheetMatrixProfiles.set(data, profile);
}

function spreadsheetProtectionCells(value: string): Map<string, string> {
  const cells = new Map<string, string>();
  for (const entry of value.split(',')) {
    if (!entry) continue;
    const separator = entry.indexOf(':');
    if (separator <= 0) continue;
    cells.set(entry.slice(0, separator), entry.slice(separator + 1));
  }
  return cells;
}

function spreadsheetProtectionSignature(
  cell: Cell | null | undefined,
): string | null {
  const hidden = (cell as (Cell & { hi?: number }) | null | undefined)?.hi;
  if (cell?.lo === undefined && hidden === undefined) return null;
  return `${cell?.lo ?? ''}:${hidden ?? ''}`;
}

function spreadsheetCoordinateKey(row: number, column: number): string {
  return `${row}_${column}`;
}

function compareSpreadsheetCoordinates(left: string, right: string): number {
  const [leftRow = 0, leftColumn = 0] = left.split('_').map(Number);
  const [rightRow = 0, rightColumn = 0] = right.split('_').map(Number);
  return leftRow - rightRow || leftColumn - rightColumn;
}

function compareSpreadsheetCells(
  left: Readonly<{ column: number; row: number }>,
  right: Readonly<{ column: number; row: number }>,
): number {
  return left.row - right.row || left.column - right.column;
}

function freezeObject(value: unknown): void {
  if (value && typeof value === 'object') Object.freeze(value);
}
