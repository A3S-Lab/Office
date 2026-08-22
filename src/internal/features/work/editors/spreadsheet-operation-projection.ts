import type { Cell, CellMatrix, Op } from '@fortune-sheet/core';
import {
  freezeImportedSpreadsheetCell,
  registerDerivedSpreadsheetMatrix,
  type SpreadsheetMatrixCellChange,
  spreadsheetMatrixProfile,
} from '../work-spreadsheet-matrix-profile';
import type { WorkSpreadsheetContent } from '../work-types';
import {
  reconcileSpreadsheetRichTextCellEdit,
  type SpreadsheetRichTextPasteIntent,
} from '../work-xlsx-rich-text-edit';
import {
  consumeSpreadsheetRichTextPaste,
  peekSpreadsheetRichTextPaste,
} from './spreadsheet-rich-text-paste';

export const MAXIMUM_INCREMENTAL_SPREADSHEET_OPERATIONS = 10_000;

export interface SpreadsheetOperationProjection {
  affectedCellCount: number;
  sheets: WorkSpreadsheetContent['sheets'];
}

export interface SpreadsheetCellOperationCoordinate {
  column: number;
  row: number;
}

interface PendingRichTextPasteCoordinate {
  column: number;
  row: number;
  sheet: WorkSpreadsheetContent['sheets'][number];
}

export function spreadsheetCellOperationKey(
  row: number,
  column: number,
): string {
  return `${row}:${column}`;
}

export function spreadsheetCellOperationCoordinates(
  operations: readonly Op[],
): Map<string, Map<string, SpreadsheetCellOperationCoordinate>> | null {
  if (operations.length > MAXIMUM_INCREMENTAL_SPREADSHEET_OPERATIONS) {
    return null;
  }
  const coordinatesBySheet = new Map<
    string,
    Map<string, SpreadsheetCellOperationCoordinate>
  >();
  for (const operation of operations) {
    if (
      operation.op === 'insertRowCol' ||
      operation.op === 'deleteRowCol' ||
      operation.op === 'addSheet' ||
      operation.op === 'deleteSheet' ||
      operation.path.length === 0 ||
      !operation.id
    ) {
      return null;
    }
    const root = operation.path[0];
    if (root === 'id' || root === 'celldata') return null;
    if (root !== 'data') continue;
    const row = operation.path[1];
    const column = operation.path[2];
    if (
      operation.path.length < 3 ||
      !isSpreadsheetCoordinate(row) ||
      !isSpreadsheetCoordinate(column)
    ) {
      return null;
    }
    const coordinates = coordinatesBySheet.get(operation.id) ?? new Map();
    coordinatesBySheet.set(operation.id, coordinates);
    coordinates.set(spreadsheetCellOperationKey(row, column), { column, row });
  }
  return coordinatesBySheet;
}

/**
 * Reconciles a bounded Fortune operation batch against authenticated
 * controlled matrices. Only the outer matrix and affected rows/cells are
 * copied; structural or incomplete batches return null for the full-scan
 * compatibility path.
 */
export function projectSpreadsheetSheetsFromFortuneOperations(
  sheets: WorkSpreadsheetContent['sheets'],
  sourceSheets: WorkSpreadsheetContent['sheets'],
  operations: readonly Op[],
): SpreadsheetOperationProjection | null {
  if (operations.length === 0 || sheets.length !== sourceSheets.length) {
    return null;
  }

  const changedSheets = spreadsheetSheetsById(sheets);
  const controlledSheets = spreadsheetSheetsById(sourceSheets);
  if (!changedSheets || !controlledSheets) return null;
  if (
    changedSheets.size !== sheets.length ||
    controlledSheets.size !== sourceSheets.length
  ) {
    return null;
  }

  const coordinatesBySheet = spreadsheetCellOperationCoordinates(operations);
  if (!coordinatesBySheet) return null;
  for (const operation of operations) {
    if (
      !operation.id ||
      !changedSheets.has(operation.id) ||
      !controlledSheets.has(operation.id)
    ) {
      return null;
    }
  }

  let affectedCellCount = 0;
  const projectedSheets: WorkSpreadsheetContent['sheets'] = [];
  const pendingRichTextPastes: PendingRichTextPasteCoordinate[] = [];
  for (const sheet of sheets) {
    const id = sheet.id;
    if (!id) return null;
    const source = controlledSheets.get(id);
    if (!source) return null;
    const coordinates = coordinatesBySheet.get(id);
    const { celldata: _cellData, data: _data, ...metadata } = sheet;

    if (!coordinates?.size) {
      projectedSheets.push(
        source.data !== undefined
          ? { ...metadata, data: source.data }
          : { ...metadata, celldata: source.celldata ?? [] },
      );
      continue;
    }
    if (
      source.data === undefined ||
      sheet.data === undefined ||
      !spreadsheetMatrixProfile(source.data)?.fortuneReady
    ) {
      return null;
    }

    const projection = projectSpreadsheetMatrixCells(
      sheet,
      source,
      coordinates.values(),
      pendingRichTextPastes,
    );
    if (!projection) return null;
    affectedCellCount += projection.affectedCellCount;
    projectedSheets.push({ ...metadata, data: projection.data });
  }

  for (const pending of pendingRichTextPastes) {
    consumeSpreadsheetRichTextPaste(pending.sheet, pending.row, pending.column);
  }

  return { affectedCellCount, sheets: projectedSheets };
}

function projectSpreadsheetMatrixCells(
  changed: WorkSpreadsheetContent['sheets'][number],
  source: WorkSpreadsheetContent['sheets'][number],
  coordinates: Iterable<SpreadsheetCellOperationCoordinate>,
  pendingRichTextPastes: PendingRichTextPasteCoordinate[],
): { affectedCellCount: number; data: CellMatrix } | null {
  if (!changed.data || !source.data) return null;
  const data = source.data.slice();
  data.length = Math.max(
    data.length,
    changed.data.length,
    positiveSpreadsheetDimension(changed.row),
  );
  const mutableRows = new Map<number, CellMatrix[number]>();
  const profileChanges: SpreadsheetMatrixCellChange[] = [];

  for (const { column, row } of coordinates) {
    const changedRow = changed.data[row];
    const sourceRow = source.data[row];
    const currentCell = changedRow?.[column];
    const formattedPaste = currentCell
      ? peekSpreadsheetRichTextPaste(source, row, column, currentCell)
      : undefined;
    const nextCell =
      currentCell == null
        ? currentCell
        : cloneControlledSpreadsheetCell(
            currentCell,
            sourceRow?.[column],
            formattedPaste,
          );
    if (formattedPaste) {
      pendingRichTextPastes.push({ column, row, sheet: source });
    }
    let nextRow = mutableRows.get(row);
    if (!nextRow) {
      nextRow = sourceRow?.slice() ?? [];
      nextRow.length = Math.max(
        nextRow.length,
        changedRow?.length ?? 0,
        positiveSpreadsheetDimension(changed.column),
      );
      mutableRows.set(row, nextRow);
      data[row] = nextRow;
    }
    if (nextCell == null) delete nextRow[column];
    else nextRow[column] = nextCell;
    profileChanges.push({
      column,
      current: nextCell,
      previous: sourceRow?.[column],
      row,
    });
  }

  if (!registerDerivedSpreadsheetMatrix(data, source.data, profileChanges)) {
    return null;
  }
  return { affectedCellCount: profileChanges.length, data };
}

function cloneControlledSpreadsheetCell(
  cell: Cell,
  previous: Cell | null | undefined,
  formattedPaste?: SpreadsheetRichTextPasteIntent,
): Cell {
  return freezeImportedSpreadsheetCell(
    structuredClone(
      reconcileSpreadsheetRichTextCellEdit(previous, cell, formattedPaste),
    ),
  );
}

function spreadsheetSheetsById(
  sheets: WorkSpreadsheetContent['sheets'],
): Map<string, WorkSpreadsheetContent['sheets'][number]> | null {
  const byId = new Map<string, WorkSpreadsheetContent['sheets'][number]>();
  for (const sheet of sheets) {
    if (!sheet.id || byId.has(sheet.id)) return null;
    byId.set(sheet.id, sheet);
  }
  return byId;
}

function isSpreadsheetCoordinate(value: unknown): value is number {
  return Number.isSafeInteger(value) && Number(value) >= 0;
}

function positiveSpreadsheetDimension(value: unknown): number {
  return typeof value === 'number' && Number.isFinite(value) && value > 0
    ? Math.floor(value)
    : 0;
}
