import type { Cell } from '@fortune-sheet/core';
import * as Y from 'yjs';
import type { WorkSpreadsheetSheet } from '../features/work/work-types';
import {
  OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
  OFFICE_KERNEL_SPREADSHEET_MAX_ROWS,
} from '../kernel/office-kernel-spreadsheet-protocol';
import { WorkOfficeCollaborationError } from './office-collaboration';
import { workOfficeCollaborationJsonEqual as jsonEqual } from './office-collaboration-json';
import {
  patchSpreadsheetFlatJsonMap,
  readSpreadsheetFlatJsonMap,
} from './office-spreadsheet-collaboration-flat-json';
import { invalidWorkOfficeSpreadsheetShared as invalidSharedSpreadsheet } from './office-spreadsheet-collaboration-validation';

export const SPREADSHEET_RECORD_CELLS = 'cells';
export const SPREADSHEET_RECORD_CELL_PRESENCE = 'cellPresence';
export const SPREADSHEET_RECORD_CELL_MODE = 'cellMode';
export const SPREADSHEET_RECORD_DATA_ROW_LENGTHS = 'dataRowLengths';

type SpreadsheetCellMode = 'data' | 'celldata';

interface SpreadsheetCellEntry {
  cell: Cell;
  column: number;
  coordinate: string;
  row: number;
}

/**
 * Cells use one field map rather than assigning a nested map per coordinate.
 * Two clients can therefore populate different fields of the same blank cell
 * without Y.Map's concurrent child-type assignment discarding one child map.
 */
export function patchSpreadsheetCells(
  record: Y.Map<unknown>,
  previous: WorkSpreadsheetSheet | undefined,
  next: WorkSpreadsheetSheet,
): void {
  const fields = nestedSpreadsheetMap(
    record,
    SPREADSHEET_RECORD_CELLS,
    'sheet cell fields',
  );
  const presence = nestedSpreadsheetMap(
    record,
    SPREADSHEET_RECORD_CELL_PRESENCE,
    'sheet cell presence',
  );
  const rowLengths = nestedSpreadsheetArray(
    record,
    SPREADSHEET_RECORD_DATA_ROW_LENGTHS,
    'sheet data row lengths',
  );
  const before = spreadsheetCellEntries(previous);
  const after = spreadsheetCellEntries(next);
  const beforeByCoordinate = new Map(
    before.map((entry) => [entry.coordinate, entry]),
  );
  const afterByCoordinate = new Map(
    after.map((entry) => [entry.coordinate, entry]),
  );

  for (const entry of before) {
    if (afterByCoordinate.has(entry.coordinate)) continue;
    if (!presence.has(entry.coordinate)) continue;
    const current = readCell(fields, entry.row, entry.column);
    if (!jsonEqual(entry.cell, current)) {
      throw staleSpreadsheetConflict(
        `Cell '${entry.coordinate}' in sheet '${next.id}' changed before it was deleted`,
      );
    }
    presence.delete(entry.coordinate);
    deleteCellFields(fields, entry.row, entry.column);
  }

  for (const entry of after) {
    const beforeEntry = beforeByCoordinate.get(entry.coordinate);
    if (beforeEntry && jsonEqual(beforeEntry.cell, entry.cell)) continue;
    if (beforeEntry && !presence.has(entry.coordinate)) {
      throw staleSpreadsheetConflict(
        `Cell '${entry.coordinate}' in sheet '${next.id}' was removed before this change could be applied`,
      );
    }
    presence.set(entry.coordinate, true);
    patchCellFields(
      fields,
      beforeEntry?.cell,
      entry.cell,
      entry.row,
      entry.column,
      next.id as string,
    );
  }

  patchCellProjectionMetadata(record, rowLengths, previous, next);
}

export function readSpreadsheetCells(
  record: Y.Map<unknown>,
  sheet: WorkSpreadsheetSheet,
): void {
  const fields = requiredSpreadsheetMap(
    record,
    SPREADSHEET_RECORD_CELLS,
    'sheet cell fields',
  );
  const presence = requiredSpreadsheetMap(
    record,
    SPREADSHEET_RECORD_CELL_PRESENCE,
    'sheet cell presence',
  );
  const rowLengths = requiredSpreadsheetArray(
    record,
    SPREADSHEET_RECORD_DATA_ROW_LENGTHS,
    'sheet data row lengths',
  );
  const grouped = new Map<string, SpreadsheetCellEntry>();
  for (const [coordinate, value] of presence.entries()) {
    if (value !== true) invalidSharedSpreadsheet('sheet cell presence');
    const decoded = decodedCoordinate(coordinate);
    grouped.set(coordinate, {
      ...decoded,
      cell: {},
      coordinate,
    });
  }
  for (const [encoded, value] of fields.entries()) {
    const { column, row } = decodedCellField(encoded);
    const coordinate = encodedCoordinate(row, column);
    const entry = grouped.get(coordinate);
    if (!entry || value instanceof Y.AbstractType) {
      invalidSharedSpreadsheet(`cell field '${encoded}'`);
    }
  }
  for (const entry of grouped.values()) {
    entry.cell = readCell(fields, entry.row, entry.column);
  }
  const entries = Array.from(grouped.values()).sort(
    (left, right) => left.row - right.row || left.column - right.column,
  );
  const mode = optionalCellMode(record.get(SPREADSHEET_RECORD_CELL_MODE));
  if (mode === undefined) {
    if (entries.length > 0 || rowLengths.length > 0) {
      invalidSharedSpreadsheet('cell projection mode');
    }
    return;
  }
  if (mode === 'celldata') {
    if (rowLengths.length > 0) {
      invalidSharedSpreadsheet('sparse cell projection dimensions');
    }
    sheet.celldata = entries.map(({ row, column, cell }) => ({
      r: row,
      c: column,
      v: cell,
    }));
    return;
  }

  const lengths = rowLengths.toArray().map((value, row) => {
    if (!Number.isSafeInteger(value) || value < 0) {
      invalidSharedSpreadsheet(`data row length at row ${row}`);
    }
    return value;
  });
  const data: Array<Array<Cell | null>> = lengths.map((length) =>
    Array.from({ length }, () => null),
  );
  for (const entry of entries) {
    if (entry.row >= data.length || entry.column >= (lengths[entry.row] ?? 0)) {
      invalidSharedSpreadsheet(
        `cell '${entry.coordinate}' outside the stored data matrix`,
      );
    }
    data[entry.row][entry.column] = entry.cell;
  }
  sheet.data = data;
}

function spreadsheetCellEntries(
  sheet: WorkSpreadsheetSheet | undefined,
): SpreadsheetCellEntry[] {
  if (!sheet) return [];
  const entries: SpreadsheetCellEntry[] = [];
  if (sheet.data !== undefined) {
    for (const [row, values] of sheet.data.entries()) {
      for (const [column, cell] of values.entries()) {
        if (cell !== null) {
          entries.push({
            cell,
            column,
            coordinate: encodedCoordinate(row, column),
            row,
          });
        }
      }
    }
    return entries;
  }
  for (const entry of sheet.celldata ?? []) {
    if (entry.v !== null) {
      entries.push({
        cell: entry.v,
        column: entry.c,
        coordinate: encodedCoordinate(entry.r, entry.c),
        row: entry.r,
      });
    }
  }
  return entries;
}

function patchCellFields(
  target: Y.Map<unknown>,
  previous: Cell | undefined,
  next: Cell,
  row: number,
  column: number,
  sheetId: string,
): void {
  patchSpreadsheetFlatJsonMap(
    new SpreadsheetCellFieldView(target, row, column),
    previous as Record<string, unknown> | undefined,
    next as Record<string, unknown>,
    `cell '${row}:${column}' in sheet '${sheetId}'`,
  );
}

function readCell(target: Y.Map<unknown>, row: number, column: number): Cell {
  return readSpreadsheetFlatJsonMap(
    new SpreadsheetCellFieldView(target, row, column),
    `cell '${row}:${column}'`,
  ) as Cell;
}

function patchCellProjectionMetadata(
  record: Y.Map<unknown>,
  rowLengths: Y.Array<number>,
  previous: WorkSpreadsheetSheet | undefined,
  next: WorkSpreadsheetSheet,
): void {
  const previousMode = cellMode(previous);
  const nextMode = cellMode(next);
  if (previousMode !== nextMode) {
    if (nextMode === undefined) record.delete(SPREADSHEET_RECORD_CELL_MODE);
    else record.set(SPREADSHEET_RECORD_CELL_MODE, nextMode);
  }
  const before = dataRowLengths(previous);
  const after = dataRowLengths(next);
  if (!jsonEqual(before, after)) {
    const shared = rowLengths.toArray();
    const merged = mergeStaleRowLengths(before, after, shared);
    if (rowLengths.length > 0) rowLengths.delete(0, rowLengths.length);
    if (merged.length > 0) rowLengths.push(merged);
  }
}

function mergeStaleRowLengths(
  previous: number[],
  next: number[],
  shared: number[],
): number[] {
  const length = Math.max(next.length, shared.length);
  const result: number[] = [];
  for (let index = 0; index < length; index += 1) {
    const before = previous[index];
    const after = next[index];
    const current = shared[index];
    if (after === before) result[index] = current ?? after ?? 0;
    else if (current === before || current === undefined)
      result[index] = after ?? 0;
    else result[index] = Math.max(after ?? 0, current);
  }
  while (result.length > next.length && result.at(-1) === 0) result.pop();
  return result;
}

function cellMode(
  sheet: WorkSpreadsheetSheet | undefined,
): SpreadsheetCellMode | undefined {
  if (sheet?.data !== undefined) return 'data';
  if (sheet?.celldata !== undefined) return 'celldata';
  return undefined;
}

function dataRowLengths(sheet: WorkSpreadsheetSheet | undefined): number[] {
  return sheet?.data?.map((row) => row.length) ?? [];
}

function nestedSpreadsheetMap(
  record: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Map<unknown> {
  const existing = record.get(key);
  if (existing instanceof Y.Map) return existing;
  if (existing !== undefined) invalidSharedSpreadsheet(label);
  const value = new Y.Map<unknown>();
  record.set(key, value);
  return value;
}

function nestedSpreadsheetArray(
  record: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Array<number> {
  const existing = record.get(key);
  if (existing instanceof Y.Array) return existing as Y.Array<number>;
  if (existing !== undefined) invalidSharedSpreadsheet(label);
  const value = new Y.Array<number>();
  record.set(key, value);
  return value;
}

function requiredSpreadsheetMap(
  record: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Map<unknown> {
  const value = record.get(key);
  if (!(value instanceof Y.Map)) invalidSharedSpreadsheet(label);
  return value as Y.Map<unknown>;
}

function requiredSpreadsheetArray(
  record: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Array<number> {
  const value = record.get(key);
  if (!(value instanceof Y.Array)) invalidSharedSpreadsheet(label);
  return value as Y.Array<number>;
}

function optionalCellMode(value: unknown): SpreadsheetCellMode | undefined {
  if (value === undefined) return undefined;
  if (value === 'data' || value === 'celldata') return value;
  invalidSharedSpreadsheet('cell projection mode');
}

function encodedCoordinate(row: number, column: number): string {
  return `${row}:${column}`;
}

function decodedCoordinate(coordinate: string): {
  column: number;
  row: number;
} {
  const match = /^(0|[1-9]\d*):(0|[1-9]\d*)$/.exec(coordinate);
  if (!match) invalidSharedSpreadsheet(`cell coordinate '${coordinate}'`);
  const row = Number(match[1]);
  const column = Number(match[2]);
  if (
    !Number.isSafeInteger(row) ||
    !Number.isSafeInteger(column) ||
    row >= OFFICE_KERNEL_SPREADSHEET_MAX_ROWS ||
    column >= OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS
  ) {
    invalidSharedSpreadsheet(`cell coordinate '${coordinate}'`);
  }
  return { row, column };
}

function encodedCellField(row: number, column: number, key: string): string {
  return JSON.stringify([row, column, key]);
}

class SpreadsheetCellFieldView {
  readonly #source: Y.Map<unknown>;
  readonly #row: number;
  readonly #column: number;

  constructor(source: Y.Map<unknown>, row: number, column: number) {
    this.#source = source;
    this.#row = row;
    this.#column = column;
  }

  get(key: string): unknown {
    return this.#source.get(encodedCellField(this.#row, this.#column, key));
  }

  set(key: string, value: unknown): unknown {
    this.#source.set(encodedCellField(this.#row, this.#column, key), value);
    return value;
  }

  delete(key: string): unknown {
    return this.#source.delete(encodedCellField(this.#row, this.#column, key));
  }

  *entries(): IterableIterator<[string, unknown]> {
    for (const [encoded, value] of this.#source.entries()) {
      const identity = decodedCellField(encoded);
      if (identity.row === this.#row && identity.column === this.#column) {
        yield [identity.key, value];
      }
    }
  }
}

function deleteCellFields(
  source: Y.Map<unknown>,
  row: number,
  column: number,
): void {
  for (const encoded of Array.from(source.keys())) {
    const identity = decodedCellField(encoded);
    if (identity.row === row && identity.column === column) {
      source.delete(encoded);
    }
  }
}

function decodedCellField(encoded: string): {
  column: number;
  key: string;
  row: number;
} {
  let value: unknown;
  try {
    value = JSON.parse(encoded);
  } catch {
    invalidSharedSpreadsheet(`cell field identity '${encoded}'`);
  }
  if (
    !Array.isArray(value) ||
    value.length !== 3 ||
    typeof value[2] !== 'string' ||
    !value[2] ||
    encodedCellField(value[0] as number, value[1] as number, value[2]) !==
      encoded
  ) {
    invalidSharedSpreadsheet(`cell field identity '${encoded}'`);
  }
  const coordinate = decodedCoordinate(`${value[0]}:${value[1]}`);
  return { ...coordinate, key: value[2] };
}

function staleSpreadsheetConflict(
  message: string,
): WorkOfficeCollaborationError {
  return new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `${message}. Refresh the shared snapshot before retrying.`,
  );
}
