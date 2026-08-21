import type { Cell, Op } from '@fortune-sheet/core';
import { sparseArrayEntries, sparseArrayIndexes } from '../spreadsheet-sparse';
import type {
  WorkSpreadsheetSheet,
  WorkSpreadsheetTable,
  WorkSpreadsheetTableColumn,
  WorkSpreadsheetTableFilter,
} from '../work-types';

export type SpreadsheetTableStructureChange =
  | {
      axis: 'column' | 'row';
      count: number;
      direction: 'lefttop' | 'rightbottom';
      index: number;
      kind: 'insert';
    }
  | {
      axis: 'column' | 'row';
      end: number;
      kind: 'delete';
      start: number;
    };

interface HeaderCellUpdate {
  column: number;
  name: string;
  row: number;
}

export function canApplySpreadsheetTableStructureChange(
  sheet: WorkSpreadsheetSheet | undefined,
  change: SpreadsheetTableStructureChange,
): boolean {
  if (!sheet || !validSpreadsheetTableStructureChange(change)) return false;
  for (const table of sheet.tables ?? []) {
    if (
      change.kind === 'delete' &&
      change.axis === 'row' &&
      spreadsheetTableDeleteRemovesSemanticRow(table, change.start, change.end)
    ) {
      return false;
    }
    if (!transformSpreadsheetTableStructure(table, change)) return false;
  }
  return true;
}

/**
 * Reconciles Fortune's cell and whole-row/column operation stream with the
 * browser-owned ListObject model. Fortune preserves unknown sheet metadata,
 * but it cannot update table ranges, column identities, or filter offsets.
 */
export function reconcileSpreadsheetTablesAfterFortune(
  sheets: WorkSpreadsheetSheet[],
  sourceSheets: WorkSpreadsheetSheet[],
  operations: readonly Op[] = [],
): WorkSpreadsheetSheet[] {
  return sheets.map((sheet, index) => {
    const source = sourceSheetForReconciliation(sheet, sourceSheets, index);
    if (!source?.tables?.length) return sheet;
    const sheetOperations = operations.filter(
      (operation) => operation.id === source.id,
    );
    let tables = source.tables;
    const reconcileColumns = new Set<string>();
    if (!operations.length) {
      for (const table of tables) reconcileColumns.add(table.id);
    }

    for (const operation of sheetOperations) {
      const structure = spreadsheetTableStructureChangeFromOperation(operation);
      if (structure) {
        const transformed: WorkSpreadsheetTable[] = [];
        for (const table of tables) {
          const next = transformSpreadsheetTableStructure(table, structure);
          if (!next) continue;
          transformed.push(next);
          if (next !== table) reconcileColumns.add(table.id);
        }
        tables = transformed;
        continue;
      }
      const coordinate = spreadsheetCellCoordinateFromOperation(operation);
      if (!coordinate) continue;
      const table = tables.find(
        (candidate) =>
          candidate.headerRow &&
          candidate.range.row[0] === coordinate.row &&
          coordinate.column >= candidate.range.column[0] &&
          coordinate.column <= candidate.range.column[1],
      );
      if (table) reconcileColumns.add(table.id);
    }

    const updates: HeaderCellUpdate[] = [];
    const readCell = spreadsheetSheetCellReader(sheet);
    tables = tables.map((table) => {
      if (!reconcileColumns.has(table.id)) return table;
      const columns = canonicalSpreadsheetTableColumns(table, readCell);
      if (table.headerRow) {
        for (const [offset, column] of columns.entries()) {
          const cellColumn = table.range.column[0] + offset;
          if (
            spreadsheetCellText(readCell(table.range.row[0], cellColumn)) !==
            column.name
          ) {
            updates.push({
              column: cellColumn,
              name: column.name,
              row: table.range.row[0],
            });
          }
        }
      }
      return sameSpreadsheetTableColumns(table.columns, columns)
        ? table
        : { ...table, columns };
    });

    const withHeaders = stampSpreadsheetTableHeaderCells(sheet, updates);
    return {
      ...withHeaders,
      tables: tables.length ? tables : undefined,
    };
  });
}

function transformSpreadsheetTableStructure(
  table: WorkSpreadsheetTable,
  change: SpreadsheetTableStructureChange,
): WorkSpreadsheetTable | null {
  if (!validSpreadsheetTableStructureChange(change)) return null;
  const axis = change.axis === 'row' ? table.range.row : table.range.column;
  const transformed =
    change.kind === 'insert'
      ? transformSpreadsheetTableAxisForInsertion(axis, change)
      : transformSpreadsheetTableAxisForDeletion(axis, change);
  if (!transformed) return null;
  const range = {
    row:
      change.axis === 'row'
        ? transformed.axis
        : ([...table.range.row] as [number, number]),
    column:
      change.axis === 'column'
        ? transformed.axis
        : ([...table.range.column] as [number, number]),
  };
  const height = range.row[1] - range.row[0] + 1;
  if (height <= Number(table.headerRow) + Number(table.totalsRow)) return null;

  if (change.axis !== 'column' || !transformed.inside) {
    if (
      range.row[0] === table.range.row[0] &&
      range.row[1] === table.range.row[1] &&
      range.column[0] === table.range.column[0] &&
      range.column[1] === table.range.column[1]
    ) {
      return table;
    }
    return { ...table, range };
  }

  if (change.kind === 'insert') {
    const columns = [...table.columns];
    columns.splice(
      transformed.offset,
      0,
      ...Array.from({ length: change.count }, () => ({ name: '' })),
    );
    return {
      ...table,
      range,
      columns,
      filters: table.filters.map((filter) =>
        filter.column >= transformed.offset
          ? { ...filter, column: filter.column + change.count }
          : filter,
      ),
    };
  }

  if (!('removed' in transformed)) return null;
  const removed = transformed.removed;
  if (typeof removed !== 'number') return null;
  const columns = [...table.columns];
  columns.splice(transformed.offset, removed);
  if (!columns.length) return null;
  return {
    ...table,
    range,
    columns,
    filters: filtersAfterSpreadsheetTableColumnDeletion(
      table.filters,
      transformed.offset,
      removed,
    ),
  };
}

function transformSpreadsheetTableAxisForInsertion(
  axis: readonly number[],
  change: Extract<SpreadsheetTableStructureChange, { kind: 'insert' }>,
): { axis: [number, number]; inside: boolean; offset: number } {
  const start = axis[0] ?? 0;
  const end = axis[1] ?? start;
  const insertion = change.index + (change.direction === 'rightbottom' ? 1 : 0);
  if (insertion <= start) {
    return {
      axis: [start + change.count, end + change.count],
      inside: false,
      offset: 0,
    };
  }
  if (insertion <= end) {
    return {
      axis: [start, end + change.count],
      inside: true,
      offset: insertion - start,
    };
  }
  return { axis: [start, end], inside: false, offset: 0 };
}

function transformSpreadsheetTableAxisForDeletion(
  axis: readonly number[],
  change: Extract<SpreadsheetTableStructureChange, { kind: 'delete' }>,
): {
  axis: [number, number];
  inside: boolean;
  offset: number;
  removed: number;
} | null {
  const start = axis[0] ?? 0;
  const end = axis[1] ?? start;
  const count = change.end - change.start + 1;
  if (change.end < start) {
    return {
      axis: [start - count, end - count],
      inside: false,
      offset: 0,
      removed: 0,
    };
  }
  if (change.start > end) {
    return { axis: [start, end], inside: false, offset: 0, removed: 0 };
  }
  const overlapStart = Math.max(start, change.start);
  const overlapEnd = Math.min(end, change.end);
  const removed = overlapEnd - overlapStart + 1;
  const remaining = end - start + 1 - removed;
  if (remaining <= 0) return null;
  const deletedBeforeStart = Math.max(
    0,
    Math.min(change.end, start - 1) - change.start + 1,
  );
  const nextStart = start - deletedBeforeStart;
  return {
    axis: [nextStart, nextStart + remaining - 1],
    inside: true,
    offset: overlapStart - start,
    removed,
  };
}

function filtersAfterSpreadsheetTableColumnDeletion(
  filters: readonly WorkSpreadsheetTableFilter[],
  offset: number,
  removed: number,
): WorkSpreadsheetTableFilter[] {
  const end = offset + removed - 1;
  return filters.flatMap((filter) => {
    if (filter.column >= offset && filter.column <= end) return [];
    return [
      filter.column > end
        ? { ...filter, column: filter.column - removed }
        : filter,
    ];
  });
}

function canonicalSpreadsheetTableColumns(
  table: WorkSpreadsheetTable,
  readCell: (row: number, column: number) => Cell | null,
): WorkSpreadsheetTableColumn[] {
  const width = table.range.column[1] - table.range.column[0] + 1;
  const observed = new Set<string>();
  return Array.from({ length: width }, (_, offset) => {
    const raw = table.headerRow
      ? spreadsheetCellText(
          readCell(table.range.row[0], table.range.column[0] + offset),
        ).trim()
      : (table.columns[offset]?.name.trim() ?? '');
    const base = validSpreadsheetTableColumnName(raw)
      ? raw
      : `Column${offset + 1}`;
    let candidate = base;
    let suffix = 2;
    while (observed.has(candidate.toLocaleLowerCase())) {
      const suffixText = String(suffix);
      candidate = `${truncateSpreadsheetTableColumnName(
        base,
        255 - suffixText.length,
      )}${suffixText}`;
      suffix += 1;
    }
    observed.add(candidate.toLocaleLowerCase());
    return { name: candidate };
  });
}

function spreadsheetTableDeleteRemovesSemanticRow(
  table: WorkSpreadsheetTable,
  start: number,
  end: number,
): boolean {
  return Boolean(
    (table.headerRow &&
      start <= table.range.row[0] &&
      end >= table.range.row[0]) ||
      (table.totalsRow &&
        start <= table.range.row[1] &&
        end >= table.range.row[1]),
  );
}

function spreadsheetTableStructureChangeFromOperation(
  operation: Op,
): SpreadsheetTableStructureChange | null {
  if (!isRecord(operation.value)) return null;
  if (operation.op === 'insertRowCol') {
    const { count, direction, index, type } = operation.value;
    const change: SpreadsheetTableStructureChange = {
      axis: type === 'column' ? 'column' : 'row',
      count: Number(count),
      direction: direction === 'rightbottom' ? 'rightbottom' : 'lefttop',
      index: Number(index),
      kind: 'insert',
    };
    return (type === 'column' || type === 'row') &&
      (direction === 'lefttop' || direction === 'rightbottom') &&
      validSpreadsheetTableStructureChange(change)
      ? change
      : null;
  }
  if (operation.op === 'deleteRowCol') {
    const { end, start, type } = operation.value;
    const change: SpreadsheetTableStructureChange = {
      axis: type === 'column' ? 'column' : 'row',
      end: Number(end),
      kind: 'delete',
      start: Number(start),
    };
    return (type === 'column' || type === 'row') &&
      validSpreadsheetTableStructureChange(change)
      ? change
      : null;
  }
  return null;
}

function spreadsheetCellCoordinateFromOperation(
  operation: Op,
): { column: number; row: number } | null {
  if (operation.path[0] !== 'data' || operation.path.length < 3) return null;
  const row = operation.path[1];
  const column = operation.path[2];
  return Number.isSafeInteger(row) &&
    Number(row) >= 0 &&
    Number.isSafeInteger(column) &&
    Number(column) >= 0
    ? { column: Number(column), row: Number(row) }
    : null;
}

function validSpreadsheetTableStructureChange(
  change: SpreadsheetTableStructureChange,
): boolean {
  if (change.kind === 'insert') {
    return (
      Number.isSafeInteger(change.index) &&
      change.index >= 0 &&
      Number.isSafeInteger(change.count) &&
      change.count > 0
    );
  }
  return (
    Number.isSafeInteger(change.start) &&
    change.start >= 0 &&
    Number.isSafeInteger(change.end) &&
    change.end >= change.start
  );
}

function sourceSheetForReconciliation(
  sheet: WorkSpreadsheetSheet,
  sourceSheets: WorkSpreadsheetSheet[],
  index: number,
): WorkSpreadsheetSheet | undefined {
  return (
    (sheet.id
      ? sourceSheets.find((candidate) => candidate.id === sheet.id)
      : undefined) ?? sourceSheets[index]
  );
}

function spreadsheetSheetCellReader(
  sheet: WorkSpreadsheetSheet,
): (row: number, column: number) => Cell | null {
  if (sheet.data !== undefined) {
    return (row, column) => sheet.data?.[row]?.[column] ?? null;
  }
  const cells = new Map(
    (sheet.celldata ?? []).map((entry) => [`${entry.r}:${entry.c}`, entry.v]),
  );
  return (row, column) => cells.get(`${row}:${column}`) ?? null;
}

function stampSpreadsheetTableHeaderCells(
  sheet: WorkSpreadsheetSheet,
  updates: readonly HeaderCellUpdate[],
): WorkSpreadsheetSheet {
  if (!updates.length) return sheet;
  if (sheet.data !== undefined) {
    const data: NonNullable<WorkSpreadsheetSheet['data']> = [];
    data.length = sheet.data.length;
    for (const [row, value] of sparseArrayEntries(sheet.data))
      data[row] = value;
    const mutableRows = new Map<number, NonNullable<(typeof data)[number]>>();
    for (const update of updates) {
      let row = mutableRows.get(update.row);
      if (!row) {
        const source = data[update.row];
        row = [];
        row.length = source?.length ?? 0;
        for (const column of sparseArrayIndexes(source))
          row[column] = source?.[column];
        mutableRows.set(update.row, row);
        data[update.row] = row;
      }
      row[update.column] = stampSpreadsheetTableHeaderCell(
        row[update.column],
        update.name,
      );
    }
    return { ...sheet, data };
  }

  const celldata = [...(sheet.celldata ?? [])];
  const indexes = new Map(
    celldata.map((entry, index) => [`${entry.r}:${entry.c}`, index]),
  );
  for (const update of updates) {
    const key = `${update.row}:${update.column}`;
    const index = indexes.get(key);
    const source = index === undefined ? undefined : celldata[index];
    const entry = {
      r: update.row,
      c: update.column,
      v: stampSpreadsheetTableHeaderCell(source?.v, update.name),
    };
    if (index === undefined) {
      indexes.set(key, celldata.length);
      celldata.push(entry);
    } else {
      celldata[index] = entry;
    }
  }
  return { ...sheet, celldata };
}

function stampSpreadsheetTableHeaderCell(
  cell: Cell | null | undefined,
  name: string,
): Cell {
  return { ...(cell ?? {}), m: name, v: name };
}

function spreadsheetCellText(cell: Cell | null | undefined): string {
  const value = cell?.m ?? cell?.v;
  return value === undefined || value === null ? '' : String(value);
}

function validSpreadsheetTableColumnName(name: string): boolean {
  const characters = Array.from(name);
  return (
    name.trim() === name &&
    characters.length >= 1 &&
    characters.length <= 255 &&
    !characters.some(
      (character) =>
        /\p{Cc}/u.test(character) ||
        character === '\uFFFE' ||
        character === '\uFFFF',
    )
  );
}

function truncateSpreadsheetTableColumnName(
  value: string,
  maximum: number,
): string {
  return Array.from(value).slice(0, maximum).join('');
}

function sameSpreadsheetTableColumns(
  left: readonly WorkSpreadsheetTableColumn[],
  right: readonly WorkSpreadsheetTableColumn[],
): boolean {
  return (
    left.length === right.length &&
    left.every((column, index) => column.name === right[index]?.name)
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
