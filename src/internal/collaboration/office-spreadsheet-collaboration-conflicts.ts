import type { Cell } from '@fortune-sheet/core';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../features/work/work-types';
import { sparseArrayEntries } from '../features/work/spreadsheet-sparse';
import { WorkOfficeCollaborationError } from './office-collaboration';
import {
  isWorkOfficeCollaborationRecord as isRecord,
  workOfficeCollaborationJsonEqual as jsonEqual,
} from './office-collaboration-json';

export function assertWorkOfficeSpreadsheetPatchSafe(
  previous: WorkSpreadsheetContent,
  next: WorkSpreadsheetContent,
  shared: WorkSpreadsheetContent,
): void {
  assertCompatibleValue(
    previous.dateSystem,
    next.dateSystem,
    shared.dateSystem,
    'date system',
  );
  assertCompatibleValue(
    previous.calculation,
    next.calculation,
    shared.calculation,
    'calculation settings',
  );
  assertRecordCollection(
    previous.sheets,
    next.sheets,
    shared.sheets,
    (sheet) => sheet.id as string,
    'sheet',
    assertCompatibleSheet,
  );
  assertRecordCollection(
    previous.namedRanges ?? [],
    next.namedRanges ?? [],
    shared.namedRanges ?? [],
    (range) => range.id,
    'named range',
  );
  assertSheetSidecars(previous, next, shared, 'printAreas', 'print area');
  assertSheetSidecars(previous, next, shared, 'printTitles', 'print titles');
  assertSheetSidecars(previous, next, shared, 'pageBreaks', 'page breaks');
  assertSheetSidecars(previous, next, shared, 'pageSetups', 'page setup');
}

function assertSheetSidecars<
  K extends 'printAreas' | 'printTitles' | 'pageBreaks' | 'pageSetups',
>(
  previous: WorkSpreadsheetContent,
  next: WorkSpreadsheetContent,
  shared: WorkSpreadsheetContent,
  key: K,
  label: string,
): void {
  type Sidecar = NonNullable<WorkSpreadsheetContent[K]>[number];
  assertRecordCollection(
    (previous[key] ?? []) as Sidecar[],
    (next[key] ?? []) as Sidecar[],
    (shared[key] ?? []) as Sidecar[],
    (record) => record.sheetId,
    label,
  );
}

function assertCompatibleSheet(
  previous: WorkSpreadsheetSheet,
  next: WorkSpreadsheetSheet,
  shared: WorkSpreadsheetSheet,
  label: string,
): void {
  const omitted = new Set([
    'data',
    'celldata',
    'config',
    'formulaMetadata',
    'images',
    'charts',
    'pivotTables',
    'tables',
  ]);
  assertCompatibleValue(
    withoutKeys(previous, omitted),
    withoutKeys(next, omitted),
    withoutKeys(shared, omitted),
    label,
  );
  assertCompatibleValue(
    previous.config,
    next.config,
    shared.config,
    `${label} config`,
  );
  assertCompatibleValue(
    previous.formulaMetadata,
    next.formulaMetadata,
    shared.formulaMetadata,
    `${label} formula metadata`,
  );
  assertRecordCollection(
    previous.images ?? [],
    next.images ?? [],
    shared.images ?? [],
    (image) => image.id,
    `${label} image`,
  );
  assertRecordCollection(
    previous.charts ?? [],
    next.charts ?? [],
    shared.charts ?? [],
    (chart) => chart.id,
    `${label} chart`,
  );
  assertRecordCollection(
    previous.pivotTables ?? [],
    next.pivotTables ?? [],
    shared.pivotTables ?? [],
    (pivot) => pivot.id,
    `${label} pivot table`,
  );
  assertRecordCollection(
    previous.tables ?? [],
    next.tables ?? [],
    shared.tables ?? [],
    (table) => table.id,
    `${label} table`,
  );
  assertCellCollection(
    spreadsheetCells(previous),
    spreadsheetCells(next),
    spreadsheetCells(shared),
    label,
  );
}

function assertCellCollection(
  previous: Map<string, Cell>,
  next: Map<string, Cell>,
  shared: Map<string, Cell>,
  sheetLabel: string,
): void {
  for (const [coordinate, before] of previous) {
    const after = next.get(coordinate);
    const current = shared.get(coordinate);
    const label = `cell '${coordinate}' in ${sheetLabel}`;
    if (after === undefined) {
      if (current !== undefined && !jsonEqual(before, current)) conflict(label);
      continue;
    }
    if (jsonEqual(before, after)) continue;
    if (current === undefined) removedConflict(label);
    assertCompatibleValue(before, after, current, label);
  }
  for (const [coordinate, after] of next) {
    if (previous.has(coordinate)) continue;
    const current = shared.get(coordinate);
    if (current !== undefined && !jsonEqual(after, current)) {
      conflict(`cell '${coordinate}' in ${sheetLabel}`);
    }
  }
}

function assertRecordCollection<T>(
  previous: readonly T[],
  next: readonly T[],
  shared: readonly T[],
  identity: (value: T) => string,
  label: string,
  compare: (
    previous: T,
    next: T,
    shared: T,
    label: string,
  ) => void = assertCompatibleValue,
): void {
  const nextById = new Map(next.map((value) => [identity(value), value]));
  const sharedById = new Map(shared.map((value) => [identity(value), value]));
  const previousIds = new Set<string>();
  for (const before of previous) {
    const id = identity(before);
    previousIds.add(id);
    const after = nextById.get(id);
    const current = sharedById.get(id);
    const recordLabel = `${label} '${id}'`;
    if (after === undefined) {
      if (current !== undefined && !jsonEqual(before, current)) {
        conflict(recordLabel);
      }
      continue;
    }
    if (jsonEqual(before, after)) continue;
    if (current === undefined) removedConflict(recordLabel);
    compare(before, after, current, recordLabel);
  }
  for (const after of next) {
    const id = identity(after);
    if (previousIds.has(id)) continue;
    const current = sharedById.get(id);
    if (current !== undefined && !jsonEqual(after, current)) {
      conflict(`${label} '${id}'`);
    }
  }
}

function assertCompatibleValue<T>(
  previous: T,
  next: T,
  shared: T,
  label: string,
): void {
  if (jsonEqual(previous, next)) return;
  if (isRecord(previous) && isRecord(next) && isRecord(shared)) {
    const keys = new Set([...Object.keys(previous), ...Object.keys(next)]);
    for (const key of keys) {
      const hadPrevious = Object.hasOwn(previous, key);
      const hasNext = Object.hasOwn(next, key);
      const hasShared = Object.hasOwn(shared, key);
      const fieldLabel = `${label} field '${key}'`;
      if (!hasNext) {
        if (
          hadPrevious &&
          hasShared &&
          !jsonEqual(previous[key], shared[key])
        ) {
          conflict(fieldLabel);
        }
        continue;
      }
      if (!hadPrevious) {
        if (hasShared && !jsonEqual(next[key], shared[key]))
          conflict(fieldLabel);
        continue;
      }
      if (jsonEqual(previous[key], next[key])) continue;
      if (!hasShared) removedConflict(fieldLabel);
      assertCompatibleValue(previous[key], next[key], shared[key], fieldLabel);
    }
    return;
  }
  if (!jsonEqual(previous, shared) && !jsonEqual(next, shared)) conflict(label);
}

function spreadsheetCells(sheet: WorkSpreadsheetSheet): Map<string, Cell> {
  const result = new Map<string, Cell>();
  if (sheet.data !== undefined) {
    for (const [row, values] of sparseArrayEntries(sheet.data)) {
      for (const [column, cell] of sparseArrayEntries(values)) {
        if (cell !== null) result.set(`${row}:${column}`, cell);
      }
    }
    return result;
  }
  for (const entry of sheet.celldata ?? []) {
    if (entry.v !== null) result.set(`${entry.r}:${entry.c}`, entry.v);
  }
  return result;
}

function withoutKeys(
  value: WorkSpreadsheetSheet,
  omitted: ReadonlySet<string>,
): Record<string, unknown> {
  return Object.fromEntries(
    Object.entries(value).filter(([key]) => !omitted.has(key)),
  );
}

function removedConflict(label: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `The ${label} was removed before this change could be applied. Refresh the shared snapshot before retrying.`,
  );
}

function conflict(label: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `The ${label} changed concurrently. Refresh the shared snapshot before retrying.`,
  );
}
