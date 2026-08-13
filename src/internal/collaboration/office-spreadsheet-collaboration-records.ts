import * as Y from 'yjs';
import type { WorkSpreadsheetSheet } from '../features/work/work-types';
import { WorkOfficeCollaborationError } from './office-collaboration';
import {
  cloneWorkOfficeCollaborationJson as cloneJsonValue,
  workOfficeCollaborationJsonEqual as jsonEqual,
} from './office-collaboration-json';
import {
  patchSpreadsheetFlatJsonMap,
  readSpreadsheetFlatJsonMap,
} from './office-spreadsheet-collaboration-flat-json';
import {
  patchSpreadsheetCells,
  readSpreadsheetCells,
  SPREADSHEET_RECORD_CELLS,
  SPREADSHEET_RECORD_CELL_MODE,
  SPREADSHEET_RECORD_CELL_PRESENCE,
  SPREADSHEET_RECORD_DATA_ROW_LENGTHS,
} from './office-spreadsheet-collaboration-cells';
import { invalidWorkOfficeSpreadsheetShared as invalidSharedSpreadsheet } from './office-spreadsheet-collaboration-validation';

export const SPREADSHEET_RECORD_CONFIG = 'config';
export const SPREADSHEET_RECORD_FORMULA_METADATA = 'formulaMetadata';
export const SPREADSHEET_RECORD_IMAGES = 'images';
export const SPREADSHEET_RECORD_IMAGE_ORDER = 'imageOrder';
export const SPREADSHEET_RECORD_CHARTS = 'charts';
export const SPREADSHEET_RECORD_CHART_ORDER = 'chartOrder';
export const SPREADSHEET_RECORD_PIVOTS = 'pivotTables';
export const SPREADSHEET_RECORD_PIVOT_ORDER = 'pivotOrder';

const NESTED_SHEET_FIELDS = new Set([
  SPREADSHEET_RECORD_CELLS,
  SPREADSHEET_RECORD_CELL_PRESENCE,
  SPREADSHEET_RECORD_CELL_MODE,
  SPREADSHEET_RECORD_DATA_ROW_LENGTHS,
  SPREADSHEET_RECORD_CONFIG,
  SPREADSHEET_RECORD_FORMULA_METADATA,
  SPREADSHEET_RECORD_IMAGES,
  SPREADSHEET_RECORD_IMAGE_ORDER,
  SPREADSHEET_RECORD_CHARTS,
  SPREADSHEET_RECORD_CHART_ORDER,
  SPREADSHEET_RECORD_PIVOTS,
  SPREADSHEET_RECORD_PIVOT_ORDER,
]);

export function initializeSpreadsheetSheetRecord(
  record: Y.Map<unknown>,
  sheet: WorkSpreadsheetSheet,
): void {
  nestedSpreadsheetMap(record, SPREADSHEET_RECORD_CONFIG, 'sheet config');
  nestedSpreadsheetMap(
    record,
    SPREADSHEET_RECORD_FORMULA_METADATA,
    'sheet formula metadata',
  );
  initializeIdCollection(
    record,
    SPREADSHEET_RECORD_IMAGES,
    SPREADSHEET_RECORD_IMAGE_ORDER,
  );
  initializeIdCollection(
    record,
    SPREADSHEET_RECORD_CHARTS,
    SPREADSHEET_RECORD_CHART_ORDER,
  );
  initializeIdCollection(
    record,
    SPREADSHEET_RECORD_PIVOTS,
    SPREADSHEET_RECORD_PIVOT_ORDER,
  );
  patchSpreadsheetSheetRecord(record, undefined, sheet);
}

export function patchSpreadsheetSheetRecord(
  record: Y.Map<unknown>,
  previous: WorkSpreadsheetSheet | undefined,
  next: WorkSpreadsheetSheet,
): void {
  const before = (previous ?? {}) as Record<string, unknown>;
  const after = next as Record<string, unknown>;
  const keys = new Set([...Object.keys(before), ...Object.keys(after)]);
  for (const key of keys) {
    if (
      key === 'data' ||
      key === 'celldata' ||
      key === 'config' ||
      key === 'formulaMetadata' ||
      key === 'images' ||
      key === 'charts' ||
      key === 'pivotTables'
    ) {
      continue;
    }
    patchOptionalJson(
      record,
      key,
      before[key],
      after[key],
      previous !== undefined,
    );
  }
  patchJsonMap(
    nestedSpreadsheetMap(record, SPREADSHEET_RECORD_CONFIG, 'sheet config'),
    previous?.config as Record<string, unknown> | undefined,
    next.config as Record<string, unknown> | undefined,
  );
  patchJsonMap(
    nestedSpreadsheetMap(
      record,
      SPREADSHEET_RECORD_FORMULA_METADATA,
      'sheet formula metadata',
    ),
    previous?.formulaMetadata as Record<string, unknown> | undefined,
    next.formulaMetadata as Record<string, unknown> | undefined,
  );
  patchSpreadsheetCells(record, previous, next);
  patchIdCollection(
    record,
    SPREADSHEET_RECORD_IMAGES,
    SPREADSHEET_RECORD_IMAGE_ORDER,
    previous?.images ?? [],
    next.images ?? [],
    `image in sheet '${next.id}'`,
  );
  patchIdCollection(
    record,
    SPREADSHEET_RECORD_CHARTS,
    SPREADSHEET_RECORD_CHART_ORDER,
    previous?.charts ?? [],
    next.charts ?? [],
    `chart in sheet '${next.id}'`,
  );
  patchIdCollection(
    record,
    SPREADSHEET_RECORD_PIVOTS,
    SPREADSHEET_RECORD_PIVOT_ORDER,
    previous?.pivotTables ?? [],
    next.pivotTables ?? [],
    `pivot table in sheet '${next.id}'`,
  );
}

export function readSpreadsheetSheetRecord(
  record: Y.Map<unknown>,
  id: string,
): WorkSpreadsheetSheet {
  const result: Record<string, unknown> = {};
  for (const [key, value] of record.entries()) {
    if (NESTED_SHEET_FIELDS.has(key)) continue;
    if (value instanceof Y.AbstractType) {
      invalidSharedSpreadsheet(`sheet '${id}' field '${key}'`);
    }
    result[key] = cloneJsonValue(value);
  }
  if (result.id !== id) invalidSharedSpreadsheet(`sheet '${id}' identity`);
  const config = readJsonMap(
    requiredSpreadsheetMap(record, SPREADSHEET_RECORD_CONFIG, 'sheet config'),
    'sheet config',
  );
  if (Object.keys(config).length > 0) result.config = config;
  const formulaMetadata = readJsonMap(
    requiredSpreadsheetMap(
      record,
      SPREADSHEET_RECORD_FORMULA_METADATA,
      'sheet formula metadata',
    ),
    'sheet formula metadata',
  );
  if (Object.keys(formulaMetadata).length > 0) {
    result.formulaMetadata = formulaMetadata;
  }
  const sheet = result as WorkSpreadsheetSheet;
  readSpreadsheetCells(record, sheet);
  readOptionalCollection(
    record,
    SPREADSHEET_RECORD_IMAGES,
    SPREADSHEET_RECORD_IMAGE_ORDER,
    `image in sheet '${id}'`,
    'images',
    result,
  );
  readOptionalCollection(
    record,
    SPREADSHEET_RECORD_CHARTS,
    SPREADSHEET_RECORD_CHART_ORDER,
    `chart in sheet '${id}'`,
    'charts',
    result,
  );
  readOptionalCollection(
    record,
    SPREADSHEET_RECORD_PIVOTS,
    SPREADSHEET_RECORD_PIVOT_ORDER,
    `pivot table in sheet '${id}'`,
    'pivotTables',
    result,
  );
  return sheet;
}

export function patchSpreadsheetOrder(
  order: Y.Array<string>,
  previous: string[],
  next: string[],
): void {
  if (jsonEqual(previous, next)) return;
  const nextIds = new Set(next);
  for (const id of previous) {
    if (!nextIds.has(id)) removeAll(order, id);
  }
  const previousIds = new Set(previous);
  for (const id of next) {
    if (!previousIds.has(id) && !order.toArray().includes(id)) {
      insertRelative(order, next, id);
    }
  }
  if (!sameRelativeOrder(previous, next)) {
    for (const id of next.filter((value) => previousIds.has(value))) {
      removeAll(order, id);
    }
    for (const id of next) {
      if (!order.toArray().includes(id)) insertRelative(order, next, id);
    }
  }
}

export function validatedSpreadsheetOrder(
  order: Y.Array<string>,
  records: Y.Map<unknown>,
  label: string,
): string[] {
  const seen = new Set<string>();
  const values: string[] = [];
  for (const value of order.toArray()) {
    if (typeof value !== 'string' || !value.trim()) {
      invalidSharedSpreadsheet(`${label} order`);
    }
    if (seen.has(value)) continue;
    seen.add(value);
    values.push(value);
  }
  if (
    seen.size !== records.size ||
    Array.from(records.keys()).some((id) => !seen.has(id))
  ) {
    invalidSharedSpreadsheet(`${label} order and record set`);
  }
  return values;
}

export function nestedSpreadsheetMap(
  parent: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Map<unknown> {
  const existing = parent.get(key);
  if (existing instanceof Y.Map) return existing;
  if (existing !== undefined) invalidSharedSpreadsheet(label);
  const value = new Y.Map<unknown>();
  parent.set(key, value);
  return value;
}

export function requiredSpreadsheetMap(
  parent: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Map<unknown> {
  const value = parent.get(key);
  if (!(value instanceof Y.Map)) invalidSharedSpreadsheet(label);
  return value as Y.Map<unknown>;
}

function initializeIdCollection(
  record: Y.Map<unknown>,
  recordsKey: string,
  orderKey: string,
): void {
  nestedSpreadsheetMap(record, recordsKey, recordsKey);
  nestedSpreadsheetArray(record, orderKey, orderKey);
}

function patchIdCollection<T extends { id: string }>(
  parent: Y.Map<unknown>,
  recordsKey: string,
  orderKey: string,
  previous: T[],
  next: T[],
  label: string,
): void {
  const records = nestedSpreadsheetMap(parent, recordsKey, label);
  const order = nestedSpreadsheetArray(parent, orderKey, `${label} order`);
  const previousById = new Map(previous.map((value) => [value.id, value]));
  const nextById = new Map(next.map((value) => [value.id, value]));
  for (const value of previous) {
    if (nextById.has(value.id)) continue;
    const current = records.get(value.id);
    if (!(current instanceof Y.Map)) {
      if (current === undefined) continue;
      invalidSharedSpreadsheet(label);
    }
    if (!jsonEqual(value, readJsonMap(current as Y.Map<unknown>, label))) {
      throw staleConflict(
        `${label} '${value.id}' changed before it was deleted`,
      );
    }
    records.delete(value.id);
    removeAll(order, value.id);
  }
  for (const value of next) {
    const before = previousById.get(value.id);
    if (before && jsonEqual(before, value)) continue;
    let record = records.get(value.id);
    if (record === undefined) {
      if (before) {
        throw staleConflict(
          `${label} '${value.id}' was removed before this change`,
        );
      }
      record = new Y.Map<unknown>();
      records.set(value.id, record);
    }
    if (!(record instanceof Y.Map)) invalidSharedSpreadsheet(label);
    patchJsonMap(
      record as Y.Map<unknown>,
      before as Record<string, unknown> | undefined,
      value as Record<string, unknown>,
    );
  }
  patchSpreadsheetOrder(
    order,
    previous.map(({ id }) => id),
    next.map(({ id }) => id),
  );
}

function readOptionalCollection(
  parent: Y.Map<unknown>,
  recordsKey: string,
  orderKey: string,
  label: string,
  outputKey: string,
  output: Record<string, unknown>,
): void {
  const records = requiredSpreadsheetMap(parent, recordsKey, label);
  const order = requiredSpreadsheetArray(parent, orderKey, `${label} order`);
  const result = validatedSpreadsheetOrder(order, records, label).map((id) => {
    const record = requiredSpreadsheetMap(records, id, label);
    const value = readJsonMap(record, label);
    if (value.id !== id) invalidSharedSpreadsheet(`${label} identity`);
    return value;
  });
  if (result.length > 0) output[outputKey] = result;
}

function patchJsonMap(
  target: Y.Map<unknown>,
  previous: Record<string, unknown> | undefined,
  next: Record<string, unknown> | undefined,
): void {
  patchSpreadsheetFlatJsonMap(target, previous, next, 'Spreadsheet record');
}

function readJsonMap(
  target: Y.Map<unknown>,
  label: string,
): Record<string, unknown> {
  return readSpreadsheetFlatJsonMap(target, label);
}

function patchOptionalJson(
  target: Y.Map<unknown>,
  key: string,
  previous: unknown,
  next: unknown,
  hadPrevious: boolean,
): void {
  if (hadPrevious && jsonEqual(previous, next)) return;
  if (next === undefined) {
    if (hadPrevious && !jsonEqual(target.get(key), previous)) {
      throw staleConflict(`Field '${key}' changed before it was removed`);
    }
    target.delete(key);
  } else {
    target.set(key, cloneJsonValue(next));
  }
}

function nestedSpreadsheetArray(
  parent: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Array<string> {
  const existing = parent.get(key);
  if (existing instanceof Y.Array) return existing as Y.Array<string>;
  if (existing !== undefined) invalidSharedSpreadsheet(label);
  const value = new Y.Array<string>();
  parent.set(key, value);
  return value;
}

function requiredSpreadsheetArray(
  parent: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Array<string> {
  const value = parent.get(key);
  if (!(value instanceof Y.Array)) invalidSharedSpreadsheet(label);
  return value as Y.Array<string>;
}

function removeAll(order: Y.Array<string>, id: string): void {
  for (let index = order.length - 1; index >= 0; index -= 1) {
    if (order.get(index) === id) order.delete(index, 1);
  }
}

function insertRelative(
  order: Y.Array<string>,
  desired: string[],
  id: string,
): void {
  const desiredIndex = desired.indexOf(id);
  for (let index = desiredIndex - 1; index >= 0; index -= 1) {
    const sibling = desired[index];
    const position = order.toArray().lastIndexOf(sibling);
    if (position >= 0) {
      order.insert(position + 1, [id]);
      return;
    }
  }
  for (let index = desiredIndex + 1; index < desired.length; index += 1) {
    const sibling = desired[index];
    const position = order.toArray().indexOf(sibling);
    if (position >= 0) {
      order.insert(position, [id]);
      return;
    }
  }
  order.push([id]);
}

function sameRelativeOrder(previous: string[], next: string[]): boolean {
  const nextIds = new Set(next);
  const previousIds = new Set(previous);
  return jsonEqual(
    previous.filter((id) => nextIds.has(id)),
    next.filter((id) => previousIds.has(id)),
  );
}

function staleConflict(message: string): WorkOfficeCollaborationError {
  return new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `${message}. Refresh the shared snapshot before retrying.`,
  );
}
