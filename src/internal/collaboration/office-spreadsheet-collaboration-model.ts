import * as Y from 'yjs';
import type { WorkSpreadsheetContent } from '../features/work/work-types';
import { WorkOfficeCollaborationError } from './office-collaboration';
import { workOfficeCollaborationJsonEqual as jsonEqual } from './office-collaboration-json';
import {
  appendWorkOfficeSpreadsheetRecordClaims,
  assertWorkOfficeSpreadsheetRecordClaims,
} from './office-spreadsheet-collaboration-claims';
import { assertWorkOfficeSpreadsheetPatchSafe } from './office-spreadsheet-collaboration-conflicts';
import {
  patchSpreadsheetFlatJsonMap,
  readSpreadsheetFlatJsonMap,
} from './office-spreadsheet-collaboration-flat-json';
import {
  initializeSpreadsheetSheetRecord,
  patchSpreadsheetOrder,
  patchSpreadsheetSheetRecord,
  readSpreadsheetSheetRecord,
  requiredSpreadsheetMap,
  validatedSpreadsheetOrder,
} from './office-spreadsheet-collaboration-records';
import {
  invalidWorkOfficeSpreadsheetShared as invalidSharedSpreadsheet,
  validateSharedWorkOfficeSpreadsheetContent,
} from './office-spreadsheet-collaboration-validation';

export { validateWorkOfficeSpreadsheetContent } from './office-spreadsheet-collaboration-validation';

export const SPREADSHEET_OPTIONS_ROOT = 'spreadsheet.options';
export const SPREADSHEET_SHEETS_ROOT = 'spreadsheet.sheets';
export const SPREADSHEET_SHEET_ORDER_ROOT = 'spreadsheet.sheet-order';
export const SPREADSHEET_NAMED_RANGES_ROOT = 'spreadsheet.named-ranges';
export const SPREADSHEET_NAMED_RANGE_ORDER_ROOT =
  'spreadsheet.named-range-order';
export const SPREADSHEET_RECORD_CLAIMS_ROOT = 'spreadsheet.record-claims';
export const SPREADSHEET_PRINT_AREAS_ROOT = 'spreadsheet.print-areas';
export const SPREADSHEET_PRINT_TITLES_ROOT = 'spreadsheet.print-titles';
export const SPREADSHEET_PAGE_BREAKS_ROOT = 'spreadsheet.page-breaks';
export const SPREADSHEET_PAGE_SETUPS_ROOT = 'spreadsheet.page-setups';

export interface WorkOfficeSpreadsheetRoots {
  options: Y.Map<unknown>;
  sheets: Y.Map<unknown>;
  sheetOrder: Y.Array<string>;
  namedRanges: Y.Map<unknown>;
  namedRangeOrder: Y.Array<string>;
  recordClaims: Y.Array<string>;
  printAreas: Y.Map<unknown>;
  printTitles: Y.Map<unknown>;
  pageBreaks: Y.Map<unknown>;
  pageSetups: Y.Map<unknown>;
}

export function workOfficeSpreadsheetRoots(
  document: Y.Doc,
  rootName: (suffix: string) => string,
): WorkOfficeSpreadsheetRoots {
  return {
    options: document.getMap(rootName(SPREADSHEET_OPTIONS_ROOT)),
    sheets: document.getMap(rootName(SPREADSHEET_SHEETS_ROOT)),
    sheetOrder: document.getArray(rootName(SPREADSHEET_SHEET_ORDER_ROOT)),
    namedRanges: document.getMap(rootName(SPREADSHEET_NAMED_RANGES_ROOT)),
    namedRangeOrder: document.getArray(
      rootName(SPREADSHEET_NAMED_RANGE_ORDER_ROOT),
    ),
    recordClaims: document.getArray(rootName(SPREADSHEET_RECORD_CLAIMS_ROOT)),
    printAreas: document.getMap(rootName(SPREADSHEET_PRINT_AREAS_ROOT)),
    printTitles: document.getMap(rootName(SPREADSHEET_PRINT_TITLES_ROOT)),
    pageBreaks: document.getMap(rootName(SPREADSHEET_PAGE_BREAKS_ROOT)),
    pageSetups: document.getMap(rootName(SPREADSHEET_PAGE_SETUPS_ROOT)),
  };
}

export function initializeWorkOfficeSpreadsheetRoots(
  roots: WorkOfficeSpreadsheetRoots,
  content: WorkSpreadsheetContent,
): void {
  appendWorkOfficeSpreadsheetRecordClaims(
    roots.recordClaims,
    undefined,
    content,
  );
  patchSpreadsheetFlatJsonMap(
    roots.options,
    undefined,
    spreadsheetWorkbookOptions(content),
    'workbook options',
  );
  patchSheets(roots, undefined, content);
  patchIdRecords(
    roots.namedRanges,
    roots.namedRangeOrder,
    [],
    content.namedRanges ?? [],
    'named range',
  );
  patchSheetSidecars(
    roots.printAreas,
    [],
    content.printAreas ?? [],
    'print area',
  );
  patchSheetSidecars(
    roots.printTitles,
    [],
    content.printTitles ?? [],
    'print titles',
  );
  patchSheetSidecars(
    roots.pageBreaks,
    [],
    content.pageBreaks ?? [],
    'page breaks',
  );
  patchSheetSidecars(
    roots.pageSetups,
    [],
    content.pageSetups ?? [],
    'page setup',
  );
}

export function readWorkOfficeSpreadsheetRoots(
  roots: WorkOfficeSpreadsheetRoots,
): WorkSpreadsheetContent {
  const result: WorkSpreadsheetContent = {
    type: 'spreadsheet',
    sheets: validatedSpreadsheetOrder(
      roots.sheetOrder,
      roots.sheets,
      'sheet',
    ).map((id) =>
      readSpreadsheetSheetRecord(
        requiredSpreadsheetMap(roots.sheets, id, `sheet '${id}'`),
        id,
      ),
    ),
  };
  const options = readSpreadsheetFlatJsonMap(roots.options, 'workbook options');
  const { dateSystem, ...calculation } = options;
  if (dateSystem !== undefined) {
    result.dateSystem = dateSystem as WorkSpreadsheetContent['dateSystem'];
  }
  if (Object.keys(calculation).length > 0) {
    result.calculation =
      calculation as unknown as WorkSpreadsheetContent['calculation'];
  }
  const namedRanges = readIdRecords(
    roots.namedRanges,
    roots.namedRangeOrder,
    'named range',
  );
  if (namedRanges.length > 0) {
    result.namedRanges = namedRanges as unknown as NonNullable<
      WorkSpreadsheetContent['namedRanges']
    >;
  }
  readSheetSidecars(roots.printAreas, 'printAreas', 'print area', result);
  readSheetSidecars(roots.printTitles, 'printTitles', 'print titles', result);
  readSheetSidecars(roots.pageBreaks, 'pageBreaks', 'page breaks', result);
  readSheetSidecars(roots.pageSetups, 'pageSetups', 'page setup', result);
  const validated = validateSharedWorkOfficeSpreadsheetContent(result);
  assertWorkOfficeSpreadsheetRecordClaims(roots.recordClaims, validated);
  return validated;
}

export function patchWorkOfficeSpreadsheetRoots(
  roots: WorkOfficeSpreadsheetRoots,
  previous: WorkSpreadsheetContent,
  next: WorkSpreadsheetContent,
): void {
  const shared = readWorkOfficeSpreadsheetRoots(roots);
  assertWorkOfficeSpreadsheetPatchSafe(previous, next, shared);
  appendWorkOfficeSpreadsheetRecordClaims(roots.recordClaims, previous, next);
  patchSpreadsheetFlatJsonMap(
    roots.options,
    spreadsheetWorkbookOptions(previous),
    spreadsheetWorkbookOptions(next),
    'workbook options',
  );
  patchSheets(roots, previous, next);
  patchIdRecords(
    roots.namedRanges,
    roots.namedRangeOrder,
    previous.namedRanges ?? [],
    next.namedRanges ?? [],
    'named range',
  );
  patchSheetSidecars(
    roots.printAreas,
    previous.printAreas ?? [],
    next.printAreas ?? [],
    'print area',
  );
  patchSheetSidecars(
    roots.printTitles,
    previous.printTitles ?? [],
    next.printTitles ?? [],
    'print titles',
  );
  patchSheetSidecars(
    roots.pageBreaks,
    previous.pageBreaks ?? [],
    next.pageBreaks ?? [],
    'page breaks',
  );
  patchSheetSidecars(
    roots.pageSetups,
    previous.pageSetups ?? [],
    next.pageSetups ?? [],
    'page setup',
  );
}

function spreadsheetWorkbookOptions(
  content: WorkSpreadsheetContent,
): Record<string, unknown> | undefined {
  const options: Record<string, unknown> = {
    ...(content.calculation as Record<string, unknown> | undefined),
  };
  if (content.dateSystem !== undefined) {
    options.dateSystem = content.dateSystem;
  }
  return Object.keys(options).length ? options : undefined;
}

export function assertWorkOfficeSpreadsheetRootsEmpty(
  roots: WorkOfficeSpreadsheetRoots,
): void {
  if (
    roots.options.size > 0 ||
    roots.sheets.size > 0 ||
    roots.sheetOrder.length > 0 ||
    roots.namedRanges.size > 0 ||
    roots.namedRangeOrder.length > 0 ||
    roots.recordClaims.length > 0 ||
    roots.printAreas.size > 0 ||
    roots.printTitles.size > 0 ||
    roots.pageBreaks.size > 0 ||
    roots.pageSetups.size > 0
  ) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.bootstrap_ambiguous',
      'The Spreadsheet collaboration roots contain data without initialized metadata.',
    );
  }
}

export function workOfficeSpreadsheetUndoScope(
  roots: WorkOfficeSpreadsheetRoots,
): Array<Y.Map<unknown> | Y.Array<string>> {
  const scope: Array<Y.Map<unknown> | Y.Array<string>> = [
    roots.options,
    roots.sheets,
    roots.sheetOrder,
    roots.namedRanges,
    roots.namedRangeOrder,
    roots.recordClaims,
    roots.printAreas,
    roots.printTitles,
    roots.pageBreaks,
    roots.pageSetups,
  ];
  return scope;
}

function patchSheets(
  roots: WorkOfficeSpreadsheetRoots,
  previous: WorkSpreadsheetContent | undefined,
  next: WorkSpreadsheetContent,
): void {
  const before = previous?.sheets ?? [];
  const beforeById = new Map(
    before.map((sheet) => [sheet.id as string, sheet]),
  );
  const afterById = new Map(
    next.sheets.map((sheet) => [sheet.id as string, sheet]),
  );
  for (const sheet of before) {
    const id = sheet.id as string;
    if (afterById.has(id)) continue;
    const current = roots.sheets.get(id);
    if (current === undefined) continue;
    if (!(current instanceof Y.Map)) invalidSharedSpreadsheet(`sheet '${id}'`);
    if (
      !jsonEqual(
        sheet,
        readSpreadsheetSheetRecord(current as Y.Map<unknown>, id),
      )
    ) {
      throw staleConflict(`Sheet '${id}' changed before it was deleted`);
    }
    roots.sheets.delete(id);
    removeAll(roots.sheetOrder, id);
  }
  for (const sheet of next.sheets) {
    const id = sheet.id as string;
    const beforeSheet = beforeById.get(id);
    if (beforeSheet && jsonEqual(beforeSheet, sheet)) continue;
    let record = roots.sheets.get(id);
    if (record === undefined) {
      if (beforeSheet)
        throw staleConflict(`Sheet '${id}' was removed before this change`);
      record = new Y.Map<unknown>();
      roots.sheets.set(id, record);
      initializeSpreadsheetSheetRecord(record as Y.Map<unknown>, sheet);
      continue;
    }
    if (!(record instanceof Y.Map)) invalidSharedSpreadsheet(`sheet '${id}'`);
    patchSpreadsheetSheetRecord(record as Y.Map<unknown>, beforeSheet, sheet);
  }
  patchSpreadsheetOrder(
    roots.sheetOrder,
    before.map((sheet) => sheet.id as string),
    next.sheets.map((sheet) => sheet.id as string),
  );
}

function patchIdRecords<T extends { id: string }>(
  records: Y.Map<unknown>,
  order: Y.Array<string>,
  previous: T[],
  next: T[],
  label: string,
): void {
  const beforeById = new Map(previous.map((value) => [value.id, value]));
  const afterById = new Map(next.map((value) => [value.id, value]));
  for (const value of previous) {
    if (afterById.has(value.id)) continue;
    const current = records.get(value.id);
    if (current === undefined) continue;
    if (!(current instanceof Y.Map)) invalidSharedSpreadsheet(label);
    if (!jsonEqual(value, readSpreadsheetFlatJsonMap(current, label))) {
      throw staleConflict(
        `${label} '${value.id}' changed before it was deleted`,
      );
    }
    records.delete(value.id);
    removeAll(order, value.id);
  }
  for (const value of next) {
    const before = beforeById.get(value.id);
    if (before && jsonEqual(before, value)) continue;
    let record = records.get(value.id);
    if (record === undefined) {
      if (before)
        throw staleConflict(
          `${label} '${value.id}' was removed before this change`,
        );
      record = new Y.Map<unknown>();
      records.set(value.id, record);
    }
    if (!(record instanceof Y.Map)) invalidSharedSpreadsheet(label);
    patchSpreadsheetFlatJsonMap(
      record as Y.Map<unknown>,
      before as Record<string, unknown> | undefined,
      value as Record<string, unknown>,
      label,
    );
  }
  patchSpreadsheetOrder(
    order,
    previous.map(({ id }) => id),
    next.map(({ id }) => id),
  );
}

function readIdRecords(
  records: Y.Map<unknown>,
  order: Y.Array<string>,
  label: string,
): Array<Record<string, unknown> & { id: string }> {
  return validatedSpreadsheetOrder(order, records, label).map((id) => {
    const value = readSpreadsheetFlatJsonMap(
      requiredSpreadsheetMap(records, id, label),
      label,
    ) as Record<string, unknown> & { id: string };
    if (value.id !== id) invalidSharedSpreadsheet(`${label} identity`);
    return value;
  });
}

function patchSheetSidecars<T extends { sheetId: string }>(
  records: Y.Map<unknown>,
  previous: T[],
  next: T[],
  label: string,
): void {
  const beforeBySheet = new Map(
    previous.map((value) => [value.sheetId, value]),
  );
  const afterBySheet = new Map(next.map((value) => [value.sheetId, value]));
  for (const value of previous) {
    if (afterBySheet.has(value.sheetId)) continue;
    const current = records.get(value.sheetId);
    if (current === undefined) continue;
    if (!(current instanceof Y.Map)) invalidSharedSpreadsheet(label);
    if (!jsonEqual(value, readSpreadsheetFlatJsonMap(current, label))) {
      throw staleConflict(
        `${label} for sheet '${value.sheetId}' changed before deletion`,
      );
    }
    records.delete(value.sheetId);
  }
  for (const value of next) {
    const before = beforeBySheet.get(value.sheetId);
    if (before && jsonEqual(before, value)) continue;
    let record = records.get(value.sheetId);
    if (record === undefined) {
      if (before)
        throw staleConflict(
          `${label} for sheet '${value.sheetId}' was removed`,
        );
      record = new Y.Map<unknown>();
      records.set(value.sheetId, record);
    }
    if (!(record instanceof Y.Map)) invalidSharedSpreadsheet(label);
    patchSpreadsheetFlatJsonMap(
      record as Y.Map<unknown>,
      before as Record<string, unknown> | undefined,
      value as Record<string, unknown>,
      label,
    );
  }
}

function readSheetSidecars(
  records: Y.Map<unknown>,
  key: 'printAreas' | 'printTitles' | 'pageBreaks' | 'pageSetups',
  label: string,
  output: WorkSpreadsheetContent,
): void {
  const result = Array.from(records.keys())
    .sort()
    .map((sheetId) => {
      const value = readSpreadsheetFlatJsonMap(
        requiredSpreadsheetMap(records, sheetId, label),
        label,
      );
      if (value.sheetId !== sheetId)
        invalidSharedSpreadsheet(`${label} identity`);
      return value;
    });
  if (result.length > 0) {
    (output as unknown as Record<string, unknown>)[key] = result;
  }
}

function removeAll(order: Y.Array<string>, id: string): void {
  for (let index = order.length - 1; index >= 0; index -= 1) {
    if (order.get(index) === id) order.delete(index, 1);
  }
}

function staleConflict(message: string): WorkOfficeCollaborationError {
  return new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `${message}. Refresh the shared snapshot before retrying.`,
  );
}
