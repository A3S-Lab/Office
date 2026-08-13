import type { Cell, CellWithRowAndCol } from '@fortune-sheet/core';
import {
  OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
  OFFICE_KERNEL_SPREADSHEET_MAX_ROWS,
} from '../kernel/office-kernel-spreadsheet-protocol';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../features/work/work-types';
import { WorkOfficeCollaborationError } from './office-collaboration';
import {
  cloneWorkOfficeCollaborationJson as cloneJsonValue,
  isWorkOfficeCollaborationRecord as isRecord,
} from './office-collaboration-json';

const MAX_POPULATED_CELLS = 1_000_000;
const MAX_DENSE_MATRIX_CELLS = 1_000_000;
const TRANSIENT_SHEET_FIELDS = [
  'calcChain',
  'dynamicArray_compute',
  'luckysheet_select_save',
  'luckysheet_selection_range',
  'status',
  'zoomRatio',
] as const;

export function validateWorkOfficeSpreadsheetContent(
  content: WorkSpreadsheetContent,
): WorkSpreadsheetContent {
  if (!content || content.type !== 'spreadsheet') {
    invalidWorkOfficeSpreadsheetInput('a Spreadsheet content value');
  }
  if (!Array.isArray(content.sheets)) {
    invalidWorkOfficeSpreadsheetInput('an array of sheets');
  }
  const sheets = validateSheets(content.sheets);
  const result: WorkSpreadsheetContent = { type: 'spreadsheet', sheets };
  copyOptionalJsonField(content, result, 'calculation', 'calculation settings');
  if (content.namedRanges !== undefined) {
    result.namedRanges = validateIdRecords(
      content.namedRanges,
      'named range',
    ) as unknown as WorkSpreadsheetContent['namedRanges'];
  }
  copyOptionalSheetSidecars(content, result, 'printAreas', 'print area');
  copyOptionalSheetSidecars(content, result, 'printTitles', 'print titles');
  copyOptionalSheetSidecars(content, result, 'pageBreaks', 'page breaks');
  copyOptionalSheetSidecars(content, result, 'pageSetups', 'page setup');
  assertSpreadsheetReferences(result);
  return result;
}

export function validateSharedWorkOfficeSpreadsheetContent(
  content: WorkSpreadsheetContent,
): WorkSpreadsheetContent {
  try {
    return validateWorkOfficeSpreadsheetContent(content);
  } catch (error) {
    if (error instanceof WorkOfficeCollaborationError) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `The shared Spreadsheet collaboration content is invalid: ${error.message}`,
      );
    }
    throw error;
  }
}

export function invalidWorkOfficeSpreadsheetShared(label: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `The shared Spreadsheet collaboration ${label} is invalid.`,
  );
}

function validateSheets(value: WorkSpreadsheetSheet[]): WorkSpreadsheetSheet[] {
  const ids = new Set<string>();
  const names = new Set<string>();
  return value.map((candidate) => {
    const sheet = validateSheet(candidate);
    const id = sheet.id as string;
    if (ids.has(id)) {
      invalidWorkOfficeSpreadsheetInput(
        `a unique sheet ID; '${id}' is repeated`,
      );
    }
    ids.add(id);
    const normalizedName = sheet.name.toLocaleLowerCase();
    if (names.has(normalizedName)) {
      invalidWorkOfficeSpreadsheetInput(
        `a unique sheet name; '${sheet.name}' is repeated`,
      );
    }
    names.add(normalizedName);
    return sheet;
  });
}

function validateSheet(value: unknown): WorkSpreadsheetSheet {
  const source = requiredInputRecord(value, 'sheet');
  const result = validateJsonRecord(source, 'sheet') as WorkSpreadsheetSheet;
  result.id = requiredIdentifier(source.id, 'sheet');
  result.name = requiredNonEmptyString(source.name, 'sheet name');
  for (const field of TRANSIENT_SHEET_FIELDS) delete result[field];
  if (source.data !== undefined) {
    result.data = validateCellMatrix(source.data, result.id);
    delete result.celldata;
  } else if (source.celldata !== undefined) {
    result.celldata = validateSparseCells(source.celldata, result.id);
    delete result.data;
  }
  validateOptionalDimension(result, 'row', OFFICE_KERNEL_SPREADSHEET_MAX_ROWS);
  validateOptionalDimension(
    result,
    'column',
    OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
  );
  if (source.images !== undefined) {
    result.images = validateIdRecords(
      source.images,
      `image in sheet '${result.id}'`,
    ).filter(
      ({ id }) => !id.startsWith('work-chart-preview-'),
    ) as unknown as NonNullable<WorkSpreadsheetSheet['images']>;
  }
  if (source.charts !== undefined) {
    result.charts = validateIdRecords(
      source.charts,
      `chart in sheet '${result.id}'`,
    ) as unknown as NonNullable<WorkSpreadsheetSheet['charts']>;
  }
  if (source.pivotTables !== undefined) {
    result.pivotTables = validateIdRecords(
      source.pivotTables,
      `pivot table in sheet '${result.id}'`,
    ) as unknown as NonNullable<WorkSpreadsheetSheet['pivotTables']>;
  }
  if (source.formulaMetadata !== undefined) {
    result.formulaMetadata = validateJsonRecord(
      requiredInputRecord(source.formulaMetadata, 'formula metadata'),
      `formula metadata in sheet '${result.id}'`,
    ) as WorkSpreadsheetSheet['formulaMetadata'];
  }
  return result;
}

function validateCellMatrix(
  value: unknown,
  sheetId: string,
): (Cell | null)[][] {
  if (
    !Array.isArray(value) ||
    value.length > OFFICE_KERNEL_SPREADSHEET_MAX_ROWS
  ) {
    invalidWorkOfficeSpreadsheetInput(
      `a cell matrix within the Spreadsheet row limit in sheet '${sheetId}'`,
    );
  }
  let populated = 0;
  let materialized = 0;
  return value.map((candidate, row) => {
    if (
      !Array.isArray(candidate) ||
      candidate.length > OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS
    ) {
      invalidWorkOfficeSpreadsheetInput(
        `a cell row within the Spreadsheet column limit at row ${row} in sheet '${sheetId}'`,
      );
    }
    materialized += candidate.length;
    if (materialized > MAX_DENSE_MATRIX_CELLS) {
      invalidWorkOfficeSpreadsheetInput(
        `at most ${MAX_DENSE_MATRIX_CELLS.toLocaleString()} materialized dense cells in sheet '${sheetId}'; use celldata for larger sparse sheets`,
      );
    }
    return candidate.map((cell, column) => {
      if (cell === null || cell === undefined) return null;
      populated += 1;
      assertPopulatedCellLimit(populated, sheetId);
      return validateCell(cell, row, column, sheetId);
    });
  });
}

function validateSparseCells(
  value: unknown,
  sheetId: string,
): CellWithRowAndCol[] {
  if (!Array.isArray(value)) {
    invalidWorkOfficeSpreadsheetInput(
      `an array of sparse cells in sheet '${sheetId}'`,
    );
  }
  if (value.length > MAX_POPULATED_CELLS) {
    assertPopulatedCellLimit(value.length, sheetId);
  }
  const coordinates = new Set<string>();
  const result = value.map((candidate) => {
    const record = requiredInputRecord(candidate, 'sparse cell');
    const row = requiredCoordinate(
      record.r,
      OFFICE_KERNEL_SPREADSHEET_MAX_ROWS,
      'row',
      sheetId,
    );
    const column = requiredCoordinate(
      record.c,
      OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
      'column',
      sheetId,
    );
    const coordinate = `${row}:${column}`;
    if (coordinates.has(coordinate)) {
      invalidWorkOfficeSpreadsheetInput(
        `a unique sparse cell coordinate; '${coordinate}' is repeated in sheet '${sheetId}'`,
      );
    }
    coordinates.add(coordinate);
    return {
      r: row,
      c: column,
      v:
        record.v === null || record.v === undefined
          ? null
          : validateCell(record.v, row, column, sheetId),
    };
  });
  return result.sort((left, right) => left.r - right.r || left.c - right.c);
}

function validateCell(
  value: unknown,
  row: number,
  column: number,
  sheetId: string,
): Cell {
  const record = requiredInputRecord(value, 'cell');
  return validateJsonRecord(
    record,
    `cell ${row}:${column} in sheet '${sheetId}'`,
  ) as Cell;
}

function validateIdRecords(
  value: unknown,
  label: string,
): Array<Record<string, unknown> & { id: string }> {
  if (!Array.isArray(value)) {
    invalidWorkOfficeSpreadsheetInput(`an array of ${label}s`);
  }
  const ids = new Set<string>();
  return value.map((candidate) => {
    const source = requiredInputRecord(candidate, label);
    const result = validateJsonRecord(source, label) as Record<
      string,
      unknown
    > & { id: string };
    result.id = requiredIdentifier(source.id, label);
    if (ids.has(result.id)) {
      invalidWorkOfficeSpreadsheetInput(
        `a unique ${label} ID; '${result.id}' is repeated`,
      );
    }
    ids.add(result.id);
    return result;
  });
}

function copyOptionalSheetSidecars<
  K extends 'printAreas' | 'printTitles' | 'pageBreaks' | 'pageSetups',
>(
  source: WorkSpreadsheetContent,
  target: WorkSpreadsheetContent,
  key: K,
  label: string,
): void {
  const value = source[key];
  if (value === undefined) return;
  if (!Array.isArray(value)) {
    invalidWorkOfficeSpreadsheetInput(`an array of ${label} records`);
  }
  const sheetIds = new Set<string>();
  const records = value.map((candidate) => {
    const input = requiredInputRecord(candidate, label);
    const record = validateJsonRecord(input, label);
    const sheetId = requiredIdentifier(input.sheetId, `${label} sheet`);
    if (sheetIds.has(sheetId)) {
      invalidWorkOfficeSpreadsheetInput(
        `at most one ${label} record for sheet '${sheetId}'`,
      );
    }
    sheetIds.add(sheetId);
    record.sheetId = sheetId;
    return record;
  });
  (target as unknown as Record<string, unknown>)[key] = records;
}

function copyOptionalJsonField(
  source: WorkSpreadsheetContent,
  target: WorkSpreadsheetContent,
  key: 'calculation',
  label: string,
): void {
  const value = source[key];
  if (value === undefined) return;
  (target as unknown as Record<string, unknown>)[key] = validateJsonRecord(
    requiredInputRecord(value, label),
    label,
  );
}

function assertSpreadsheetReferences(content: WorkSpreadsheetContent): void {
  const sheetIds = new Set(content.sheets.map((sheet) => sheet.id as string));
  for (const range of content.namedRanges ?? []) {
    if (range.scopeSheetId !== undefined && !sheetIds.has(range.scopeSheetId)) {
      invalidWorkOfficeSpreadsheetInput(
        `named range '${range.id}' to reference an existing scope sheet`,
      );
    }
  }
  for (const sheet of content.sheets) {
    for (const pivot of sheet.pivotTables ?? []) {
      if (!sheetIds.has(pivot.sourceSheetId)) {
        invalidWorkOfficeSpreadsheetInput(
          `pivot table '${pivot.id}' to reference an existing source sheet`,
        );
      }
    }
  }
  for (const key of [
    'printAreas',
    'printTitles',
    'pageBreaks',
    'pageSetups',
  ] as const) {
    for (const record of content[key] ?? []) {
      if (!sheetIds.has(record.sheetId)) {
        invalidWorkOfficeSpreadsheetInput(
          `${key} to reference an existing sheet '${record.sheetId}'`,
        );
      }
    }
  }
}

function validateOptionalDimension(
  sheet: WorkSpreadsheetSheet,
  key: 'row' | 'column',
  maximum: number,
): void {
  const value = sheet[key];
  if (value === undefined) return;
  if (!Number.isSafeInteger(value) || value < 0 || value > maximum) {
    invalidWorkOfficeSpreadsheetInput(
      `a ${key} count between 0 and ${maximum} in sheet '${sheet.id}'`,
    );
  }
}

function requiredCoordinate(
  value: unknown,
  maximum: number,
  label: string,
  sheetId: string,
): number {
  if (
    !Number.isSafeInteger(value) ||
    (value as number) < 0 ||
    (value as number) >= maximum
  ) {
    invalidWorkOfficeSpreadsheetInput(
      `a valid ${label} coordinate in sheet '${sheetId}'`,
    );
  }
  return value as number;
}

function requiredInputRecord(
  value: unknown,
  label: string,
): Record<string, unknown> {
  if (!isRecord(value)) {
    invalidWorkOfficeSpreadsheetInput(`a valid ${label} record`);
  }
  return value as Record<string, unknown>;
}

function requiredIdentifier(value: unknown, label: string): string {
  const result = requiredNonEmptyString(value, `${label} ID`);
  if (result !== result.trim() || result.length > 256) {
    invalidWorkOfficeSpreadsheetInput(
      `a ${label} ID containing 1 to 256 characters`,
    );
  }
  return result;
}

function requiredNonEmptyString(value: unknown, label: string): string {
  if (typeof value !== 'string' || !value.trim() || value.length > 256) {
    invalidWorkOfficeSpreadsheetInput(
      `a non-empty string of at most 256 characters for ${label}`,
    );
  }
  return value as string;
}

function validateJsonRecord(
  value: Record<string, unknown>,
  label: string,
): Record<string, unknown> {
  try {
    return cloneJsonValue(value) as Record<string, unknown>;
  } catch (error) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.content_invalid',
      `Spreadsheet collaboration requires a JSON-compatible ${label}: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
}

function assertPopulatedCellLimit(count: number, sheetId: string): void {
  if (count <= MAX_POPULATED_CELLS) return;
  invalidWorkOfficeSpreadsheetInput(
    `at most ${MAX_POPULATED_CELLS.toLocaleString()} populated cells in sheet '${sheetId}'`,
  );
}

function invalidWorkOfficeSpreadsheetInput(expected: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `Spreadsheet collaboration requires ${expected}.`,
  );
}
