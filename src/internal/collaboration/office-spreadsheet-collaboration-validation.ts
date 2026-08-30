import type { Cell, CellWithRowAndCol } from '@fortune-sheet/core';
import { normalizeSpreadsheetTableCalculatedFormula } from '../features/work/editors/spreadsheet-table-calculated-columns';
import {
  normalizeSpreadsheetTableTotalsFormula,
  normalizeSpreadsheetTableTotalsFunction,
  normalizeSpreadsheetTableTotalsLabel,
} from '../features/work/editors/spreadsheet-table-totals';
import { isValidSpreadsheetDefinedName } from '../features/work/work-spreadsheet-ranges';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetCustomFilterCondition,
  WorkSpreadsheetDynamicFilter,
  WorkSpreadsheetSheet,
  WorkSpreadsheetTable,
  WorkSpreadsheetTableFilterCriteria,
} from '../features/work/work-types';
import {
  OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
  OFFICE_KERNEL_SPREADSHEET_MAX_ROWS,
} from '../kernel/office-kernel-spreadsheet-protocol';
import { WorkOfficeCollaborationError } from './office-collaboration';
import {
  invalidWorkOfficeSpreadsheetInput,
  requiredCoordinate,
  requiredIdentifier,
  requiredInputRecord,
  requiredNonEmptyString,
  validateJsonRecord,
} from './office-spreadsheet-collaboration-validation-support';

const MAX_POPULATED_CELLS = 1_000_000;
const MAX_DENSE_MATRIX_CELLS = 1_000_000;
const MAX_FILTER_VALUES_PER_COLUMN = 10_000;
const MAX_FILTER_VALUE_CHARACTERS = 32_767;
const MAX_FILTER_TEXT_BYTES = 1_048_576;
const FILTER_TEXT_ENCODER = new TextEncoder();
const SPREADSHEET_DYNAMIC_FILTERS: ReadonlySet<string> = new Set([
  'above-average',
  'below-average',
  'tomorrow',
  'today',
  'yesterday',
  'next-week',
  'this-week',
  'last-week',
  'next-month',
  'this-month',
  'last-month',
  'next-quarter',
  'this-quarter',
  'last-quarter',
  'next-year',
  'this-year',
  'last-year',
  'year-to-date',
  'quarter-1',
  'quarter-2',
  'quarter-3',
  'quarter-4',
  'month-1',
  'month-2',
  'month-3',
  'month-4',
  'month-5',
  'month-6',
  'month-7',
  'month-8',
  'month-9',
  'month-10',
  'month-11',
  'month-12',
]);
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
  if (source.tables !== undefined) {
    result.tables = validateSpreadsheetTables(source.tables, result.id);
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

function validateSpreadsheetTables(
  value: unknown,
  sheetId: string,
): WorkSpreadsheetTable[] {
  const records = validateIdRecords(value, `table in sheet '${sheetId}'`);
  return records.map((record) => {
    const label = `table '${record.id}' in sheet '${sheetId}'`;
    const range = requiredSpreadsheetTableRange(record.range, label, sheetId);
    const width = range.column[1] - range.column[0] + 1;
    const columns = requiredSpreadsheetTableColumns(
      record.columns,
      width,
      label,
    );
    const filters = requiredSpreadsheetTableFilters(
      record.filters,
      width,
      label,
    );
    const headerRow = requiredBoolean(record.headerRow, `${label} header row`);
    const totalsRow = requiredBoolean(record.totalsRow, `${label} totals row`);
    const height = range.row[1] - range.row[0] + 1;
    if (height <= Number(headerRow) + Number(totalsRow)) {
      invalidWorkOfficeSpreadsheetInput(
        `${label} to include at least one body row`,
      );
    }
    const result = record as unknown as WorkSpreadsheetTable;
    result.name = requiredSpreadsheetTableName(record.name, `${label} name`);
    if (record.displayName !== undefined) {
      result.displayName = requiredSpreadsheetTableName(
        record.displayName,
        `${label} display name`,
      );
    }
    if (record.ooxmlId !== undefined) {
      if (!Number.isSafeInteger(record.ooxmlId) || Number(record.ooxmlId) < 1) {
        invalidWorkOfficeSpreadsheetInput(`a positive OOXML ID for ${label}`);
      }
      result.ooxmlId = Number(record.ooxmlId);
    }
    result.range = range;
    result.columns = columns;
    result.filters = filters as WorkSpreadsheetTable['filters'];
    result.headerRow = headerRow;
    result.totalsRow = totalsRow;
    const style = requiredSpreadsheetTableStyle(record.style, label);
    const showFirstColumn = requiredBoolean(
      record.showFirstColumn,
      `${label} first-column emphasis`,
    );
    const showLastColumn = requiredBoolean(
      record.showLastColumn,
      `${label} last-column emphasis`,
    );
    const showRowStripes = requiredBoolean(
      record.showRowStripes,
      `${label} row stripes`,
    );
    const showColumnStripes = requiredBoolean(
      record.showColumnStripes,
      `${label} column stripes`,
    );
    if (!headerRow && filters.length > 0) {
      invalidWorkOfficeSpreadsheetInput(
        `${label} filters to require an enabled header row`,
      );
    }
    if (
      style.family === 'none' &&
      (showFirstColumn || showLastColumn || showRowStripes || showColumnStripes)
    ) {
      invalidWorkOfficeSpreadsheetInput(
        `${label} style flags to require a built-in style`,
      );
    }
    result.style = style;
    result.showFirstColumn = showFirstColumn;
    result.showLastColumn = showLastColumn;
    result.showRowStripes = showRowStripes;
    result.showColumnStripes = showColumnStripes;
    return result;
  });
}

function requiredSpreadsheetTableRange(
  value: unknown,
  label: string,
  sheetId: string,
): WorkSpreadsheetTable['range'] {
  const record = requiredInputRecord(value, `${label} range`);
  const readAxis = (
    candidate: unknown,
    maximum: number,
    axis: 'column' | 'row',
  ): [number, number] => {
    if (!Array.isArray(candidate) || candidate.length !== 2) {
      invalidWorkOfficeSpreadsheetInput(`a bounded ${axis} range for ${label}`);
    }
    const start = requiredCoordinate(
      candidate[0],
      maximum,
      `${axis} start`,
      sheetId,
    );
    const end = requiredCoordinate(
      candidate[1],
      maximum,
      `${axis} end`,
      sheetId,
    );
    if (start > end) {
      invalidWorkOfficeSpreadsheetInput(
        `an ordered ${axis} range for ${label}`,
      );
    }
    return [start, end];
  };
  return {
    row: readAxis(record.row, OFFICE_KERNEL_SPREADSHEET_MAX_ROWS, 'row'),
    column: readAxis(
      record.column,
      OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
      'column',
    ),
  };
}

function requiredSpreadsheetTableColumns(
  value: unknown,
  width: number,
  label: string,
): WorkSpreadsheetTable['columns'] {
  if (!Array.isArray(value) || value.length !== width) {
    invalidWorkOfficeSpreadsheetInput(
      `${label} to have one column definition per worksheet column`,
    );
  }
  const names = new Set<string>();
  return value.map((candidate) => {
    const record = requiredInputRecord(candidate, `${label} column`);
    assertOptionalRecordKeys(
      record,
      [
        'name',
        'calculatedFormula',
        'totalsFunction',
        'totalsLabel',
        'totalsFormula',
      ],
      `column for ${label}`,
    );
    const name = requiredTableColumnName(record.name, label);
    const normalized = name.toLocaleLowerCase();
    if (names.has(normalized)) {
      invalidWorkOfficeSpreadsheetInput(`unique column names for ${label}`);
    }
    names.add(normalized);
    let calculatedFormula: string | undefined;
    if (record.calculatedFormula !== undefined) {
      calculatedFormula = normalizeSpreadsheetTableCalculatedFormula(
        record.calculatedFormula,
      );
      if (
        !calculatedFormula ||
        calculatedFormula !== record.calculatedFormula
      ) {
        invalidWorkOfficeSpreadsheetInput(
          `a bounded structured calculated-column formula for ${label}`,
        );
      }
    }
    const totalsFunction =
      record.totalsFunction === undefined
        ? undefined
        : normalizeSpreadsheetTableTotalsFunction(record.totalsFunction);
    if (
      record.totalsFunction !== undefined &&
      (!totalsFunction || totalsFunction !== record.totalsFunction)
    ) {
      invalidWorkOfficeSpreadsheetInput(
        `a supported totals-row function for ${label}`,
      );
    }
    const totalsLabel =
      record.totalsLabel === undefined
        ? undefined
        : normalizeSpreadsheetTableTotalsLabel(record.totalsLabel);
    if (
      record.totalsLabel !== undefined &&
      (!totalsLabel || totalsLabel !== record.totalsLabel)
    ) {
      invalidWorkOfficeSpreadsheetInput(
        `a bounded totals-row label for ${label}`,
      );
    }
    const totalsFormula =
      record.totalsFormula === undefined
        ? undefined
        : normalizeSpreadsheetTableTotalsFormula(record.totalsFormula);
    if (
      record.totalsFormula !== undefined &&
      (!totalsFormula || totalsFormula !== record.totalsFormula)
    ) {
      invalidWorkOfficeSpreadsheetInput(
        `a bounded totals-row formula for ${label}`,
      );
    }
    if (totalsFormula && totalsFunction !== 'custom') {
      invalidWorkOfficeSpreadsheetInput(
        `custom totals-row formulas to declare the custom function for ${label}`,
      );
    }
    if (totalsFunction === 'custom' && !totalsFormula) {
      invalidWorkOfficeSpreadsheetInput(
        `custom totals-row functions to include a formula for ${label}`,
      );
    }
    if ((totalsFunction || totalsFormula) && totalsLabel) {
      invalidWorkOfficeSpreadsheetInput(
        `totals-row labels not to share a cell with a formula for ${label}`,
      );
    }
    if (
      !calculatedFormula &&
      !totalsFunction &&
      !totalsLabel &&
      !totalsFormula
    ) {
      return { name };
    }
    return {
      name,
      ...(calculatedFormula ? { calculatedFormula } : {}),
      ...(totalsFunction ? { totalsFunction } : {}),
      ...(totalsLabel ? { totalsLabel } : {}),
      ...(totalsFormula ? { totalsFormula } : {}),
    };
  });
}

function assertOptionalRecordKeys(
  record: Record<string, unknown>,
  keys: readonly string[],
  label: string,
): void {
  const allowed = new Set(keys);
  if (
    !Object.hasOwn(record, keys[0] ?? '') ||
    Object.keys(record).some((key) => !allowed.has(key))
  ) {
    invalidWorkOfficeSpreadsheetInput(
      `a complete ${label} record without unknown fields`,
    );
  }
}

function requiredSpreadsheetTableFilters(
  value: unknown,
  width: number,
  label: string,
): WorkSpreadsheetTable['filters'] {
  if (!Array.isArray(value)) {
    invalidWorkOfficeSpreadsheetInput(`an array of filters for ${label}`);
  }
  const columns = new Set<number>();
  const textBudget = { bytes: 0 };
  return value.map((candidate) => {
    const record = validateJsonRecord(
      requiredInputRecord(candidate, `${label} filter`),
      `${label} filter`,
    );
    assertExactRecordKeys(
      record,
      ['column', 'criteria'],
      `filter for ${label}`,
    );
    if (
      !Number.isSafeInteger(record.column) ||
      Number(record.column) < 0 ||
      Number(record.column) >= width ||
      columns.has(Number(record.column))
    ) {
      invalidWorkOfficeSpreadsheetInput(
        `unique in-range filter columns for ${label}`,
      );
    }
    const column = Number(record.column);
    const criteria = requiredSpreadsheetTableFilterCriteria(
      record.criteria,
      label,
      textBudget,
    );
    columns.add(column);
    return { column, criteria };
  });
}

function requiredSpreadsheetTableFilterCriteria(
  value: unknown,
  label: string,
  textBudget: { bytes: number },
): WorkSpreadsheetTableFilterCriteria {
  const record = validateJsonRecord(
    requiredInputRecord(value, `filter criteria for ${label}`),
    `filter criteria for ${label}`,
  );
  const type = record.type;
  switch (type) {
    case 'values': {
      assertExactRecordKeys(
        record,
        ['type', 'values', 'includeBlanks'],
        `filter criteria for ${label}`,
      );
      if (
        !Array.isArray(record.values) ||
        record.values.length > MAX_FILTER_VALUES_PER_COLUMN
      ) {
        invalidWorkOfficeSpreadsheetInput(
          `at most ${MAX_FILTER_VALUES_PER_COLUMN.toLocaleString()} values in filter criteria for ${label}`,
        );
      }
      const includeBlanks = requiredBoolean(
        record.includeBlanks,
        `filter blank inclusion for ${label}`,
      );
      if (record.values.length === 0 && !includeBlanks) {
        invalidWorkOfficeSpreadsheetInput(
          `at least one value or includeBlanks=true in filter criteria for ${label}`,
        );
      }
      const observed = new Set<string>();
      const values: string[] = [];
      for (const candidate of record.values) {
        const filterValue = requiredSpreadsheetFilterText(
          candidate,
          label,
          textBudget,
        );
        if (observed.has(filterValue)) {
          invalidWorkOfficeSpreadsheetInput(
            `unique filter values in criteria for ${label}`,
          );
        }
        observed.add(filterValue);
        values.push(filterValue);
      }
      return { type, values, includeBlanks };
    }
    case 'equals':
    case 'not-equals':
    case 'contains':
    case 'does-not-contain':
    case 'begins-with':
    case 'does-not-begin-with':
    case 'ends-with':
    case 'does-not-end-with':
    case 'greater-than':
    case 'greater-than-or-equal':
    case 'less-than':
    case 'less-than-or-equal':
      return requiredSpreadsheetCustomFilterCondition(
        record,
        label,
        textBudget,
      );
    case 'between':
    case 'not-between':
      assertExactRecordKeys(
        record,
        ['type', 'lower', 'upper'],
        `filter criteria for ${label}`,
      );
      return {
        type,
        lower: requiredSpreadsheetFilterText(record.lower, label, textBudget),
        upper: requiredSpreadsheetFilterText(record.upper, label, textBudget),
      };
    case 'blanks':
    case 'non-blanks':
      assertExactRecordKeys(record, ['type'], `filter criteria for ${label}`);
      return { type };
    case 'compound': {
      assertExactRecordKeys(
        record,
        ['type', 'conjunction', 'conditions'],
        `filter criteria for ${label}`,
      );
      if (record.conjunction !== 'and' && record.conjunction !== 'or') {
        invalidWorkOfficeSpreadsheetInput(
          `an 'and' or 'or' conjunction for compound filter criteria for ${label}`,
        );
      }
      if (!Array.isArray(record.conditions) || record.conditions.length !== 2) {
        invalidWorkOfficeSpreadsheetInput(
          `exactly two custom filter conditions for ${label}`,
        );
      }
      return {
        type,
        conjunction: record.conjunction,
        conditions: [
          requiredSpreadsheetCustomFilterCondition(
            record.conditions[0],
            label,
            textBudget,
          ),
          requiredSpreadsheetCustomFilterCondition(
            record.conditions[1],
            label,
            textBudget,
          ),
        ],
      };
    }
    case 'top':
    case 'bottom':
      assertExactRecordKeys(
        record,
        ['type', 'count'],
        `filter criteria for ${label}`,
      );
      return {
        type,
        count: requiredSpreadsheetFilterInteger(
          record.count,
          1,
          500,
          `a ${type} filter count from 1 through 500 for ${label}`,
        ),
      };
    case 'top-percent':
    case 'bottom-percent':
      assertExactRecordKeys(
        record,
        ['type', 'percent'],
        `filter criteria for ${label}`,
      );
      return {
        type,
        percent: requiredSpreadsheetFilterInteger(
          record.percent,
          1,
          100,
          `a ${type} filter percentage from 1 through 100 for ${label}`,
        ),
      };
    case 'dynamic':
      assertExactRecordKeys(
        record,
        ['type', 'kind'],
        `filter criteria for ${label}`,
      );
      if (
        typeof record.kind !== 'string' ||
        !SPREADSHEET_DYNAMIC_FILTERS.has(record.kind)
      ) {
        invalidWorkOfficeSpreadsheetInput(
          `a supported dynamic filter for ${label}`,
        );
      }
      return { type, kind: record.kind as WorkSpreadsheetDynamicFilter };
    default:
      invalidWorkOfficeSpreadsheetInput(
        `supported filter criteria for ${label}`,
      );
  }
}

function requiredSpreadsheetCustomFilterCondition(
  value: unknown,
  label: string,
  textBudget: { bytes: number },
): WorkSpreadsheetCustomFilterCondition {
  const record = validateJsonRecord(
    requiredInputRecord(value, `custom filter condition for ${label}`),
    `custom filter condition for ${label}`,
  );
  switch (record.type) {
    case 'equals':
    case 'not-equals':
    case 'contains':
    case 'does-not-contain':
    case 'begins-with':
    case 'does-not-begin-with':
    case 'ends-with':
    case 'does-not-end-with':
    case 'greater-than':
    case 'greater-than-or-equal':
    case 'less-than':
    case 'less-than-or-equal':
      assertExactRecordKeys(
        record,
        ['type', 'value'],
        `custom filter condition for ${label}`,
      );
      return {
        type: record.type,
        value: requiredSpreadsheetFilterText(record.value, label, textBudget),
      };
    default:
      invalidWorkOfficeSpreadsheetInput(
        `supported custom filter condition for ${label}`,
      );
  }
}

function requiredSpreadsheetFilterText(
  value: unknown,
  label: string,
  textBudget: { bytes: number },
): string {
  if (typeof value !== 'string') {
    invalidWorkOfficeSpreadsheetInput(
      `filter text containing 1 to ${MAX_FILTER_VALUE_CHARACTERS.toLocaleString()} characters for ${label}`,
    );
  }
  let characters = 0;
  for (const character of value) {
    characters += 1;
    const codePoint = character.codePointAt(0);
    if (
      codePoint === undefined ||
      !(
        codePoint === 0x9 ||
        codePoint === 0xa ||
        codePoint === 0xd ||
        (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
        (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
        (codePoint >= 0x10000 && codePoint <= 0x10ffff)
      )
    ) {
      invalidWorkOfficeSpreadsheetInput(
        `XML-compatible filter text for ${label}`,
      );
    }
  }
  if (characters < 1 || characters > MAX_FILTER_VALUE_CHARACTERS) {
    invalidWorkOfficeSpreadsheetInput(
      `filter text containing 1 to ${MAX_FILTER_VALUE_CHARACTERS.toLocaleString()} characters for ${label}`,
    );
  }
  textBudget.bytes += FILTER_TEXT_ENCODER.encode(value).byteLength;
  if (textBudget.bytes > MAX_FILTER_TEXT_BYTES) {
    invalidWorkOfficeSpreadsheetInput(
      `filter text totaling at most ${MAX_FILTER_TEXT_BYTES.toLocaleString()} bytes for ${label}`,
    );
  }
  return value;
}

function requiredSpreadsheetFilterInteger(
  value: unknown,
  minimum: number,
  maximum: number,
  expected: string,
): number {
  if (
    !Number.isSafeInteger(value) ||
    Number(value) < minimum ||
    Number(value) > maximum
  ) {
    invalidWorkOfficeSpreadsheetInput(expected);
  }
  return Number(value);
}

function assertExactRecordKeys(
  record: Record<string, unknown>,
  keys: readonly string[],
  label: string,
): void {
  const allowed = new Set(keys);
  if (
    keys.some((key) => !Object.hasOwn(record, key)) ||
    Object.keys(record).some((key) => !allowed.has(key))
  ) {
    invalidWorkOfficeSpreadsheetInput(
      `a complete ${label} record without unknown fields`,
    );
  }
}

function requiredSpreadsheetTableStyle(
  value: unknown,
  label: string,
): WorkSpreadsheetTable['style'] {
  const record = validateJsonRecord(
    requiredInputRecord(value, `${label} style`),
    `${label} style`,
  );
  const family = record.family;
  if (family === 'none') {
    assertExactRecordKeys(record, ['family'], `style for ${label}`);
    return { family };
  }
  assertExactRecordKeys(record, ['family', 'number'], `style for ${label}`);
  const maximum = family === 'light' ? 21 : family === 'medium' ? 28 : 11;
  if (
    (family !== 'light' && family !== 'medium' && family !== 'dark') ||
    !Number.isSafeInteger(record.number) ||
    Number(record.number) < 1 ||
    Number(record.number) > maximum
  ) {
    invalidWorkOfficeSpreadsheetInput(`a built-in table style for ${label}`);
  }
  return { family, number: Number(record.number) };
}

function requiredSpreadsheetTableName(value: unknown, label: string): string {
  if (
    typeof value !== 'string' ||
    value.trim() !== value ||
    Array.from(value).length < 1 ||
    Array.from(value).length > 255
  ) {
    invalidWorkOfficeSpreadsheetInput(`a valid table name for ${label}`);
  }
  const result = value as string;
  if (
    !isValidSpreadsheetDefinedName(result) ||
    ['r', 'c'].includes(result.toLocaleLowerCase())
  ) {
    invalidWorkOfficeSpreadsheetInput(`a valid table name for ${label}`);
  }
  return result;
}

function requiredTableColumnName(value: unknown, label: string): string {
  const characters = typeof value === 'string' ? Array.from(value) : [];
  if (
    typeof value !== 'string' ||
    characters.length < 1 ||
    characters.length > 255 ||
    value.trim() !== value ||
    characters.some(
      (character) =>
        /\p{Cc}/u.test(character) ||
        character === '\uFFFE' ||
        character === '\uFFFF',
    )
  ) {
    invalidWorkOfficeSpreadsheetInput(`a valid column name for ${label}`);
  }
  return value as string;
}

function requiredBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') {
    invalidWorkOfficeSpreadsheetInput(`a Boolean value for ${label}`);
  }
  return value as boolean;
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
  const definedNames = new Set(
    (content.namedRanges ?? []).map((range) =>
      range.name.trim().toLocaleLowerCase(),
    ),
  );
  const tableNames = new Set<string>();
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
    for (const [index, table] of (sheet.tables ?? []).entries()) {
      const names = [table.name, table.displayName].filter(
        (value): value is string => Boolean(value),
      );
      for (const name of names) {
        const normalized = name.toLocaleLowerCase();
        if (definedNames.has(normalized) || tableNames.has(normalized)) {
          invalidWorkOfficeSpreadsheetInput(
            `unique table and defined names; '${name}' is repeated`,
          );
        }
        tableNames.add(normalized);
      }
      if (
        table.range.row[1] >=
          (sheet.row ?? OFFICE_KERNEL_SPREADSHEET_MAX_ROWS) ||
        table.range.column[1] >=
          (sheet.column ?? OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS)
      ) {
        invalidWorkOfficeSpreadsheetInput(
          `table '${table.id}' to fit within sheet '${sheet.id}'`,
        );
      }
      for (const other of (sheet.tables ?? []).slice(0, index)) {
        if (spreadsheetTableRangesIntersect(table.range, other.range)) {
          invalidWorkOfficeSpreadsheetInput(
            `non-overlapping tables in sheet '${sheet.id}'`,
          );
        }
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

function spreadsheetTableRangesIntersect(
  left: WorkSpreadsheetTable['range'],
  right: WorkSpreadsheetTable['range'],
): boolean {
  return (
    left.row[0] <= right.row[1] &&
    left.row[1] >= right.row[0] &&
    left.column[0] <= right.column[1] &&
    left.column[1] >= right.column[0]
  );
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

function assertPopulatedCellLimit(count: number, sheetId: string): void {
  if (count <= MAX_POPULATED_CELLS) return;
  invalidWorkOfficeSpreadsheetInput(
    `at most ${MAX_POPULATED_CELLS.toLocaleString()} populated cells in sheet '${sheetId}'`,
  );
}
