import type {
  OfficeKernelSpreadsheetCalculationRequest,
  OfficeKernelSpreadsheetInputSheet,
  OfficeKernelSpreadsheetValue,
} from './office-kernel-spreadsheet-protocol';
import {
  isOfficeKernelSpreadsheetError,
  OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
  OFFICE_KERNEL_SPREADSHEET_MAX_ROWS,
  OFFICE_KERNEL_SPREADSHEET_MAX_TEXT_BYTES,
} from './office-kernel-spreadsheet-protocol';
import { OFFICE_KERNEL_PROTOCOL_VERSION } from './office-kernel-version';

const MAX_SPREADSHEET_INPUT_CELLS = 100_000;
const MAX_SPREADSHEET_FORMULA_CHARACTERS = 8_192;
const MAX_SPREADSHEET_SHEETS = 1_024;
const MAX_SPREADSHEET_IDENTIFIER_BYTES = 256;
const MAX_UNSIGNED_32_BIT_INTEGER = 4_294_967_295;

export const MAX_SPREADSHEET_DEPENDENCY_DEPTH = 64;
export const MAX_SPREADSHEET_RANGE_CELLS = 100_000;

export function validateSpreadsheetCalculationRequest(
  request: OfficeKernelSpreadsheetCalculationRequest,
): void {
  if (request.protocol !== OFFICE_KERNEL_PROTOCOL_VERSION) {
    throw kernelError(
      'office.kernel.protocol_unsupported',
      `Office kernel protocol ${request.protocol} is unsupported.`,
    );
  }
  if (request.kind !== 'spreadsheetCalculation') {
    throw kernelError(
      'office.kernel.request_kind_invalid',
      'The Spreadsheet kernel only accepts calculation requests.',
    );
  }
  if (
    !unsigned32BitInteger(request.requestId) ||
    !unsigned32BitInteger(request.revision) ||
    !nonNegativeSafeInteger(request.documentRevision)
  ) {
    throw kernelError(
      'office.kernel.revision_invalid',
      'Request, calculation, and document revisions must be non-negative safe integers.',
    );
  }
  if (request.sheets.length > MAX_SPREADSHEET_SHEETS) {
    throw kernelError(
      'office.kernel.spreadsheet.sheet_limit_exceeded',
      `A Spreadsheet calculation request may contain at most ${MAX_SPREADSHEET_SHEETS} sheets.`,
    );
  }
  const totalCells = request.sheets.reduce(
    (count, sheet) => count + sheet.cells.length,
    0,
  );
  if (totalCells > MAX_SPREADSHEET_INPUT_CELLS) {
    throw kernelError(
      'office.kernel.spreadsheet.cell_limit_exceeded',
      `A Spreadsheet calculation request may contain at most ${MAX_SPREADSHEET_INPUT_CELLS} populated cells.`,
    );
  }
  if (request.targets && request.targets.length > MAX_SPREADSHEET_INPUT_CELLS) {
    throw kernelError(
      'office.kernel.spreadsheet.target_limit_exceeded',
      `A Spreadsheet calculation request may contain at most ${MAX_SPREADSHEET_INPUT_CELLS} targets.`,
    );
  }

  const sheetIds = new Set<string>();
  const sheetNames = new Set<string>();
  for (const sheet of request.sheets) {
    const normalizedName = sheet.name.toLowerCase();
    if (
      !sheet.id.trim() ||
      utf8ByteLength(sheet.id) > MAX_SPREADSHEET_IDENTIFIER_BYTES ||
      !sheet.name.trim() ||
      utf8ByteLength(sheet.name) > MAX_SPREADSHEET_IDENTIFIER_BYTES ||
      sheetIds.has(sheet.id) ||
      sheetNames.has(normalizedName)
    ) {
      throw kernelError(
        'office.kernel.spreadsheet.sheet_invalid',
        'Spreadsheet sheet IDs and names must be unique and non-empty.',
      );
    }
    sheetIds.add(sheet.id);
    sheetNames.add(normalizedName);

    validateSpreadsheetTables(sheet, request.sheets);

    const coordinates = new Set<string>();
    for (const cell of sheet.cells) {
      const key = `${cell.row}:${cell.column}`;
      if (
        !boundedSpreadsheetIndex(
          cell.row,
          OFFICE_KERNEL_SPREADSHEET_MAX_ROWS,
        ) ||
        !boundedSpreadsheetIndex(
          cell.column,
          OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
        ) ||
        coordinates.has(key)
      ) {
        throw kernelError(
          'office.kernel.spreadsheet.cell_invalid',
          'Spreadsheet cells require unique, bounded row and column coordinates.',
        );
      }
      coordinates.add(key);
      validateSpreadsheetValue(cell.value);
      if (
        cell.formula !== undefined &&
        (cell.formula.length === 0 ||
          Array.from(cell.formula).length > MAX_SPREADSHEET_FORMULA_CHARACTERS)
      ) {
        throw kernelError(
          'office.kernel.spreadsheet.formula_invalid',
          `Spreadsheet formulas must contain 1-${MAX_SPREADSHEET_FORMULA_CHARACTERS} characters.`,
        );
      }
    }
  }

  for (const target of request.targets ?? []) {
    if (
      !sheetIds.has(target.sheetId) ||
      !boundedSpreadsheetIndex(
        target.row,
        OFFICE_KERNEL_SPREADSHEET_MAX_ROWS,
      ) ||
      !boundedSpreadsheetIndex(
        target.column,
        OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
      )
    ) {
      throw kernelError(
        'office.kernel.spreadsheet.target_invalid',
        'Spreadsheet calculation targets must reference an existing, bounded cell.',
      );
    }
  }
}

function validateSpreadsheetTables(
  sheet: OfficeKernelSpreadsheetInputSheet,
  sheets: readonly OfficeKernelSpreadsheetInputSheet[],
): void {
  const tables = sheet.tables ?? [];
  const tableCount = sheets.reduce(
    (count, candidate) => count + (candidate.tables?.length ?? 0),
    0,
  );
  if (tableCount > 1_024) {
    throw kernelError(
      'office.kernel.spreadsheet.table_limit_exceeded',
      'A Spreadsheet calculation request may contain at most 1024 tables.',
    );
  }
  const aliases = new Map<string, string>();
  for (const candidate of sheets) {
    for (const [tableIndex, table] of (candidate.tables ?? []).entries()) {
      const identity = `${candidate.id}\u0000${tableIndex}`;
      for (const alias of [table.name, table.displayName]) {
        if (!alias) continue;
        const normalized = alias.toLocaleLowerCase();
        const existing = aliases.get(normalized);
        if (existing && existing !== identity) {
          throw kernelError(
            'office.kernel.spreadsheet.table_invalid',
            `Spreadsheet table name '${alias}' is ambiguous.`,
          );
        }
        aliases.set(normalized, identity);
      }
    }
  }
  const ranges: Array<{
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
    name: string;
  }> = [];
  for (const table of tables) {
    for (const [kind, value] of [
      ['startRow', table.startRow],
      ['endRow', table.endRow],
      ['startColumn', table.startColumn],
      ['endColumn', table.endColumn],
    ] as const) {
      if (
        !boundedSpreadsheetIndex(
          value,
          kind.endsWith('Row')
            ? OFFICE_KERNEL_SPREADSHEET_MAX_ROWS
            : OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
        )
      ) {
        throw kernelError(
          'office.kernel.spreadsheet.table_invalid',
          `Spreadsheet table '${table.name}' has an out-of-bounds ${kind}.`,
        );
      }
    }
    if (
      table.startRow > table.endRow ||
      table.startColumn > table.endColumn ||
      table.columns.length !== table.endColumn - table.startColumn + 1
    ) {
      throw kernelError(
        'office.kernel.spreadsheet.table_invalid',
        `Spreadsheet table '${table.name}' has an invalid range or column count.`,
      );
    }
    validateTableName(table.name, 'name');
    if (table.displayName !== undefined)
      validateTableName(table.displayName, 'displayName');
    const columnNames = new Set<string>();
    for (const column of table.columns) {
      validateTableName(column, 'column');
      const normalized = column.toLocaleLowerCase();
      if (columnNames.has(normalized)) {
        throw kernelError(
          'office.kernel.spreadsheet.table_invalid',
          `Spreadsheet table '${table.name}' contains duplicate column names.`,
        );
      }
      columnNames.add(normalized);
    }
    if (table.headerRow && table.totalsRow && table.startRow === table.endRow) {
      throw kernelError(
        'office.kernel.spreadsheet.table_invalid',
        `Spreadsheet table '${table.name}' cannot use header and totals rows in one row.`,
      );
    }
    for (const previous of ranges) {
      if (
        table.startRow <= previous.endRow &&
        table.endRow >= previous.startRow &&
        table.startColumn <= previous.endColumn &&
        table.endColumn >= previous.startColumn
      ) {
        throw kernelError(
          'office.kernel.spreadsheet.table_invalid',
          `Spreadsheet tables '${previous.name}' and '${table.name}' overlap.`,
        );
      }
    }
    ranges.push({
      startRow: table.startRow,
      endRow: table.endRow,
      startColumn: table.startColumn,
      endColumn: table.endColumn,
      name: table.name,
    });
  }
}

function validateTableName(value: string, kind: string): void {
  if (
    !value.trim() ||
    utf8ByteLength(value) > MAX_SPREADSHEET_IDENTIFIER_BYTES
  ) {
    throw kernelError(
      'office.kernel.spreadsheet.table_invalid',
      `Spreadsheet table ${kind} must contain 1-${MAX_SPREADSHEET_IDENTIFIER_BYTES} UTF-8 bytes.`,
    );
  }
}

export function boundedSpreadsheetIndex(
  value: number,
  exclusiveMaximum: number,
): boolean {
  return Number.isSafeInteger(value) && value >= 0 && value < exclusiveMaximum;
}

function validateSpreadsheetValue(value: OfficeKernelSpreadsheetValue): void {
  if (value.kind === 'number' && !Number.isFinite(value.value)) {
    throw kernelError(
      'office.kernel.spreadsheet.value_invalid',
      'Spreadsheet numeric values must be finite.',
    );
  }
  if (
    value.kind === 'text' &&
    utf8ByteLength(value.value) > OFFICE_KERNEL_SPREADSHEET_MAX_TEXT_BYTES
  ) {
    throw kernelError(
      'office.kernel.spreadsheet.value_invalid',
      `Spreadsheet text values may contain at most ${OFFICE_KERNEL_SPREADSHEET_MAX_TEXT_BYTES} UTF-8 bytes.`,
    );
  }
  if (value.kind === 'error' && !isOfficeKernelSpreadsheetError(value.value)) {
    throw kernelError(
      'office.kernel.spreadsheet.value_invalid',
      `Spreadsheet error value '${value.value}' is not recognized.`,
    );
  }
}

function nonNegativeSafeInteger(value: number): boolean {
  return Number.isSafeInteger(value) && value >= 0;
}

function unsigned32BitInteger(value: number): boolean {
  return nonNegativeSafeInteger(value) && value <= MAX_UNSIGNED_32_BIT_INTEGER;
}

function utf8ByteLength(value: string): number {
  return new TextEncoder().encode(value).byteLength;
}

function kernelError(code: string, message: string): Error {
  return Object.assign(new Error(message), { code });
}
