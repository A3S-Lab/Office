import type { OFFICE_KERNEL_PROTOCOL_VERSION } from './office-kernel-version';

export const OFFICE_KERNEL_SPREADSHEET_MAX_ROWS = 1_048_576;
export const OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS = 16_384;
export const OFFICE_KERNEL_SPREADSHEET_MAX_TEXT_BYTES = 32_767;

export type OfficeKernelSpreadsheetValue =
  | { kind: 'blank' }
  | { kind: 'number'; value: number }
  | { kind: 'text'; value: string }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'error'; value: OfficeKernelSpreadsheetError };

export type OfficeKernelSpreadsheetError =
  | '#BLOCKED!'
  | '#BUSY!'
  | '#CALC!'
  | '#CONNECT!'
  | '#DIV/0!'
  | '#FIELD!'
  | '#GETTING_DATA'
  | '#N/A'
  | '#NAME?'
  | '#NULL!'
  | '#NUM!'
  | '#PYTHON!'
  | '#REF!'
  | '#SPILL!'
  | '#UNKNOWN!'
  | '#VALUE!';

export interface OfficeKernelSpreadsheetCoordinate {
  sheetId: string;
  row: number;
  column: number;
}

export interface OfficeKernelSpreadsheetInputCell {
  row: number;
  column: number;
  formula?: string;
  value: OfficeKernelSpreadsheetValue;
}

export interface OfficeKernelSpreadsheetInputSheet {
  id: string;
  name: string;
  cells: OfficeKernelSpreadsheetInputCell[];
}

export interface OfficeKernelSpreadsheetCalculationRequest {
  protocol: typeof OFFICE_KERNEL_PROTOCOL_VERSION;
  kind: 'spreadsheetCalculation';
  requestId: number;
  revision: number;
  documentRevision: number;
  sheets: OfficeKernelSpreadsheetInputSheet[];
  targets?: OfficeKernelSpreadsheetCoordinate[];
}

export interface OfficeKernelSpreadsheetCalculatedCell
  extends OfficeKernelSpreadsheetCoordinate {
  value: OfficeKernelSpreadsheetValue;
}

export interface OfficeKernelSpreadsheetCalculationIssue {
  cell: OfficeKernelSpreadsheetCoordinate;
  code: string;
  message: string;
}

export interface OfficeKernelSpreadsheetCalculationResult {
  protocol: typeof OFFICE_KERNEL_PROTOCOL_VERSION;
  kind: 'spreadsheetCalculationResult';
  requestId: number;
  revision: number;
  documentRevision: number;
  engine: 'wasm' | 'javascript';
  cells: OfficeKernelSpreadsheetCalculatedCell[];
  calculationOrder: OfficeKernelSpreadsheetCoordinate[];
  issues: OfficeKernelSpreadsheetCalculationIssue[];
}

const spreadsheetErrors = new Set<OfficeKernelSpreadsheetError>([
  '#BLOCKED!',
  '#BUSY!',
  '#CALC!',
  '#CONNECT!',
  '#DIV/0!',
  '#FIELD!',
  '#GETTING_DATA',
  '#N/A',
  '#NAME?',
  '#NULL!',
  '#NUM!',
  '#PYTHON!',
  '#REF!',
  '#SPILL!',
  '#UNKNOWN!',
  '#VALUE!',
]);
const spreadsheetTextEncoder = new TextEncoder();

export function isOfficeKernelSpreadsheetError(
  value: unknown,
): value is OfficeKernelSpreadsheetError {
  return (
    typeof value === 'string' &&
    spreadsheetErrors.has(value as OfficeKernelSpreadsheetError)
  );
}

export function isOfficeKernelSpreadsheetCalculationResult(
  value: unknown,
): value is OfficeKernelSpreadsheetCalculationResult {
  if (!value || typeof value !== 'object') return false;
  const candidate = value as Record<string, unknown>;
  return (
    Array.isArray(candidate.cells) &&
    candidate.cells.every(isCalculatedCell) &&
    Array.isArray(candidate.calculationOrder) &&
    candidate.calculationOrder.every(isCoordinate) &&
    Array.isArray(candidate.issues) &&
    candidate.issues.every(isCalculationIssue)
  );
}

function isCalculatedCell(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  const cell = value as Record<string, unknown>;
  return isCoordinate(cell) && isSpreadsheetValue(cell.value);
}

function isCalculationIssue(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  const issue = value as Record<string, unknown>;
  return (
    isCoordinate(issue.cell) &&
    typeof issue.code === 'string' &&
    issue.code.length > 0 &&
    typeof issue.message === 'string' &&
    issue.message.length > 0
  );
}

function isCoordinate(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  const coordinate = value as Record<string, unknown>;
  return (
    typeof coordinate.sheetId === 'string' &&
    coordinate.sheetId.length > 0 &&
    coordinate.sheetId.length <= 256 &&
    isBoundedInteger(coordinate.row, OFFICE_KERNEL_SPREADSHEET_MAX_ROWS - 1) &&
    isBoundedInteger(
      coordinate.column,
      OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS - 1,
    )
  );
}

function isSpreadsheetValue(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  const scalar = value as Record<string, unknown>;
  switch (scalar.kind) {
    case 'blank':
      return true;
    case 'number':
      return typeof scalar.value === 'number' && Number.isFinite(scalar.value);
    case 'text':
      return (
        typeof scalar.value === 'string' &&
        spreadsheetTextEncoder.encode(scalar.value).byteLength <=
          OFFICE_KERNEL_SPREADSHEET_MAX_TEXT_BYTES
      );
    case 'boolean':
      return typeof scalar.value === 'boolean';
    case 'error':
      return isOfficeKernelSpreadsheetError(scalar.value);
    default:
      return false;
  }
}

function isBoundedInteger(value: unknown, maximum: number): boolean {
  return (
    typeof value === 'number' &&
    Number.isSafeInteger(value) &&
    value >= 0 &&
    value <= maximum
  );
}
