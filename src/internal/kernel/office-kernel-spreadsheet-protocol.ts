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

/**
 * The bounded table metadata needed to resolve native structured references.
 * Ranges are zero-based and inclusive, matching the controlled workbook
 * model. Styling, filters, and drawing metadata intentionally stay outside
 * the calculation kernel.
 */
export interface OfficeKernelSpreadsheetInputTable {
  name: string;
  displayName?: string;
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
  columns: string[];
  headerRow: boolean;
  totalsRow: boolean;
}

export interface OfficeKernelSpreadsheetInputSheet {
  id: string;
  name: string;
  cells: OfficeKernelSpreadsheetInputCell[];
  tables?: OfficeKernelSpreadsheetInputTable[];
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

export type OfficeKernelSpreadsheetSessionCellChange =
  | (OfficeKernelSpreadsheetCoordinate & {
      kind: 'upsert';
      formula?: string;
      value: OfficeKernelSpreadsheetValue;
    })
  | (OfficeKernelSpreadsheetCoordinate & {
      kind: 'remove';
    });

export type OfficeKernelSpreadsheetSessionUpdate =
  | {
      kind: 'replace';
      sheets: OfficeKernelSpreadsheetInputSheet[];
    }
  | {
      kind: 'patch';
      baseDocumentRevision: number;
      changes: OfficeKernelSpreadsheetSessionCellChange[];
    };

export type OfficeKernelSpreadsheetSessionCalculationScope =
  | { kind: 'workbook' }
  | { kind: 'dirty' }
  | {
      kind: 'targets';
      targets: OfficeKernelSpreadsheetCoordinate[];
    };

export interface OfficeKernelSpreadsheetSessionCalculationRequest {
  protocol: typeof OFFICE_KERNEL_PROTOCOL_VERSION;
  kind: 'spreadsheetSessionCalculation';
  requestId: number;
  revision: number;
  documentRevision: number;
  update: OfficeKernelSpreadsheetSessionUpdate;
  calculation: OfficeKernelSpreadsheetSessionCalculationScope;
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

export interface OfficeKernelSpreadsheetSessionCalculationStats {
  updateKind: 'replace' | 'patch';
  calculationScope: 'workbook' | 'dirty' | 'targets';
  formulaCellCount: number;
  dirtyFormulaCellCount: number;
  evaluatedFormulaCellCount: number;
  reusedFormulaCellCount: number;
  dependencyEdgeCount: number;
}

export interface OfficeKernelSpreadsheetSessionCalculationResult
  extends Omit<OfficeKernelSpreadsheetCalculationResult, 'kind'> {
  kind: 'spreadsheetSessionCalculationResult';
  stats: OfficeKernelSpreadsheetSessionCalculationStats;
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
  return isOfficeKernelSpreadsheetCalculationResultShape(value);
}

export function isOfficeKernelSpreadsheetSessionCalculationResult(
  value: unknown,
): value is OfficeKernelSpreadsheetSessionCalculationResult {
  if (!isOfficeKernelSpreadsheetCalculationResultShape(value)) return false;
  const candidate = value as Record<string, unknown>;
  return isSessionCalculationStats(candidate.stats);
}

function isOfficeKernelSpreadsheetCalculationResultShape(
  value: unknown,
): boolean {
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

function isSessionCalculationStats(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  const stats = value as Record<string, unknown>;
  return (
    (stats.updateKind === 'replace' || stats.updateKind === 'patch') &&
    (stats.calculationScope === 'workbook' ||
      stats.calculationScope === 'dirty' ||
      stats.calculationScope === 'targets') &&
    isBoundedInteger(stats.formulaCellCount, 100_000) &&
    isBoundedInteger(stats.dirtyFormulaCellCount, 100_000) &&
    isBoundedInteger(stats.evaluatedFormulaCellCount, 100_000) &&
    isBoundedInteger(stats.reusedFormulaCellCount, 100_000) &&
    isBoundedInteger(stats.dependencyEdgeCount, 1_000_000)
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
