import { calculateSpreadsheetInJavaScript } from './office-kernel-spreadsheet-fallback';
import type {
  OfficeKernelSpreadsheetCalculationRequest,
  OfficeKernelSpreadsheetInputSheet,
  OfficeKernelSpreadsheetSessionCalculationRequest,
  OfficeKernelSpreadsheetSessionCalculationResult,
} from './office-kernel-spreadsheet-protocol';
import {
  OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
  OFFICE_KERNEL_SPREADSHEET_MAX_ROWS,
} from './office-kernel-spreadsheet-protocol';

const MAX_SPREADSHEET_SESSION_CHANGES = 100_000;

export class JavaScriptSpreadsheetCalculationSession {
  private documentRevision: number | null = null;
  private sheets: OfficeKernelSpreadsheetInputSheet[] | null = null;

  async calculate(
    request: OfficeKernelSpreadsheetSessionCalculationRequest,
  ): Promise<OfficeKernelSpreadsheetSessionCalculationResult> {
    const nextSheets =
      request.update.kind === 'replace'
        ? cloneSheets(request.update.sheets)
        : this.patchSheets(request);
    const result = await calculateSpreadsheetSessionInJavaScript(
      request,
      nextSheets,
    );
    this.sheets = nextSheets;
    this.documentRevision = request.documentRevision;
    return result;
  }

  private patchSheets(
    request: OfficeKernelSpreadsheetSessionCalculationRequest,
  ): OfficeKernelSpreadsheetInputSheet[] {
    if (request.update.kind !== 'patch') {
      throw new Error('Spreadsheet session patch update is required.');
    }
    if (!this.sheets || this.documentRevision === null) {
      throw sessionError(
        'office.kernel.spreadsheet.session_uninitialized',
        'A Spreadsheet calculation session requires a replace update before patches.',
      );
    }
    if (request.update.baseDocumentRevision !== this.documentRevision) {
      throw sessionError(
        'office.kernel.spreadsheet.session_revision_mismatch',
        `Spreadsheet patch revision ${request.update.baseDocumentRevision} does not match session revision ${this.documentRevision}.`,
      );
    }
    if (
      request.documentRevision < request.update.baseDocumentRevision ||
      (request.update.changes.length > 0 &&
        request.documentRevision === request.update.baseDocumentRevision)
    ) {
      throw sessionError(
        'office.kernel.spreadsheet.session_revision_invalid',
        'A Spreadsheet patch cannot move a session backwards, and a non-empty patch requires a newer document revision.',
      );
    }
    if (request.update.changes.length > MAX_SPREADSHEET_SESSION_CHANGES) {
      throw sessionError(
        'office.kernel.spreadsheet.patch_limit_exceeded',
        `A Spreadsheet session patch may contain at most ${MAX_SPREADSHEET_SESSION_CHANGES} cell changes.`,
      );
    }
    const nextSheets = cloneSheets(this.sheets);
    const sheetsById = new Map(nextSheets.map((sheet) => [sheet.id, sheet]));
    const coordinates = new Set<string>();
    for (const change of request.update.changes) {
      if (
        !boundedIndex(change.row, OFFICE_KERNEL_SPREADSHEET_MAX_ROWS) ||
        !boundedIndex(change.column, OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS)
      ) {
        throw sessionError(
          'office.kernel.spreadsheet.cell_invalid',
          'Spreadsheet cell coordinates must remain within XFD1048576.',
        );
      }
      const sheet = sheetsById.get(change.sheetId);
      if (!sheet) {
        throw sessionError(
          'office.kernel.spreadsheet.patch_invalid',
          `Spreadsheet patch references missing sheet '${change.sheetId}'.`,
        );
      }
      const key = `${change.sheetId}\u0000${change.row}\u0000${change.column}`;
      if (coordinates.has(key)) {
        throw sessionError(
          'office.kernel.spreadsheet.patch_invalid',
          'Spreadsheet session patches require unique cell coordinates.',
        );
      }
      coordinates.add(key);
      const index = sheet.cells.findIndex(
        (cell) => cell.row === change.row && cell.column === change.column,
      );
      if (change.kind === 'remove') {
        if (index >= 0) sheet.cells.splice(index, 1);
        continue;
      }
      const cell = {
        row: change.row,
        column: change.column,
        formula: change.formula,
        value: change.value,
      };
      if (index >= 0) {
        sheet.cells[index] = cell;
      } else {
        sheet.cells.push(cell);
      }
    }
    for (const sheet of nextSheets) {
      sheet.cells.sort(
        (left, right) => left.row - right.row || left.column - right.column,
      );
    }
    return nextSheets;
  }
}

export async function calculateSpreadsheetSessionInJavaScript(
  request: OfficeKernelSpreadsheetSessionCalculationRequest,
  sheets: OfficeKernelSpreadsheetInputSheet[],
): Promise<OfficeKernelSpreadsheetSessionCalculationResult> {
  const calculationRequest: OfficeKernelSpreadsheetCalculationRequest = {
    protocol: request.protocol,
    kind: 'spreadsheetCalculation',
    requestId: request.requestId,
    revision: request.revision,
    documentRevision: request.documentRevision,
    sheets,
    targets:
      request.calculation.kind === 'targets'
        ? request.calculation.targets
        : undefined,
  };
  const result = await calculateSpreadsheetInJavaScript(calculationRequest);
  const formulaKeys = new Set(
    sheets.flatMap((sheet) =>
      sheet.cells.flatMap((cell) =>
        cell.formula
          ? [`${sheet.id}\u0000${cell.row}\u0000${cell.column}`]
          : [],
      ),
    ),
  );
  const dirtyFormulaCellCount =
    request.calculation.kind === 'targets'
      ? new Set(
          request.calculation.targets.flatMap((target) => {
            const key = `${target.sheetId}\u0000${target.row}\u0000${target.column}`;
            return formulaKeys.has(key) ? [key] : [];
          }),
        ).size
      : formulaKeys.size;
  return {
    ...result,
    kind: 'spreadsheetSessionCalculationResult',
    stats: {
      updateKind: request.update.kind,
      calculationScope: request.calculation.kind,
      formulaCellCount: formulaKeys.size,
      dirtyFormulaCellCount,
      evaluatedFormulaCellCount: result.calculationOrder.length,
      reusedFormulaCellCount: 0,
      dependencyEdgeCount: 0,
    },
  };
}

function cloneSheets(
  sheets: readonly OfficeKernelSpreadsheetInputSheet[],
): OfficeKernelSpreadsheetInputSheet[] {
  return sheets.map((sheet) => ({
    id: sheet.id,
    name: sheet.name,
    ...(sheet.tables?.length
      ? {
          tables: sheet.tables.map((table) => ({
            ...table,
            columns: [...table.columns],
          })),
        }
      : {}),
    cells: sheet.cells.map((cell) => ({
      ...cell,
      value: { ...cell.value },
    })),
  }));
}

function sessionError(code: string, message: string): Error {
  return Object.assign(new Error(message), { code });
}

function boundedIndex(value: number, exclusiveMaximum: number): boolean {
  return Number.isSafeInteger(value) && value >= 0 && value < exclusiveMaximum;
}
