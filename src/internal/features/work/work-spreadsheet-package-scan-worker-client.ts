import type { Cell, CellMatrix } from '@fortune-sheet/core';
import type { WorkBook } from 'xlsx';
import {
  recordSpreadsheetImportMeasure,
  spreadsheetImportNow,
} from './work-spreadsheet-import-diagnostics';
import { freezeImportedSpreadsheetCell } from './work-spreadsheet-matrix-profile';
import type {
  SpreadsheetPackageScanResult,
  SpreadsheetPlainWorksheet,
  SpreadsheetPackageScanWorkerResponse,
} from './work-spreadsheet-package-scan-worker-protocol';
import {
  PLAIN_XLSX_CELL_BOOLEAN,
  PLAIN_XLSX_CELL_CHUNK_SIZE,
  PLAIN_XLSX_CELL_ERROR,
  PLAIN_XLSX_CELL_NUMBER,
  PLAIN_XLSX_CELL_TEXT,
  PLAIN_XLSX_COLUMN_BITS,
  PLAIN_XLSX_COLUMN_MASK,
  PLAIN_XLSX_ROW_CHUNK_SIZE,
  plainXlsxRangeExtent,
  type PlainXlsxCellChunk,
} from './work-xlsx-plain-fast-path';

const SPREADSHEET_PACKAGE_SCAN_WORKER_TIMEOUT_MS = 120_000;
const MAX_XLSX_COLUMN_COUNT = 16_384;
const MAX_XLSX_ROW_COUNT = 1_048_576;

interface PendingPlainWorksheet {
  columnCount: number | null;
  data: CellMatrix;
  expectedColumnCount: number | null;
  expectedRowCount: number | null;
  lastCoordinate: number;
  populatedCellCount: number | null;
  receivedCellCount: number;
  rowCount: number | null;
}

/**
 * Decompresses and classifies worksheet XML away from the UI thread. A null
 * result preserves the existing main-thread parser as a compatibility path.
 */
export function scanSpreadsheetPackageInWorker(
  bytes: ArrayBuffer,
  signal?: AbortSignal,
  onFastPathCandidate?: () => void,
): Promise<SpreadsheetPackageScanResult | null> {
  if (typeof Worker === 'undefined') return Promise.resolve(null);
  if (signal?.aborted) return Promise.reject(packageScanAbortError());

  let worker: Worker;
  try {
    worker = new Worker(
      new URL('./work-spreadsheet-package-scan.worker.js', import.meta.url),
      { name: 'a3s-office-spreadsheet-package-scan' },
    );
  } catch {
    return Promise.resolve(null);
  }

  const startedAt = spreadsheetImportNow();
  return new Promise((resolve, reject) => {
    let settled = false;
    let timeout: ReturnType<typeof setTimeout> | undefined;
    let fastPathCandidateReported = false;
    let fastPathRejected = false;
    let fastWorkbook: WorkBook | null = null;
    let pendingPlainWorksheets: Record<string, PendingPlainWorksheet> | null =
      null;
    const finish = (
      result: SpreadsheetPackageScanResult | null,
      error?: Error,
    ) => {
      if (settled) return;
      settled = true;
      if (timeout !== undefined) clearTimeout(timeout);
      signal?.removeEventListener('abort', handleAbort);
      worker.removeEventListener('error', handleFailure);
      worker.removeEventListener('message', handleMessage);
      worker.removeEventListener('messageerror', handleFailure);
      worker.terminate();
      if (error) reject(error);
      else resolve(result);
    };
    const handleAbort = () => finish(null, packageScanAbortError());
    const handleFailure = () => finish(null);
    const handleMessage = (
      event: MessageEvent<SpreadsheetPackageScanWorkerResponse>,
    ) => {
      const response = event.data;
      if (!response || typeof response !== 'object') {
        finish(null);
        return;
      }
      if (response.kind === 'failure') {
        finish(null);
        return;
      }
      if (response.kind === 'fast-path-rejected') {
        fastPathRejected = true;
        fastWorkbook = null;
        pendingPlainWorksheets = null;
        return;
      }
      if (response.kind === 'workbook') {
        if (fastPathRejected) return;
        if (fastWorkbook !== null || pendingPlainWorksheets !== null) {
          fastPathRejected = true;
          fastWorkbook = null;
          pendingPlainWorksheets = null;
          return;
        }
        const sheetNames = response.workbook?.SheetNames;
        if (
          !Array.isArray(sheetNames) ||
          !sheetNames.length ||
          sheetNames.some((name) => typeof name !== 'string') ||
          new Set(sheetNames).size !== sheetNames.length
        ) {
          fastPathRejected = true;
          fastWorkbook = null;
          pendingPlainWorksheets = null;
          return;
        }
        fastWorkbook = {
          ...response.workbook,
          Sheets: Object.fromEntries(sheetNames.map((name) => [name, []])),
        } as WorkBook;
        pendingPlainWorksheets = Object.fromEntries(
          sheetNames.map((name) => [
            name,
            {
              columnCount: null,
              data: [],
              expectedColumnCount: null,
              expectedRowCount: null,
              lastCoordinate: -1,
              populatedCellCount: null,
              receivedCellCount: 0,
              rowCount: null,
            } satisfies PendingPlainWorksheet,
          ]),
        );
        if (!fastPathCandidateReported) {
          fastPathCandidateReported = true;
          try {
            onFastPathCandidate?.();
          } catch {
            // Candidate notification cannot weaken the worker fallback.
          }
        }
        return;
      }
      if (response.kind === 'plain-worksheet-start') {
        if (fastPathRejected) return;
        const worksheet = pendingPlainWorksheets?.[response.name];
        if (
          !worksheet ||
          worksheet.expectedColumnCount !== null ||
          worksheet.expectedRowCount !== null ||
          worksheet.columnCount !== null ||
          worksheet.receivedCellCount !== 0 ||
          !validPlainWorksheetExtent(response.columnCount, response.rowCount)
        ) {
          fastPathRejected = true;
          fastWorkbook = null;
          pendingPlainWorksheets = null;
          return;
        }
        worksheet.expectedColumnCount = response.columnCount;
        worksheet.expectedRowCount = response.rowCount;
        return;
      }
      if (response.kind === 'plain-cells') {
        if (fastPathRejected) return;
        const worksheet = pendingPlainWorksheets?.[response.name];
        if (
          !worksheet ||
          worksheet.expectedColumnCount === null ||
          worksheet.expectedRowCount === null ||
          worksheet.columnCount !== null ||
          !appendPlainXlsxCellChunk(worksheet, response.chunk)
        ) {
          fastPathRejected = true;
          fastWorkbook = null;
          pendingPlainWorksheets = null;
          return;
        }
        return;
      }
      if (response.kind === 'worksheet') {
        if (fastPathRejected) return;
        const worksheet = fastWorkbook?.Sheets[response.name];
        const pending = pendingPlainWorksheets?.[response.name];
        if (
          !Array.isArray(worksheet) ||
          response.dense !== true ||
          !pending ||
          pending.expectedColumnCount !== response.columnCount ||
          pending.expectedRowCount !== response.rowCount ||
          pending.columnCount !== null ||
          !validPlainWorksheetSize(response)
        ) {
          fastPathRejected = true;
          fastWorkbook = null;
          pendingPlainWorksheets = null;
          return;
        }
        Object.assign(worksheet, response.properties);
        pending.columnCount = response.columnCount;
        pending.populatedCellCount = response.populatedCellCount;
        pending.rowCount = response.rowCount;
        return;
      }
      if (
        response.kind !== 'success' ||
        !response.worksheets ||
        typeof response.worksheets !== 'object'
      ) {
        finish(null);
        return;
      }
      const plainWorksheets =
        response.fastPath && !fastPathRejected
          ? completePlainWorksheets(fastWorkbook, pendingPlainWorksheets)
          : null;
      const workbook = plainWorksheets ? fastWorkbook : null;
      recordSpreadsheetImportMeasure(
        'a3s-office.spreadsheet.package-scan-worker',
        startedAt,
        spreadsheetImportNow(),
        {
          fastPath: workbook !== null,
          populatedCellCount: plainWorksheets
            ? Object.values(plainWorksheets).reduce(
                (sum, worksheet) => sum + worksheet.populatedCellCount,
                0,
              )
            : 0,
          protocol: plainWorksheets ? 'columnar-transfer' : undefined,
          worksheetCount: Object.keys(response.worksheets).length,
        },
      );
      finish({ plainWorksheets, workbook, worksheets: response.worksheets });
    };
    timeout = setTimeout(
      handleFailure,
      SPREADSHEET_PACKAGE_SCAN_WORKER_TIMEOUT_MS,
    );
    signal?.addEventListener('abort', handleAbort, { once: true });
    worker.addEventListener('error', handleFailure, { once: true });
    worker.addEventListener('message', handleMessage);
    worker.addEventListener('messageerror', handleFailure, { once: true });
    if (signal?.aborted) {
      handleAbort();
      return;
    }
    try {
      worker.postMessage({ bytes, kind: 'scan' }, { transfer: [bytes] });
    } catch {
      handleFailure();
    }
  });
}

function completePlainWorksheets(
  workbook: WorkBook | null,
  pending: Record<string, PendingPlainWorksheet> | null,
): Record<string, SpreadsheetPlainWorksheet> | null {
  if (!workbook || !pending) return null;
  const completed: Record<string, SpreadsheetPlainWorksheet> = {};
  for (const name of workbook.SheetNames) {
    const worksheet = workbook.Sheets[name];
    const candidate = pending[name];
    if (
      !Array.isArray(worksheet) ||
      typeof (worksheet as unknown as Record<string, unknown>)['!ref'] !==
        'string' ||
      !candidate ||
      candidate.columnCount === null ||
      candidate.populatedCellCount === null ||
      candidate.rowCount === null ||
      candidate.receivedCellCount !== candidate.populatedCellCount
    ) {
      return null;
    }
    completed[name] = {
      columnCount: candidate.columnCount,
      data: candidate.data,
      populatedCellCount: candidate.populatedCellCount,
      rowCount: candidate.rowCount,
    };
  }
  return Object.keys(completed).length === Object.keys(pending).length
    ? completed
    : null;
}

function appendPlainXlsxCellChunk(
  worksheet: PendingPlainWorksheet,
  chunk: PlainXlsxCellChunk,
): boolean {
  const columnCount = worksheet.expectedColumnCount;
  const rowCount = worksheet.expectedRowCount;
  if (
    columnCount === null ||
    rowCount === null ||
    !chunk ||
    !(chunk.coordinates instanceof Uint32Array) ||
    !(chunk.kinds instanceof Uint8Array) ||
    !(chunk.numericValues instanceof Float64Array) ||
    !Array.isArray(chunk.textValues) ||
    !Number.isSafeInteger(chunk.startRow) ||
    chunk.startRow < 0 ||
    chunk.startRow >= MAX_XLSX_ROW_COUNT ||
    chunk.startRow % PLAIN_XLSX_ROW_CHUNK_SIZE !== 0 ||
    chunk.coordinates.length === 0 ||
    chunk.coordinates.length > PLAIN_XLSX_CELL_CHUNK_SIZE ||
    chunk.coordinates.length !== chunk.kinds.length ||
    chunk.numericValues.length + chunk.textValues.length !==
      chunk.coordinates.length
  ) {
    return false;
  }

  let numericIndex = 0;
  let textIndex = 0;
  for (let index = 0; index < chunk.coordinates.length; index += 1) {
    const packed = chunk.coordinates[index] ?? 0;
    const rowOffset = packed >>> PLAIN_XLSX_COLUMN_BITS;
    const column = packed & PLAIN_XLSX_COLUMN_MASK;
    const row = chunk.startRow + rowOffset;
    const absoluteCoordinate = row * MAX_XLSX_COLUMN_COUNT + column;
    if (
      rowOffset >= PLAIN_XLSX_ROW_CHUNK_SIZE ||
      row >= rowCount ||
      column >= columnCount ||
      absoluteCoordinate <= worksheet.lastCoordinate
    ) {
      return false;
    }
    const kind = chunk.kinds[index];
    let cell: Cell;
    if (kind === PLAIN_XLSX_CELL_TEXT) {
      const value = chunk.textValues[textIndex];
      if (typeof value !== 'string') return false;
      textIndex += 1;
      cell = Object.freeze({ v: value });
    } else {
      const value = chunk.numericValues[numericIndex];
      if (!Number.isFinite(value)) return false;
      numericIndex += 1;
      if (kind === PLAIN_XLSX_CELL_NUMBER) cell = Object.freeze({ v: value });
      else if (kind === PLAIN_XLSX_CELL_BOOLEAN) {
        if (value !== 0 && value !== 1) return false;
        cell = Object.freeze({ v: value === 1 });
      } else if (kind === PLAIN_XLSX_CELL_ERROR) {
        const display = PLAIN_XLSX_ERROR_TEXT[value];
        if (!display) return false;
        cell = freezeImportedSpreadsheetCell({
          ct: { t: 'e' },
          m: display,
          v: value,
        });
      } else return false;
    }
    let rowCells = worksheet.data[row];
    if (!rowCells) {
      rowCells = [];
      worksheet.data[row] = rowCells;
    }
    rowCells[column] = cell;
    worksheet.lastCoordinate = absoluteCoordinate;
  }
  if (
    numericIndex !== chunk.numericValues.length ||
    textIndex !== chunk.textValues.length
  ) {
    return false;
  }
  worksheet.receivedCellCount += chunk.coordinates.length;
  return true;
}

function validPlainWorksheetExtent(
  columnCount: number,
  rowCount: number,
): boolean {
  return (
    Number.isSafeInteger(columnCount) &&
    columnCount > 0 &&
    columnCount <= MAX_XLSX_COLUMN_COUNT &&
    Number.isSafeInteger(rowCount) &&
    rowCount > 0 &&
    rowCount <= MAX_XLSX_ROW_COUNT
  );
}

function validPlainWorksheetSize(
  response: Extract<
    SpreadsheetPackageScanWorkerResponse,
    { kind: 'worksheet' }
  >,
): boolean {
  const properties = response.properties;
  const reference =
    properties !== null &&
    typeof properties === 'object' &&
    !Array.isArray(properties) &&
    Object.keys(properties).length === 1 &&
    Object.hasOwn(properties, '!ref') &&
    typeof properties['!ref'] === 'string'
      ? properties['!ref']
      : null;
  const extent = reference ? plainXlsxRangeExtent(reference) : null;
  return (
    validPlainWorksheetExtent(response.columnCount, response.rowCount) &&
    Number.isSafeInteger(response.populatedCellCount) &&
    response.populatedCellCount > 0 &&
    response.populatedCellCount <= response.rowCount * response.columnCount &&
    extent?.columnCount === response.columnCount &&
    extent.rowCount === response.rowCount
  );
}

const PLAIN_XLSX_ERROR_TEXT: Readonly<Record<number, string>> = {
  0: '#NULL!',
  7: '#DIV/0!',
  15: '#VALUE!',
  23: '#REF!',
  29: '#NAME?',
  36: '#NUM!',
  42: '#N/A',
  43: '#GETTING_DATA',
};

function packageScanAbortError(): Error {
  if (typeof DOMException !== 'undefined') {
    return new DOMException('Spreadsheet import was cancelled.', 'AbortError');
  }
  const error = new Error('Spreadsheet import was cancelled.');
  error.name = 'AbortError';
  return error;
}
