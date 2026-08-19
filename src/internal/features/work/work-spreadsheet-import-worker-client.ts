import type { ParsingOptions, WorkBook } from 'xlsx';
import type {
  SpreadsheetImportDenseRow,
  SpreadsheetImportWorkerResponse,
} from './work-spreadsheet-import-worker-protocol';
import {
  recordSpreadsheetImportMeasure,
  spreadsheetImportNow,
} from './work-spreadsheet-import-diagnostics';

const SPREADSHEET_IMPORT_WORKER_TIMEOUT_MS = 120_000;

/**
 * Parses a workbook away from the UI thread when a dedicated Worker can be
 * created. A null result asks the caller to use the synchronous fallback.
 */
export function readSpreadsheetWorkbookInWorker(
  bytes: ArrayBuffer,
  options: ParsingOptions,
  signal?: AbortSignal,
): Promise<WorkBook | null> {
  if (typeof Worker === 'undefined') return Promise.resolve(null);
  if (signal?.aborted) return Promise.reject(spreadsheetImportAbortError());

  let worker: Worker;
  try {
    worker = new Worker(
      new URL('./work-spreadsheet-import.worker.js', import.meta.url),
      { name: 'a3s-office-spreadsheet-import' },
    );
  } catch {
    return Promise.resolve(null);
  }

  const startedAt = spreadsheetImportNow();
  return new Promise((resolve, reject) => {
    let settled = false;
    let timeout: ReturnType<typeof setTimeout> | undefined;
    let workbook: WorkBook | null = null;
    let workbookReceivedAt: number | null = null;
    let rowChunkCount = 0;
    let sparseCellChunkCount = 0;
    const finish = (result: WorkBook | null, error?: Error) => {
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
    const handleAbort = () => finish(null, spreadsheetImportAbortError());
    const handleFailure = () => finish(null);
    const handleMessage = (
      event: MessageEvent<SpreadsheetImportWorkerResponse>,
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
      if (response.kind === 'workbook') {
        if (!Array.isArray(response.workbook?.SheetNames)) {
          finish(null);
          return;
        }
        workbook = { ...response.workbook, Sheets: {} } as WorkBook;
        workbookReceivedAt = spreadsheetImportNow();
        recordSpreadsheetImportMeasure(
          'a3s-office.spreadsheet.worker-ready',
          startedAt,
          workbookReceivedAt,
        );
        return;
      }
      if (response.kind === 'worksheet') {
        if (!workbook?.SheetNames.includes(response.name)) {
          finish(null);
          return;
        }
        const worksheet = response.dense ? [] : {};
        Object.assign(worksheet, response.properties);
        workbook.Sheets[response.name] = worksheet;
        return;
      }
      if (response.kind === 'rows') {
        rowChunkCount += 1;
        const worksheet = workbook?.Sheets[response.name];
        if (
          !Array.isArray(worksheet) ||
          !Number.isSafeInteger(response.startRow) ||
          response.startRow < 0 ||
          !Array.isArray(response.rows)
        ) {
          finish(null);
          return;
        }
        const rows = worksheet as unknown as Array<
          SpreadsheetImportDenseRow | undefined
        >;
        rows.length = Math.max(
          rows.length,
          response.startRow + response.rows.length,
        );
        for (let index = 0; index < response.rows.length; index += 1) {
          const row = response.rows[index];
          if (row !== undefined) rows[response.startRow + index] = row;
        }
        return;
      }
      if (response.kind === 'cells') {
        sparseCellChunkCount += 1;
        const worksheet = workbook?.Sheets[response.name];
        if (!worksheet || Array.isArray(worksheet)) {
          finish(null);
          return;
        }
        for (const [address, cell] of response.cells) {
          worksheet[address] = cell;
        }
        return;
      }
      if (response.kind === 'success') {
        const finishedAt = spreadsheetImportNow();
        if (workbookReceivedAt !== null) {
          recordSpreadsheetImportMeasure(
            'a3s-office.spreadsheet.worker-stream',
            workbookReceivedAt,
            finishedAt,
            { rowChunkCount, sparseCellChunkCount },
          );
        }
        recordSpreadsheetImportMeasure(
          'a3s-office.spreadsheet.worker-total',
          startedAt,
          finishedAt,
          { rowChunkCount, sparseCellChunkCount },
        );
        finish(workbook);
      }
    };
    timeout = setTimeout(handleFailure, SPREADSHEET_IMPORT_WORKER_TIMEOUT_MS);
    signal?.addEventListener('abort', handleAbort, { once: true });
    worker.addEventListener('error', handleFailure, { once: true });
    worker.addEventListener('message', handleMessage);
    worker.addEventListener('messageerror', handleFailure, { once: true });
    if (signal?.aborted) {
      handleAbort();
      return;
    }
    try {
      worker.postMessage(
        { bytes, kind: 'parse', options },
        { transfer: [bytes] },
      );
    } catch {
      handleFailure();
    }
  });
}

function spreadsheetImportAbortError(): Error {
  if (typeof DOMException !== 'undefined') {
    return new DOMException('Spreadsheet import was cancelled.', 'AbortError');
  }
  const error = new Error('Spreadsheet import was cancelled.');
  error.name = 'AbortError';
  return error;
}
