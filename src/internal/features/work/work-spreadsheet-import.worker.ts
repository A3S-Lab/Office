import * as XLSX from 'xlsx';
import type { CellObject, WorkBook, WorkSheet } from 'xlsx';
import {
  SPREADSHEET_IMPORT_CELL_CHUNK_SIZE,
  type SpreadsheetImportDenseRow,
  SPREADSHEET_IMPORT_ROW_CHUNK_SIZE,
  type SpreadsheetImportWorkbookMetadata,
  type SpreadsheetImportWorkerRequest,
  type SpreadsheetImportWorkerResponse,
} from './work-spreadsheet-import-worker-protocol';

interface SpreadsheetImportWorkerScope {
  onmessage:
    | ((event: MessageEvent<SpreadsheetImportWorkerRequest>) => void)
    | null;
  postMessage: (message: SpreadsheetImportWorkerResponse) => void;
}

const scope = globalThis as unknown as SpreadsheetImportWorkerScope;

scope.onmessage = (event) => {
  if (event.data.kind !== 'parse') return;
  try {
    const workbook = XLSX.read(event.data.bytes, event.data.options);
    scope.postMessage({
      kind: 'workbook',
      workbook: workbookMetadata(workbook),
    });
    for (const name of workbook.SheetNames) {
      const worksheet = workbook.Sheets[name];
      if (worksheet) streamWorksheet(name, worksheet);
    }
    scope.postMessage({ kind: 'success' });
  } catch {
    scope.postMessage({ kind: 'failure' });
  }
};

function workbookMetadata(
  workbook: WorkBook,
): SpreadsheetImportWorkbookMetadata {
  const metadata = { ...workbook } as Partial<WorkBook>;
  Reflect.deleteProperty(metadata, 'Sheets');
  return metadata as SpreadsheetImportWorkbookMetadata;
}

function streamWorksheet(name: string, worksheet: WorkSheet): void {
  const properties = worksheetProperties(worksheet);
  if (Array.isArray(worksheet)) {
    scope.postMessage({ kind: 'worksheet', dense: true, name, properties });
    const rows = worksheet as unknown as Array<
      SpreadsheetImportDenseRow | undefined
    >;
    for (
      let startRow = 0;
      startRow < rows.length;
      startRow += SPREADSHEET_IMPORT_ROW_CHUNK_SIZE
    ) {
      const endRow = Math.min(
        rows.length,
        startRow + SPREADSHEET_IMPORT_ROW_CHUNK_SIZE,
      );
      scope.postMessage({
        kind: 'rows',
        name,
        rows: rows.slice(startRow, endRow),
        startRow,
      });
      for (let row = startRow; row < endRow; row += 1) rows[row] = undefined;
    }
    return;
  }

  scope.postMessage({ kind: 'worksheet', dense: false, name, properties });
  let cells: Array<[string, CellObject]> = [];
  for (const address in worksheet) {
    if (address.startsWith('!')) continue;
    const value = worksheet[address];
    if (!value || typeof value !== 'object' || Array.isArray(value)) continue;
    cells.push([address, value as CellObject]);
    if (cells.length < SPREADSHEET_IMPORT_CELL_CHUNK_SIZE) continue;
    scope.postMessage({ cells, kind: 'cells', name });
    cells = [];
  }
  if (cells.length) scope.postMessage({ cells, kind: 'cells', name });
}

function worksheetProperties(worksheet: WorkSheet): Record<string, unknown> {
  const properties: Record<string, unknown> = {};
  const source = worksheet as unknown as Record<string, unknown>;
  for (const name in worksheet) {
    if (name.startsWith('!')) properties[name] = source[name];
  }
  return properties;
}
