import type { CellObject, ParsingOptions, WorkBook } from 'xlsx';

export const SPREADSHEET_IMPORT_ROW_CHUNK_SIZE = 512;
export const SPREADSHEET_IMPORT_CELL_CHUNK_SIZE = 8_192;

export interface SpreadsheetImportWorkerRequest {
  bytes: ArrayBuffer;
  kind: 'parse';
  options: ParsingOptions;
}

export type SpreadsheetImportWorkbookMetadata = Omit<WorkBook, 'Sheets'>;
export type SpreadsheetImportDenseRow = Array<CellObject | undefined>;

export type SpreadsheetImportWorkerResponse =
  | {
      kind: 'workbook';
      workbook: SpreadsheetImportWorkbookMetadata;
    }
  | {
      dense: boolean;
      kind: 'worksheet';
      name: string;
      properties: Record<string, unknown>;
    }
  | {
      kind: 'rows';
      name: string;
      rows: Array<SpreadsheetImportDenseRow | undefined>;
      startRow: number;
    }
  | {
      cells: Array<[address: string, cell: CellObject]>;
      kind: 'cells';
      name: string;
    }
  | { kind: 'success' }
  | { kind: 'failure' };
