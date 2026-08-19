import type { CellMatrix } from '@fortune-sheet/core';
import type { WorkBook } from 'xlsx';
import type { SpreadsheetImportWorkbookMetadata } from './work-spreadsheet-import-worker-protocol';
import type { PlainXlsxCellChunk } from './work-xlsx-plain-fast-path';
import type { XlsxWorksheetXmlScan } from './work-xlsx-worksheet-scan';

export interface SpreadsheetPackageScanWorkerRequest {
  bytes: ArrayBuffer;
  kind: 'scan';
}

export interface SpreadsheetPackageScanResult {
  plainWorksheets: Record<string, SpreadsheetPlainWorksheet> | null;
  workbook: WorkBook | null;
  worksheets: Record<string, XlsxWorksheetXmlScan>;
}

export interface SpreadsheetPlainWorksheet {
  columnCount: number;
  data: CellMatrix;
  populatedCellCount: number;
  rowCount: number;
}

export type SpreadsheetPackageScanWorkerResponse =
  | {
      kind: 'workbook';
      workbook: SpreadsheetImportWorkbookMetadata;
    }
  | {
      columnCount: number;
      kind: 'plain-worksheet-start';
      name: string;
      rowCount: number;
    }
  | {
      columnCount: number;
      dense: true;
      kind: 'worksheet';
      name: string;
      populatedCellCount: number;
      properties: Record<string, unknown>;
      rowCount: number;
    }
  | {
      chunk: PlainXlsxCellChunk;
      kind: 'plain-cells';
      name: string;
    }
  | { kind: 'fast-path-rejected' }
  | {
      fastPath: boolean;
      kind: 'success';
      worksheets: SpreadsheetPackageScanResult['worksheets'];
    }
  | { kind: 'failure' };
