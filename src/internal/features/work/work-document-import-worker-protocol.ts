import type {
  LargeSimpleDocxParseOptions,
  LargeSimpleDocxStreamResult,
} from './work-docx-large-document-parser';

export const DOCUMENT_IMPORT_PARAGRAPH_BATCH_SIZE = 2_048;
export const DOCUMENT_IMPORT_TABLE_ROW_BATCH_SIZE = 2_048;

export interface DocumentImportWorkerRequest {
  kind: 'parse';
  options: LargeSimpleDocxParseOptions;
  xmlBytes: ArrayBuffer;
}

export interface DocumentImportWorkerTimings {
  contentMs: number;
  eligibilityMs: number;
  envelopeMs: number;
  parseMs: number;
  xmlMs: number;
}

export type DocumentImportWorkerResponse =
  | { kind: 'failure' }
  | { kind: 'ineligible' }
  | {
      kind: 'paragraphs';
      texts: string[];
    }
  | { kind: 'table-start' }
  | {
      cellParagraphCounts: Uint32Array;
      kind: 'table-rows';
      rowCellCounts: Uint32Array;
      texts: string[];
    }
  | { kind: 'table-end' }
  | {
      kind: 'success';
      streamed: LargeSimpleDocxStreamResult;
      timings: DocumentImportWorkerTimings;
    };
