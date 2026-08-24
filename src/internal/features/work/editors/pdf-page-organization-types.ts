export const PDF_PAGE_ORGANIZATION_MAX_SOURCE_BYTES = 256 * 1024 * 1024;
export const PDF_PAGE_ORGANIZATION_MAX_IMPORT_BYTES = 128 * 1024 * 1024;
export const PDF_PAGE_ORGANIZATION_MAX_PAGES = 4_096;
export const PDF_PAGE_ORGANIZATION_MIN_PAGE_POINTS = 18;
export const PDF_PAGE_ORGANIZATION_MAX_PAGE_POINTS = 14_400;

export type PdfPageOrganizationErrorCode =
  | 'pdf.pages.duplicate-index'
  | 'pdf.pages.empty-document'
  | 'pdf.pages.encrypted-source'
  | 'pdf.pages.invalid-index'
  | 'pdf.pages.invalid-order'
  | 'pdf.pages.invalid-page-size'
  | 'pdf.pages.invalid-source'
  | 'pdf.pages.invalid-split'
  | 'pdf.pages.missing-import'
  | 'pdf.pages.page-limit'
  | 'pdf.pages.risky-import'
  | 'pdf.pages.risky-structure'
  | 'pdf.pages.signed-source'
  | 'pdf.pages.source-limit'
  | 'pdf.pages.unsupported-rotation';

export class PdfPageOrganizationError extends Error {
  readonly code: PdfPageOrganizationErrorCode;

  constructor(code: PdfPageOrganizationErrorCode, message: string) {
    super(message);
    this.name = 'PdfPageOrganizationError';
    this.code = code;
  }
}

export interface PdfPageOrganizationSize {
  height: number;
  width: number;
}

export type PdfPageOrganizationMutation =
  | {
      degrees: 90 | 180 | 270;
      kind: 'rotate';
      pageIndexes: number[];
    }
  | { kind: 'delete'; pageIndexes: number[] }
  | { kind: 'reorder'; pageOrder: number[] }
  | { index: number; kind: 'insert-blank'; size?: PdfPageOrganizationSize }
  | { index: number; kind: 'merge' };

export type PdfPageOrganizationExportOperation =
  | { kind: 'extract'; pageIndexes: number[] }
  | { kind: 'split'; splitAfterPageIndexes: number[] };

export type PdfPageOrganizationSource = ArrayBuffer | Uint8Array;

export type PdfPageOrganizationJob =
  | {
      importSource?: PdfPageOrganizationSource;
      kind: 'mutate';
      mutation: PdfPageOrganizationMutation;
      source: PdfPageOrganizationSource;
    }
  | {
      kind: 'export';
      operation: PdfPageOrganizationExportOperation;
      source: PdfPageOrganizationSource;
    };

export interface PdfPageOrganizationDiagnostic {
  code: 'pdf.pages.catalog-not-copied';
  message: string;
}

export interface PdfPageOrganizationExportFile {
  pageCount: number;
  source: Uint8Array;
}

export type PdfPageOrganizationResult =
  | {
      diagnostics: PdfPageOrganizationDiagnostic[];
      kind: 'mutated';
      pageCount: number;
      source: Uint8Array;
    }
  | {
      diagnostics: PdfPageOrganizationDiagnostic[];
      files: PdfPageOrganizationExportFile[];
      kind: 'exported';
    };
