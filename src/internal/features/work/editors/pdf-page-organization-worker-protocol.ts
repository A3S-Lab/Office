import type {
  PdfPageOrganizationErrorCode,
  PdfPageOrganizationJob,
  PdfPageOrganizationResult,
} from './pdf-page-organization-types';

export interface PdfPageOrganizationWorkerRequest {
  job: PdfPageOrganizationJob;
  kind: 'run';
}

export type PdfPageOrganizationWorkerResponse =
  | {
      kind: 'success';
      result: PdfPageOrganizationResult;
    }
  | {
      code: PdfPageOrganizationErrorCode;
      kind: 'failure';
      message: string;
    };
