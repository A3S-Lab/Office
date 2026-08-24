import { applyPdfPageOrganizationJob } from './pdf-page-organization-engine';
import { PdfPageOrganizationError } from './pdf-page-organization-types';
import type {
  PdfPageOrganizationWorkerRequest,
  PdfPageOrganizationWorkerResponse,
} from './pdf-page-organization-worker-protocol';

interface PdfPageOrganizationWorkerScope {
  onmessage:
    | ((event: MessageEvent<PdfPageOrganizationWorkerRequest>) => void)
    | null;
  postMessage: (
    message: PdfPageOrganizationWorkerResponse,
    options?: StructuredSerializeOptions,
  ) => void;
}

const scope = globalThis as unknown as PdfPageOrganizationWorkerScope;

scope.onmessage = (event) => {
  if (event.data.kind !== 'run') return;
  void applyPdfPageOrganizationJob(event.data.job)
    .then((result) => {
      const transfer =
        result.kind === 'mutated'
          ? [result.source.buffer]
          : result.files.map((file) => file.source.buffer);
      scope.postMessage({ kind: 'success', result }, { transfer });
    })
    .catch((error: unknown) => {
      const typed =
        error instanceof PdfPageOrganizationError
          ? error
          : new PdfPageOrganizationError(
              'pdf.pages.invalid-source',
              error instanceof Error
                ? error.message
                : 'Unable to organize this PDF.',
            );
      scope.postMessage({
        code: typed.code,
        kind: 'failure',
        message: typed.message,
      });
    });
};
