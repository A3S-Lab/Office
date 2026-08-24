import {
  PdfPageOrganizationError,
  type PdfPageOrganizationJob,
  type PdfPageOrganizationResult,
  type PdfPageOrganizationSource,
} from './pdf-page-organization-types';
import type {
  PdfPageOrganizationWorkerRequest,
  PdfPageOrganizationWorkerResponse,
} from './pdf-page-organization-worker-protocol';

const PDF_PAGE_ORGANIZATION_WORKER_TIMEOUT_MS = 120_000;

export async function runPdfPageOrganizationJob(
  job: PdfPageOrganizationJob,
  signal?: AbortSignal,
): Promise<PdfPageOrganizationResult> {
  if (signal?.aborted) throw pdfPageOrganizationAbortError();
  if (typeof Worker === 'undefined') return runDirectly(job, signal);

  let worker: Worker;
  try {
    worker = new Worker(
      new URL('./pdf-page-organization.worker.js', import.meta.url),
      { name: 'a3s-office-pdf-page-organization' },
    );
  } catch {
    return runDirectly(job, signal);
  }

  return new Promise((resolve, reject) => {
    let settled = false;
    const timeout = setTimeout(() => {
      finish(
        undefined,
        new PdfPageOrganizationError(
          'pdf.pages.invalid-source',
          'PDF page organization timed out.',
        ),
      );
    }, PDF_PAGE_ORGANIZATION_WORKER_TIMEOUT_MS);
    const finish = (
      result?: PdfPageOrganizationResult,
      error?: Error,
    ): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      signal?.removeEventListener('abort', handleAbort);
      worker.removeEventListener('error', handleWorkerError);
      worker.removeEventListener('message', handleMessage);
      worker.removeEventListener('messageerror', handleWorkerError);
      worker.terminate();
      if (error) reject(error);
      else if (result) resolve(result);
      else reject(new Error('PDF page worker returned no result.'));
    };
    const handleAbort = () =>
      finish(undefined, pdfPageOrganizationAbortError());
    const handleWorkerError = () =>
      finish(
        undefined,
        new PdfPageOrganizationError(
          'pdf.pages.invalid-source',
          'The PDF page worker failed before producing a result.',
        ),
      );
    const handleMessage = (
      event: MessageEvent<PdfPageOrganizationWorkerResponse>,
    ) => {
      const response = event.data;
      if (!response || typeof response !== 'object') {
        handleWorkerError();
        return;
      }
      if (response.kind === 'failure') {
        finish(
          undefined,
          new PdfPageOrganizationError(response.code, response.message),
        );
        return;
      }
      if (response.kind === 'success') finish(response.result);
      else handleWorkerError();
    };

    signal?.addEventListener('abort', handleAbort, { once: true });
    worker.addEventListener('error', handleWorkerError, { once: true });
    worker.addEventListener('message', handleMessage);
    worker.addEventListener('messageerror', handleWorkerError, { once: true });
    if (signal?.aborted) {
      handleAbort();
      return;
    }
    const transferableJob = transferablePdfPageJob(job);
    const request: PdfPageOrganizationWorkerRequest = {
      job: transferableJob.job,
      kind: 'run',
    };
    try {
      worker.postMessage(request, { transfer: transferableJob.transfer });
    } catch {
      handleWorkerError();
    }
  });
}

async function runDirectly(
  job: PdfPageOrganizationJob,
  signal?: AbortSignal,
): Promise<PdfPageOrganizationResult> {
  if (signal?.aborted) throw pdfPageOrganizationAbortError();
  const { applyPdfPageOrganizationJob } = await import(
    './pdf-page-organization-engine'
  );
  if (signal?.aborted) throw pdfPageOrganizationAbortError();
  const result = await applyPdfPageOrganizationJob(job);
  if (signal?.aborted) throw pdfPageOrganizationAbortError();
  return result;
}

function transferablePdfPageJob(job: PdfPageOrganizationJob): {
  job: PdfPageOrganizationJob;
  transfer: Transferable[];
} {
  const source = transferableSource(job.source);
  if (job.kind === 'export') {
    return {
      job: { ...job, source: source.source },
      transfer: [source.source],
    };
  }
  const imported = job.importSource
    ? transferableSource(job.importSource)
    : undefined;
  return {
    job: {
      ...job,
      ...(imported ? { importSource: imported.source } : {}),
      source: source.source,
    },
    transfer: imported ? [source.source, imported.source] : [source.source],
  };
}

function transferableSource(source: PdfPageOrganizationSource): {
  source: ArrayBuffer;
} {
  if (source instanceof ArrayBuffer) return { source };
  const copy = new Uint8Array(source.byteLength);
  copy.set(source);
  return { source: copy.buffer };
}

function pdfPageOrganizationAbortError(): Error {
  if (typeof DOMException !== 'undefined') {
    return new DOMException(
      'PDF page organization was cancelled.',
      'AbortError',
    );
  }
  const error = new Error('PDF page organization was cancelled.');
  error.name = 'AbortError';
  return error;
}
