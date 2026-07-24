import { layoutOfficeDocumentInJavaScript } from './office-kernel-fallback';
import { resolveOfficePresentationGeometryInJavaScript } from './office-kernel-presentation-fallback';
import { calculateSpreadsheetInJavaScript } from './office-kernel-spreadsheet-fallback';
import { calculateSpreadsheetSessionInJavaScript } from './office-kernel-spreadsheet-session-fallback';
import { layoutOfficeTextInJavaScript } from './office-kernel-text-fallback';
import type {
  OfficeKernelClient,
  OfficeKernelLayoutInput,
  OfficeKernelPresentationGeometryInput,
  OfficeKernelSpreadsheetCalculationInput,
  OfficeKernelSpreadsheetSessionCalculationInput,
  OfficeKernelTextLayoutInput,
} from './office-kernel-client-types';
import {
  type OfficeKernelFontSource,
  type OfficeKernelLayoutRequest,
  type OfficeKernelLayoutResult,
  type OfficeKernelPresentationGeometryRequest,
  type OfficeKernelPresentationGeometryResult,
  type OfficeKernelSpreadsheetCalculationRequest,
  type OfficeKernelSpreadsheetCalculationResult,
  type OfficeKernelSpreadsheetSessionCalculationRequest,
  type OfficeKernelSpreadsheetSessionCalculationResult,
  type OfficeKernelTextLayoutRequest,
  type OfficeKernelTextLayoutResult,
  type OfficeKernelWorkerResponse,
  OFFICE_KERNEL_PROTOCOL_VERSION,
} from './office-kernel-protocol';

export type {
  OfficeKernelClient,
  OfficeKernelLayoutInput,
  OfficeKernelPresentationGeometryInput,
  OfficeKernelSpreadsheetCalculationInput,
  OfficeKernelSpreadsheetSessionCalculationInput,
  OfficeKernelTextLayoutInput,
} from './office-kernel-client-types';

interface PendingRequest {
  reject: (error: Error) => void;
  removeAbortListener?: () => void;
  signal?: AbortSignal;
  timeout?: ReturnType<typeof setTimeout>;
}

interface PendingLayout extends PendingRequest {
  request: OfficeKernelLayoutRequest;
  resolve: (result: OfficeKernelLayoutResult) => void;
}

interface PendingPresentationGeometry extends PendingRequest {
  request: OfficeKernelPresentationGeometryRequest;
  resolve: (result: OfficeKernelPresentationGeometryResult) => void;
}

interface PendingTextLayout extends PendingRequest {
  request: OfficeKernelTextLayoutRequest;
  resolve: (result: OfficeKernelTextLayoutResult) => void;
}

interface PendingSpreadsheetCalculation extends PendingRequest {
  request: OfficeKernelSpreadsheetCalculationRequest;
  resolve: (result: OfficeKernelSpreadsheetCalculationResult) => void;
}

interface PendingSpreadsheetSessionCalculation extends PendingRequest {
  request: OfficeKernelSpreadsheetSessionCalculationRequest;
  fallbackSheets: OfficeKernelSpreadsheetCalculationRequest['sheets'];
  resolve: (result: OfficeKernelSpreadsheetSessionCalculationResult) => void;
}

const OFFICE_KERNEL_WORKER_TIMEOUT_MS = 10_000;

export function createOfficeKernelClient(
  wasmUrl?: string,
  fonts: readonly OfficeKernelFontSource[] = [],
): OfficeKernelClient {
  return new BrowserOfficeKernelClient(wasmUrl, fonts);
}

class BrowserOfficeKernelClient implements OfficeKernelClient {
  private worker: Worker | null = null;
  private nextRequestId = 1;
  private pendingLayouts = new Map<number, PendingLayout>();
  private pendingPresentationGeometry = new Map<
    number,
    PendingPresentationGeometry
  >();
  private pendingTextLayouts = new Map<number, PendingTextLayout>();
  private pendingSpreadsheetCalculations = new Map<
    number,
    PendingSpreadsheetCalculation
  >();
  private pendingSpreadsheetSessionCalculations = new Map<
    number,
    PendingSpreadsheetSessionCalculation
  >();
  private disposed = false;

  constructor(wasmUrl?: string, fonts: readonly OfficeKernelFontSource[] = []) {
    if (typeof Worker === 'undefined') return;
    try {
      this.worker = new Worker(
        new URL('./office-kernel.worker.js', import.meta.url),
        {
          name: 'a3s-office-kernel',
        },
      );
      this.worker.addEventListener('message', this.handleMessage);
      this.worker.addEventListener('error', this.handleWorkerFailure);
      this.worker.postMessage({ kind: 'initialize', wasmUrl, fonts });
    } catch {
      this.worker = null;
    }
  }

  layout(
    input: OfficeKernelLayoutInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelLayoutResult> {
    if (this.disposed) return Promise.reject(disposedError());
    const request: OfficeKernelLayoutRequest = {
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'layout',
      requestId: this.nextRequestId++,
      revision: input.revision,
      documentRevision: input.documentRevision,
      startPageIndex: input.startPageIndex ?? 0,
      page: input.page,
      blocks: input.blocks,
    };
    if (signal?.aborted) return Promise.reject(abortError());
    if (!this.worker) {
      return fallbackRequest(request, signal, layoutOfficeDocumentInJavaScript);
    }
    return new Promise((resolve, reject) => {
      const pending: PendingLayout = { request, resolve, reject };
      this.attachPendingRequest(this.pendingLayouts, pending, signal);
      this.worker?.postMessage({ kind: 'layout', request });
    });
  }

  presentationGeometry(
    input: OfficeKernelPresentationGeometryInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelPresentationGeometryResult> {
    if (this.disposed) return Promise.reject(disposedError());
    const request: OfficeKernelPresentationGeometryRequest = {
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'presentationGeometry',
      requestId: this.nextRequestId++,
      revision: input.revision,
      documentRevision: input.documentRevision,
      operation: input.operation,
      elements: input.elements,
    };
    if (signal?.aborted) return Promise.reject(abortError());
    if (!this.worker) {
      return fallbackRequest(
        request,
        signal,
        resolveOfficePresentationGeometryInJavaScript,
      );
    }
    return new Promise((resolve, reject) => {
      const pending: PendingPresentationGeometry = {
        request,
        resolve,
        reject,
      };
      this.attachPendingRequest(
        this.pendingPresentationGeometry,
        pending,
        signal,
      );
      this.worker?.postMessage({ kind: 'presentationGeometry', request });
    });
  }

  textLayout(
    input: OfficeKernelTextLayoutInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelTextLayoutResult> {
    if (this.disposed) return Promise.reject(disposedError());
    const request: OfficeKernelTextLayoutRequest = {
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'textLayout',
      requestId: this.nextRequestId++,
      revision: input.revision,
      documentRevision: input.documentRevision,
      paragraphs: input.paragraphs,
    };
    if (signal?.aborted) return Promise.reject(abortError());
    if (!this.worker) {
      return fallbackRequest(request, signal, layoutOfficeTextInJavaScript);
    }
    return new Promise((resolve, reject) => {
      const pending: PendingTextLayout = { request, resolve, reject };
      this.attachPendingRequest(this.pendingTextLayouts, pending, signal);
      this.worker?.postMessage({ kind: 'textLayout', request });
    });
  }

  spreadsheetCalculation(
    input: OfficeKernelSpreadsheetCalculationInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelSpreadsheetCalculationResult> {
    if (this.disposed) return Promise.reject(disposedError());
    const request: OfficeKernelSpreadsheetCalculationRequest = {
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'spreadsheetCalculation',
      requestId: this.nextRequestId++,
      revision: input.revision,
      documentRevision: input.documentRevision,
      sheets: input.sheets,
      targets: input.targets,
    };
    if (signal?.aborted) return Promise.reject(abortError());
    if (!this.worker) {
      return fallbackRequest(request, signal, calculateSpreadsheetInJavaScript);
    }
    return new Promise((resolve, reject) => {
      const pending: PendingSpreadsheetCalculation = {
        request,
        resolve,
        reject,
      };
      this.attachPendingRequest(
        this.pendingSpreadsheetCalculations,
        pending,
        signal,
      );
      this.worker?.postMessage({ kind: 'spreadsheetCalculation', request });
    });
  }

  spreadsheetSessionCalculation(
    input: OfficeKernelSpreadsheetSessionCalculationInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelSpreadsheetSessionCalculationResult> {
    if (this.disposed) return Promise.reject(disposedError());
    const request: OfficeKernelSpreadsheetSessionCalculationRequest = {
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'spreadsheetSessionCalculation',
      requestId: this.nextRequestId++,
      revision: input.revision,
      documentRevision: input.documentRevision,
      update: input.update,
      calculation: input.calculation,
    };
    if (signal?.aborted) return Promise.reject(abortError());
    if (!this.worker) {
      return fallbackSpreadsheetSessionCalculation(
        request,
        input.fallbackSheets,
        signal,
      );
    }
    return new Promise((resolve, reject) => {
      const pending: PendingSpreadsheetSessionCalculation = {
        request,
        fallbackSheets: input.fallbackSheets,
        resolve,
        reject,
      };
      this.attachPendingRequest(
        this.pendingSpreadsheetSessionCalculations,
        pending,
        signal,
      );
      this.worker?.postMessage({
        kind: 'spreadsheetSessionCalculation',
        request,
      });
    });
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.worker?.removeEventListener('message', this.handleMessage);
    this.worker?.removeEventListener('error', this.handleWorkerFailure);
    this.worker?.terminate();
    this.worker = null;
    this.rejectPendingRequests(this.pendingLayouts, disposedError());
    this.rejectPendingRequests(
      this.pendingPresentationGeometry,
      disposedError(),
    );
    this.rejectPendingRequests(this.pendingTextLayouts, disposedError());
    this.rejectPendingRequests(
      this.pendingSpreadsheetCalculations,
      disposedError(),
    );
    this.rejectPendingRequests(
      this.pendingSpreadsheetSessionCalculations,
      disposedError(),
    );
  }

  private attachPendingRequest<T extends PendingRequest>(
    pendingRequests: Map<number, T>,
    pending: T & {
      request: { requestId: number };
    },
    signal?: AbortSignal,
  ): void {
    const requestId = pending.request.requestId;
    if (signal) {
      pending.signal = signal;
      const onAbort = () => {
        pendingRequests.delete(requestId);
        releasePendingRequest(pending);
        this.worker?.postMessage({ kind: 'cancel', requestId });
        pending.reject(abortError());
      };
      signal.addEventListener('abort', onAbort, { once: true });
      pending.removeAbortListener = () =>
        signal.removeEventListener('abort', onAbort);
    }
    pending.timeout = setTimeout(
      this.handleWorkerFailure,
      OFFICE_KERNEL_WORKER_TIMEOUT_MS,
    );
    pendingRequests.set(requestId, pending);
  }

  private rejectPendingRequests<T extends PendingRequest>(
    pendingRequests: Map<number, T>,
    error: Error,
  ): void {
    for (const pending of pendingRequests.values()) {
      releasePendingRequest(pending);
      pending.reject(error);
    }
    pendingRequests.clear();
  }

  private handleMessage = (
    event: MessageEvent<OfficeKernelWorkerResponse>,
  ): void => {
    if (event.data.kind !== 'response') return;
    const response = event.data.response;
    const layout = this.pendingLayouts.get(response.requestId);
    const geometry = this.pendingPresentationGeometry.get(response.requestId);
    const textLayout = this.pendingTextLayouts.get(response.requestId);
    const spreadsheetCalculation = this.pendingSpreadsheetCalculations.get(
      response.requestId,
    );
    const spreadsheetSessionCalculation =
      this.pendingSpreadsheetSessionCalculations.get(response.requestId);
    const pending =
      layout ??
      geometry ??
      textLayout ??
      spreadsheetCalculation ??
      spreadsheetSessionCalculation;
    if (!pending) return;
    this.pendingLayouts.delete(response.requestId);
    this.pendingPresentationGeometry.delete(response.requestId);
    this.pendingTextLayouts.delete(response.requestId);
    this.pendingSpreadsheetCalculations.delete(response.requestId);
    this.pendingSpreadsheetSessionCalculations.delete(response.requestId);
    releasePendingRequest(pending);
    if (response.kind === 'error') {
      pending.reject(
        Object.assign(new Error(response.error.message), {
          code: response.error.code,
        }),
      );
      return;
    }
    if (response.kind === 'layoutResult' && layout) {
      layout.resolve(response);
      return;
    }
    if (response.kind === 'presentationGeometryResult' && geometry) {
      geometry.resolve(response);
      return;
    }
    if (response.kind === 'textLayoutResult' && textLayout) {
      textLayout.resolve(response);
      return;
    }
    if (
      response.kind === 'spreadsheetCalculationResult' &&
      spreadsheetCalculation
    ) {
      spreadsheetCalculation.resolve(response);
      return;
    }
    if (
      response.kind === 'spreadsheetSessionCalculationResult' &&
      spreadsheetSessionCalculation
    ) {
      spreadsheetSessionCalculation.resolve(response);
      return;
    }
    pending.reject(new Error('Office kernel response kind did not match.'));
  };

  private handleWorkerFailure = (): void => {
    const layouts = [...this.pendingLayouts.values()];
    const geometryRequests = [...this.pendingPresentationGeometry.values()];
    const textLayoutRequests = [...this.pendingTextLayouts.values()];
    const spreadsheetCalculations = [
      ...this.pendingSpreadsheetCalculations.values(),
    ];
    const spreadsheetSessionCalculations = [
      ...this.pendingSpreadsheetSessionCalculations.values(),
    ];
    this.pendingLayouts.clear();
    this.pendingPresentationGeometry.clear();
    this.pendingTextLayouts.clear();
    this.pendingSpreadsheetCalculations.clear();
    this.pendingSpreadsheetSessionCalculations.clear();
    this.worker?.terminate();
    this.worker = null;
    for (const pending of layouts) {
      releasePendingRequest(pending);
      void fallbackRequest(
        pending.request,
        pending.signal,
        layoutOfficeDocumentInJavaScript,
      ).then(pending.resolve, pending.reject);
    }
    for (const pending of geometryRequests) {
      releasePendingRequest(pending);
      void fallbackRequest(
        pending.request,
        pending.signal,
        resolveOfficePresentationGeometryInJavaScript,
      ).then(pending.resolve, pending.reject);
    }
    for (const pending of textLayoutRequests) {
      releasePendingRequest(pending);
      void fallbackRequest(
        pending.request,
        pending.signal,
        layoutOfficeTextInJavaScript,
      ).then(pending.resolve, pending.reject);
    }
    for (const pending of spreadsheetCalculations) {
      releasePendingRequest(pending);
      void fallbackRequest(
        pending.request,
        pending.signal,
        calculateSpreadsheetInJavaScript,
      ).then(pending.resolve, pending.reject);
    }
    for (const pending of spreadsheetSessionCalculations) {
      releasePendingRequest(pending);
      void fallbackSpreadsheetSessionCalculation(
        pending.request,
        pending.fallbackSheets,
        pending.signal,
      ).then(pending.resolve, pending.reject);
    }
  };
}

async function fallbackSpreadsheetSessionCalculation(
  request: OfficeKernelSpreadsheetSessionCalculationRequest,
  sheets: OfficeKernelSpreadsheetCalculationRequest['sheets'],
  signal?: AbortSignal,
): Promise<OfficeKernelSpreadsheetSessionCalculationResult> {
  return fallbackRequest(
    { request, sheets },
    signal,
    ({ request: fallbackRequest, sheets: fallbackSheets }) =>
      calculateSpreadsheetSessionInJavaScript(fallbackRequest, fallbackSheets),
  );
}

async function fallbackRequest<Request, Result>(
  request: Request,
  signal: AbortSignal | undefined,
  execute: (request: Request) => Result | Promise<Result>,
): Promise<Result> {
  if (signal?.aborted) throw abortError();
  const result = await execute(request);
  if (signal?.aborted) throw abortError();
  return result;
}

function releasePendingRequest(pending: PendingRequest): void {
  pending.removeAbortListener?.();
  if (pending.timeout !== undefined) clearTimeout(pending.timeout);
}

function abortError(): DOMException {
  return new DOMException('Office kernel request was cancelled.', 'AbortError');
}

function disposedError(): Error {
  return new Error('Office kernel client is disposed.');
}
