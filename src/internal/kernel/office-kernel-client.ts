import { layoutOfficeDocumentInJavaScript } from './office-kernel-fallback';
import { alignOfficePresentationInJavaScript } from './office-kernel-presentation-fallback';
import { layoutOfficeTextInJavaScript } from './office-kernel-text-fallback';
import {
  type OfficeKernelFontSource,
  type OfficeKernelLayoutRequest,
  type OfficeKernelLayoutResult,
  type OfficeKernelPresentationGeometryRequest,
  type OfficeKernelPresentationGeometryResult,
  type OfficeKernelTextLayoutRequest,
  type OfficeKernelTextLayoutResult,
  type OfficeKernelWorkerResponse,
  OFFICE_KERNEL_PROTOCOL_VERSION,
} from './office-kernel-protocol';

interface PendingRequest {
  reject: (error: Error) => void;
  removeAbortListener?: () => void;
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

export interface OfficeKernelLayoutInput {
  revision: number;
  documentRevision: number;
  startPageIndex?: number;
  page: OfficeKernelLayoutRequest['page'];
  blocks: OfficeKernelLayoutRequest['blocks'];
}

export interface OfficeKernelPresentationGeometryInput {
  revision: number;
  documentRevision: number;
  operation: OfficeKernelPresentationGeometryRequest['operation'];
  elements: OfficeKernelPresentationGeometryRequest['elements'];
}

export interface OfficeKernelTextLayoutInput {
  revision: number;
  documentRevision: number;
  paragraphs: OfficeKernelTextLayoutRequest['paragraphs'];
}

export interface OfficeKernelClient {
  layout(
    input: OfficeKernelLayoutInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelLayoutResult>;
  presentationGeometry(
    input: OfficeKernelPresentationGeometryInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelPresentationGeometryResult>;
  textLayout(
    input: OfficeKernelTextLayoutInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelTextLayoutResult>;
  dispose(): void;
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
        alignOfficePresentationInJavaScript,
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
    const pending = layout ?? geometry ?? textLayout;
    if (!pending) return;
    this.pendingLayouts.delete(response.requestId);
    this.pendingPresentationGeometry.delete(response.requestId);
    this.pendingTextLayouts.delete(response.requestId);
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
    pending.reject(new Error('Office kernel response kind did not match.'));
  };

  private handleWorkerFailure = (): void => {
    const layouts = [...this.pendingLayouts.values()];
    const geometryRequests = [...this.pendingPresentationGeometry.values()];
    const textLayoutRequests = [...this.pendingTextLayouts.values()];
    this.pendingLayouts.clear();
    this.pendingPresentationGeometry.clear();
    this.pendingTextLayouts.clear();
    this.worker?.terminate();
    this.worker = null;
    for (const pending of layouts) {
      releasePendingRequest(pending);
      Promise.resolve()
        .then(() => layoutOfficeDocumentInJavaScript(pending.request))
        .then(pending.resolve, pending.reject);
    }
    for (const pending of geometryRequests) {
      releasePendingRequest(pending);
      Promise.resolve()
        .then(() => alignOfficePresentationInJavaScript(pending.request))
        .then(pending.resolve, pending.reject);
    }
    for (const pending of textLayoutRequests) {
      releasePendingRequest(pending);
      Promise.resolve()
        .then(() => layoutOfficeTextInJavaScript(pending.request))
        .then(pending.resolve, pending.reject);
    }
  };
}

function fallbackRequest<Request, Result>(
  request: Request,
  signal: AbortSignal | undefined,
  execute: (request: Request) => Result,
): Promise<Result> {
  return Promise.resolve().then(() => {
    if (signal?.aborted) throw abortError();
    return execute(request);
  });
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
