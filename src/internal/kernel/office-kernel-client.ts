import { layoutOfficeDocumentInJavaScript } from './office-kernel-fallback';
import {
  type OfficeKernelLayoutRequest,
  type OfficeKernelLayoutResult,
  type OfficeKernelWorkerResponse,
  OFFICE_KERNEL_PROTOCOL_VERSION,
} from './office-kernel-protocol';

interface PendingLayout {
  request: OfficeKernelLayoutRequest;
  resolve: (result: OfficeKernelLayoutResult) => void;
  reject: (error: Error) => void;
  removeAbortListener?: () => void;
  timeout?: ReturnType<typeof setTimeout>;
}

export interface OfficeKernelLayoutInput {
  revision: number;
  page: OfficeKernelLayoutRequest['page'];
  blocks: OfficeKernelLayoutRequest['blocks'];
}

export interface OfficeKernelClient {
  layout(
    input: OfficeKernelLayoutInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelLayoutResult>;
  dispose(): void;
}

const OFFICE_KERNEL_WORKER_TIMEOUT_MS = 10_000;

export function createOfficeKernelClient(wasmUrl?: string): OfficeKernelClient {
  return new BrowserOfficeKernelClient(wasmUrl);
}

class BrowserOfficeKernelClient implements OfficeKernelClient {
  private worker: Worker | null = null;
  private nextRequestId = 1;
  private pending = new Map<number, PendingLayout>();
  private disposed = false;

  constructor(wasmUrl?: string) {
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
      this.worker.postMessage({ kind: 'initialize', wasmUrl });
    } catch {
      this.worker = null;
    }
  }

  layout(
    input: OfficeKernelLayoutInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelLayoutResult> {
    if (this.disposed)
      return Promise.reject(new Error('Office kernel client is disposed.'));
    const request: OfficeKernelLayoutRequest = {
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'layout',
      requestId: this.nextRequestId++,
      revision: input.revision,
      page: input.page,
      blocks: input.blocks,
    };
    if (signal?.aborted) return Promise.reject(abortError());
    if (!this.worker) {
      return Promise.resolve().then(() => {
        if (signal?.aborted) throw abortError();
        return layoutOfficeDocumentInJavaScript(request);
      });
    }
    return new Promise((resolve, reject) => {
      const pending: PendingLayout = { request, resolve, reject };
      if (signal) {
        const onAbort = () => {
          this.pending.delete(request.requestId);
          releasePendingLayout(pending);
          this.worker?.postMessage({
            kind: 'cancel',
            requestId: request.requestId,
          });
          reject(abortError());
        };
        signal.addEventListener('abort', onAbort, { once: true });
        pending.removeAbortListener = () =>
          signal.removeEventListener('abort', onAbort);
      }
      pending.timeout = setTimeout(
        this.handleWorkerFailure,
        OFFICE_KERNEL_WORKER_TIMEOUT_MS,
      );
      this.pending.set(request.requestId, pending);
      this.worker?.postMessage({ kind: 'layout', request });
    });
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.worker?.removeEventListener('message', this.handleMessage);
    this.worker?.removeEventListener('error', this.handleWorkerFailure);
    this.worker?.terminate();
    this.worker = null;
    for (const pending of this.pending.values()) {
      releasePendingLayout(pending);
      pending.reject(new Error('Office kernel client was disposed.'));
    }
    this.pending.clear();
  }

  private handleMessage = (
    event: MessageEvent<OfficeKernelWorkerResponse>,
  ): void => {
    if (event.data.kind !== 'response') return;
    const response = event.data.response;
    const pending = this.pending.get(response.requestId);
    if (!pending) return;
    this.pending.delete(response.requestId);
    releasePendingLayout(pending);
    if (response.kind === 'error') {
      pending.reject(
        Object.assign(new Error(response.error.message), {
          code: response.error.code,
        }),
      );
      return;
    }
    pending.resolve(response);
  };

  private handleWorkerFailure = (): void => {
    const pendingLayouts = [...this.pending.values()];
    this.pending.clear();
    this.worker?.terminate();
    this.worker = null;
    for (const pending of pendingLayouts) {
      releasePendingLayout(pending);
      Promise.resolve()
        .then(() => layoutOfficeDocumentInJavaScript(pending.request))
        .then(pending.resolve, pending.reject);
    }
  };
}

function releasePendingLayout(pending: PendingLayout): void {
  pending.removeAbortListener?.();
  if (pending.timeout !== undefined) clearTimeout(pending.timeout);
}

function abortError(): DOMException {
  return new DOMException('Office kernel request was cancelled.', 'AbortError');
}
