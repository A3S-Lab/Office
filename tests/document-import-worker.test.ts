import { describe, expect, test } from '@rstest/core';
import { parseLargeSimpleDocxInWorker } from '../src/internal/features/work/work-document-import-worker-client';
import type { DocumentImportWorkerResponse } from '../src/internal/features/work/work-document-import-worker-protocol';
import { documentModelUsesWindowing } from '../src/internal/features/work/work-document-windowing';
import type { WorkDocumentSectionLayout } from '../src/internal/features/work/work-types';

describe('document import worker', () => {
  test('assembles streamed paragraphs and terminates the worker', async () => {
    const worker = await withDocumentWorker(
      (instance) => {
        queueMicrotask(() => {
          instance.emit({
            kind: 'paragraphs',
            texts: ['one', 'two', 'three'],
          });
          instance.emit(successResponse(3, 3, 0));
        });
      },
      async () => {
        const attempt = await parseLargeSimpleDocxInWorker(
          new ArrayBuffer(16),
          {
            minimumLogicalBlocks: 3,
            windowing: {
              blockSize: 2,
              blockThreshold: 2,
              tableRowSize: 2,
              tableRowThreshold: 2,
            },
          },
        );
        expect(attempt?.status).toBe('accepted');
        if (attempt?.status !== 'accepted') {
          throw new Error('Expected an accepted document Worker result.');
        }
        expect(attempt.result.html).toContain(
          '<p>one</p><p>two</p><p>three</p>',
        );
        expect(documentModelUsesWindowing(attempt.result.root)).toBe(true);
        return RecordingDocumentWorker.instance;
      },
    );

    expect(worker?.terminated).toBe(true);
    expect(worker?.transferCount).toBe(1);
  });

  test('assembles streamed table row batches into one canonical table', async () => {
    const attempt = await withDocumentWorker(
      (instance) => {
        queueMicrotask(() => {
          instance.emit({ kind: 'table-start' });
          instance.emit({
            cellParagraphCounts: new Uint32Array([1]),
            kind: 'table-rows',
            rowCellCounts: new Uint32Array([1]),
            texts: ['one'],
          });
          instance.emit({
            cellParagraphCounts: new Uint32Array([1]),
            kind: 'table-rows',
            rowCellCounts: new Uint32Array([1]),
            texts: ['two'],
          });
          instance.emit({ kind: 'table-end' });
          instance.emit(successResponse(2, 2, 2));
        });
      },
      () =>
        parseLargeSimpleDocxInWorker(new ArrayBuffer(16), {
          minimumLogicalBlocks: 2,
          windowing: {
            blockSize: 2,
            blockThreshold: 100,
            tableRowSize: 1,
            tableRowThreshold: 2,
          },
        }),
    );

    expect(attempt?.status).toBe('accepted');
    if (attempt?.status !== 'accepted') {
      throw new Error('Expected an accepted document Worker result.');
    }
    expect(attempt.result.html.match(/<table/g)).toHaveLength(1);
    expect(attempt.result.html.match(/<tr/g)).toHaveLength(2);
    expect(documentModelUsesWindowing(attempt.result.root)).toBe(true);
  });

  test('distinguishes an ineligible document from a Worker failure', async () => {
    const ineligible = await withDocumentWorker(
      (instance) => queueMicrotask(() => instance.emit({ kind: 'ineligible' })),
      () => parseLargeSimpleDocxInWorker(new ArrayBuffer(8)),
    );
    expect(ineligible).toEqual({ status: 'ineligible' });

    const failed = await withDocumentWorker(
      (instance) => queueMicrotask(() => instance.emit({ kind: 'failure' })),
      () => parseLargeSimpleDocxInWorker(new ArrayBuffer(8)),
    );
    expect(failed).toBeNull();
  });

  test('cancels and terminates an active document Worker', async () => {
    const controller = new AbortController();
    const worker = await withDocumentWorker(
      () => undefined,
      async () => {
        const pending = parseLargeSimpleDocxInWorker(
          new ArrayBuffer(8),
          {},
          controller.signal,
        );
        controller.abort();
        await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
        return RecordingDocumentWorker.instance;
      },
    );
    expect(worker?.terminated).toBe(true);
  });
});

function successResponse(
  logicalBlockCount: number,
  paragraphCount: number,
  tableRowCount: number,
): DocumentImportWorkerResponse {
  return {
    kind: 'success',
    streamed: {
      layout: documentLayout(),
      logicalBlockCount,
      paragraphCount,
      tableRowCount,
    },
    timings: {
      contentMs: 1,
      eligibilityMs: 1,
      envelopeMs: 1,
      parseMs: 2,
      xmlMs: 1,
    },
  };
}

function documentLayout(): WorkDocumentSectionLayout {
  return {
    breakAfter: 'nextPage',
    columns: { count: 1, separator: false, spacing: 12 },
    margins: { bottom: 25.4, left: 25.4, right: 25.4, top: 25.4 },
    orientation: 'portrait',
    pageSize: 'a4',
  };
}

type WorkerPost = (worker: RecordingDocumentWorker) => void;

class RecordingDocumentWorker {
  static instance: RecordingDocumentWorker | null = null;
  static post: WorkerPost = () => undefined;

  readonly listeners = new Map<string, Set<EventListener>>();
  terminated = false;
  transferCount = 0;

  constructor() {
    RecordingDocumentWorker.instance = this;
  }

  addEventListener(type: string, listener: EventListener): void {
    const listeners = this.listeners.get(type) ?? new Set<EventListener>();
    listeners.add(listener);
    this.listeners.set(type, listeners);
  }

  removeEventListener(type: string, listener: EventListener): void {
    this.listeners.get(type)?.delete(listener);
  }

  postMessage(_message: unknown, options?: StructuredSerializeOptions): void {
    this.transferCount = options?.transfer?.length ?? 0;
    RecordingDocumentWorker.post(this);
  }

  terminate(): void {
    this.terminated = true;
  }

  emit(data: DocumentImportWorkerResponse): void {
    const event = new MessageEvent('message', { data });
    for (const listener of this.listeners.get('message') ?? []) listener(event);
  }
}

async function withDocumentWorker<T>(
  post: WorkerPost,
  run: () => Promise<T>,
): Promise<T> {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
  RecordingDocumentWorker.instance = null;
  RecordingDocumentWorker.post = post;
  Object.defineProperty(globalThis, 'Worker', {
    configurable: true,
    value: RecordingDocumentWorker as unknown as typeof Worker,
  });
  try {
    return await run();
  } finally {
    if (descriptor) Object.defineProperty(globalThis, 'Worker', descriptor);
    else Reflect.deleteProperty(globalThis, 'Worker');
  }
}
