import { describe, expect, test } from '@rstest/core';
import type { WorkBook, WorkSheet } from 'xlsx';
import { readSpreadsheetWorkbookInWorker } from '../src/internal/features/work/work-spreadsheet-import-worker-client';
import { xlsxWorksheetCellEntries } from '../src/internal/features/work/work-xlsx-worksheet';

describe('spreadsheet import worker', () => {
  test('reconstructs streamed dense worksheet chunks and terminates its worker', async () => {
    const workbook = {
      SheetNames: ['Sheet 1'],
      Sheets: {
        'Sheet 1': Object.assign(
          [[{ t: 's', v: 'A1' }], undefined, [undefined, { t: 'n', v: 2 }]],
          { '!ref': 'A1:B3' },
        ) as unknown as WorkSheet,
      },
    } satisfies WorkBook;
    const worker = await withSpreadsheetWorker(
      (instance) => {
        queueMicrotask(() => {
          instance.emitMessage({
            kind: 'workbook',
            workbook: { SheetNames: ['Sheet 1'] },
          });
          instance.emitMessage({
            kind: 'worksheet',
            dense: true,
            name: 'Sheet 1',
            properties: { '!ref': 'A1:B3' },
          });
          instance.emitMessage({
            kind: 'rows',
            name: 'Sheet 1',
            rows: [[{ t: 's', v: 'A1' }], undefined],
            startRow: 0,
          });
          instance.emitMessage({
            kind: 'rows',
            name: 'Sheet 1',
            rows: [[undefined, { t: 'n', v: 2 }]],
            startRow: 2,
          });
          instance.emitMessage({ kind: 'success' });
        });
      },
      async () => {
        const result = await readSpreadsheetWorkbookInWorker(
          new ArrayBuffer(16),
          { type: 'array', dense: true },
        );
        expect(result).toEqual(workbook);
        return RecordingSpreadsheetWorker.instance;
      },
    );

    expect(worker?.terminated).toBe(true);
    expect(worker?.transferCount).toBe(1);
  });

  test('reconstructs streamed sparse worksheet cells', async () => {
    const result = await withSpreadsheetWorker(
      (instance) => {
        queueMicrotask(() => {
          instance.emitMessage({
            kind: 'workbook',
            workbook: { SheetNames: ['Sparse'] },
          });
          instance.emitMessage({
            kind: 'worksheet',
            dense: false,
            name: 'Sparse',
            properties: { '!ref': 'A1:XFD1048576' },
          });
          instance.emitMessage({
            cells: [
              ['A1', { t: 's', v: 'Anchor' }],
              ['XFD1048576', { t: 's', v: 'Tail' }],
            ],
            kind: 'cells',
            name: 'Sparse',
          });
          instance.emitMessage({ kind: 'success' });
        });
      },
      () =>
        readSpreadsheetWorkbookInWorker(new ArrayBuffer(16), {
          type: 'array',
        }),
    );

    expect(result).toEqual({
      SheetNames: ['Sparse'],
      Sheets: {
        Sparse: {
          '!ref': 'A1:XFD1048576',
          A1: { t: 's', v: 'Anchor' },
          XFD1048576: { t: 's', v: 'Tail' },
        },
      },
    });
  });

  test('falls back when the streamed protocol finishes without workbook metadata', async () => {
    const worker = await withSpreadsheetWorker(
      (instance) => {
        queueMicrotask(() => instance.emitMessage({ kind: 'success' }));
      },
      async () => {
        await expect(
          readSpreadsheetWorkbookInWorker(new ArrayBuffer(16), {
            type: 'array',
          }),
        ).resolves.toBeNull();
        return RecordingSpreadsheetWorker.instance;
      },
    );

    expect(worker?.terminated).toBe(true);
  });

  test('falls back and terminates when posting the request fails', async () => {
    const worker = await withSpreadsheetWorker(
      () => {
        throw new DOMException('Cannot clone request.', 'DataCloneError');
      },
      async () => {
        await expect(
          readSpreadsheetWorkbookInWorker(new ArrayBuffer(16), {
            type: 'array',
          }),
        ).resolves.toBeNull();
        return RecordingSpreadsheetWorker.instance;
      },
    );

    expect(worker?.terminated).toBe(true);
  });

  test('cancels an active worker with the caller AbortSignal', async () => {
    const controller = new AbortController();
    const worker = await withSpreadsheetWorker(
      () => undefined,
      async () => {
        const pending = readSpreadsheetWorkbookInWorker(
          new ArrayBuffer(16),
          { type: 'array' },
          controller.signal,
        );
        controller.abort();
        await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
        return RecordingSpreadsheetWorker.instance;
      },
    );

    expect(worker?.terminated).toBe(true);
  });
});

describe('XLSX worksheet iteration', () => {
  test('iterates dense rows without materializing sparse address entries', () => {
    const worksheet = [
      [{ v: 'A1', t: 's' }, undefined, { v: 'C1', t: 's' }],
      Array.from({ length: 27 }, (_, column) =>
        column === 26 ? { v: 'AA2', t: 's' } : undefined,
      ),
    ] as unknown as WorkSheet;

    expect(
      Array.from(xlsxWorksheetCellEntries(worksheet)).map(
        ({ address, column, row }) => ({ address, column, row }),
      ),
    ).toEqual([
      { address: 'A1', column: 0, row: 0 },
      { address: 'C1', column: 2, row: 0 },
      { address: 'AA2', column: 26, row: 1 },
    ]);
  });

  test('ignores metadata and malformed addresses in sparse worksheets', () => {
    const worksheet = {
      '!ref': 'A1:B2',
      A1: { v: 1, t: 'n' },
      B2: { v: 2, t: 'n' },
      invalid: { v: 3, t: 'n' },
    } as WorkSheet;

    expect(
      Array.from(xlsxWorksheetCellEntries(worksheet)).map(
        ({ address, column, row }) => ({ address, column, row }),
      ),
    ).toEqual([
      { address: 'A1', column: 0, row: 0 },
      { address: 'B2', column: 1, row: 1 },
    ]);
  });
});

type WorkerPost = (worker: RecordingSpreadsheetWorker) => void;

class RecordingSpreadsheetWorker {
  static instance: RecordingSpreadsheetWorker | null = null;
  static post: WorkerPost = () => undefined;

  readonly listeners = new Map<string, Set<EventListener>>();
  terminated = false;
  transferCount = 0;

  constructor() {
    RecordingSpreadsheetWorker.instance = this;
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
    RecordingSpreadsheetWorker.post(this);
  }

  terminate(): void {
    this.terminated = true;
  }

  emitMessage(data: unknown): void {
    const event = new MessageEvent('message', { data });
    for (const listener of this.listeners.get('message') ?? []) listener(event);
  }
}

async function withSpreadsheetWorker<T>(
  post: WorkerPost,
  run: () => Promise<T>,
): Promise<T> {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
  RecordingSpreadsheetWorker.instance = null;
  RecordingSpreadsheetWorker.post = post;
  Object.defineProperty(globalThis, 'Worker', {
    configurable: true,
    value: RecordingSpreadsheetWorker as unknown as typeof Worker,
  });
  try {
    return await run();
  } finally {
    if (descriptor) Object.defineProperty(globalThis, 'Worker', descriptor);
    else Reflect.deleteProperty(globalThis, 'Worker');
  }
}
