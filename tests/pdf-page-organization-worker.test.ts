import { afterEach, expect, test } from '@rstest/core';
import { PDFDocument } from 'pdf-lib';
import { runPdfPageOrganizationJob } from '../src/internal/features/work/editors/pdf-page-organization-worker-client';

const originalWorker = globalThis.Worker;

afterEach(() => {
  Object.defineProperty(globalThis, 'Worker', {
    configurable: true,
    value: originalWorker,
    writable: true,
  });
});

test('transfers a bounded page job to a dedicated worker and closes its exact worker', async () => {
  const source = await createPdf(1);
  const imported = await createPdf(2);
  let instance: MockWorker | undefined;
  Object.defineProperty(globalThis, 'Worker', {
    configurable: true,
    value: class extends MockWorker {
      constructor(url: URL, options: WorkerOptions) {
        super(url, options);
        instance = this;
      }
    },
    writable: true,
  });

  const pending = runPdfPageOrganizationJob({
    importSource: imported.buffer,
    kind: 'mutate',
    mutation: { index: 1, kind: 'merge' },
    source: source.buffer,
  });
  expect(instance?.name).toBe('a3s-office-pdf-page-organization');
  expect(instance?.transferred).toHaveLength(2);
  instance?.emitMessage({
    kind: 'success',
    result: {
      diagnostics: [],
      kind: 'mutated',
      pageCount: 3,
      source: new Uint8Array([37, 80, 68, 70]),
    },
  });
  await expect(pending).resolves.toMatchObject({
    kind: 'mutated',
    pageCount: 3,
  });
  expect(instance?.terminated).toBe(true);
});

test('reconstructs stable typed worker failures', async () => {
  const source = await createPdf(1);
  let instance: MockWorker | undefined;
  Object.defineProperty(globalThis, 'Worker', {
    configurable: true,
    value: class extends MockWorker {
      constructor(url: URL, options: WorkerOptions) {
        super(url, options);
        instance = this;
      }
    },
    writable: true,
  });
  const pending = runPdfPageOrganizationJob({
    kind: 'mutate',
    mutation: { kind: 'delete', pageIndexes: [0] },
    source: source.buffer,
  });
  instance?.emitMessage({
    code: 'pdf.pages.empty-document',
    kind: 'failure',
    message: 'A PDF must retain at least one page.',
  });
  await expect(pending).rejects.toMatchObject({
    code: 'pdf.pages.empty-document',
    message: 'A PDF must retain at least one page.',
    name: 'PdfPageOrganizationError',
  });
  expect(instance?.terminated).toBe(true);
});

test('aborts only the owned page worker', async () => {
  const source = await createPdf(1);
  let instance: MockWorker | undefined;
  Object.defineProperty(globalThis, 'Worker', {
    configurable: true,
    value: class extends MockWorker {
      constructor(url: URL, options: WorkerOptions) {
        super(url, options);
        instance = this;
      }
    },
    writable: true,
  });
  const controller = new AbortController();
  const pending = runPdfPageOrganizationJob(
    {
      kind: 'mutate',
      mutation: { degrees: 90, kind: 'rotate', pageIndexes: [0] },
      source: source.buffer,
    },
    controller.signal,
  );
  controller.abort();
  await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
  expect(instance?.terminated).toBe(true);
});

class MockWorker extends EventTarget {
  readonly name: string | undefined;
  terminated = false;
  transferred: Transferable[] = [];

  constructor(_url: URL, options: WorkerOptions) {
    super();
    this.name = options.name;
  }

  postMessage(_message: unknown, options?: StructuredSerializeOptions): void {
    this.transferred = options?.transfer ? [...options.transfer] : [];
  }

  terminate(): void {
    this.terminated = true;
  }

  emitMessage(data: unknown): void {
    this.dispatchEvent(new MessageEvent('message', { data }));
  }
}

async function createPdf(pageCount: number): Promise<Uint8Array> {
  const pdf = await PDFDocument.create();
  for (let index = 0; index < pageCount; index += 1) pdf.addPage([300, 600]);
  return pdf.save();
}
