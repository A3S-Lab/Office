import type { Op, Sheet } from '@fortune-sheet/core';
import type { WorkbookInstance } from '@fortune-sheet/react';
import { describe, expect, test } from '@rstest/core';
import { act, renderHook, waitFor } from '@testing-library/react';
import { useSpreadsheetCalculation } from '../src/internal/features/work/editors/use-spreadsheet-calculation';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';
import type {
  OfficeKernelSpreadsheetSessionCalculationRequest,
  OfficeKernelWorkerRequest,
  OfficeKernelWorkerResponse,
} from '../src/internal/kernel/office-kernel-protocol';

describe('Spreadsheet calculation hook', () => {
  test('reuses one Worker session and posts only changed cells after initialization', async () => {
    const workerDescriptor = Object.getOwnPropertyDescriptor(
      globalThis,
      'Worker',
    );
    RecordingKernelWorker.reset();
    Object.defineProperty(globalThis, 'Worker', {
      configurable: true,
      value: RecordingKernelWorker as unknown as typeof Worker,
    });
    const initial = content(2);
    const workbook = new RecordingWorkbook(initial.sheets);
    const workbookRef = { current: workbook as unknown as WorkbookInstance };

    try {
      const { rerender, unmount } = renderHook(
        ({ value }) =>
          useSpreadsheetCalculation({
            content: value,
            workbookRef,
          }),
        { initialProps: { value: initial } },
      );
      await waitFor(() =>
        expect(RecordingKernelWorker.sessionRequests).toHaveLength(1),
      );
      await waitFor(() =>
        expect(RecordingKernelWorker.completedResponses).toBe(1),
      );
      const initialOperationCount = workbook.ops.length;

      const latest = content(3);
      workbook.sheets = latest.sheets;
      rerender({ value: latest });
      await waitFor(() =>
        expect(RecordingKernelWorker.sessionRequests).toHaveLength(2),
      );
      await waitFor(() =>
        expect(workbook.ops).toHaveLength(initialOperationCount + 1),
      );

      expect(RecordingKernelWorker.sessionRequests[0]).toMatchObject({
        update: { kind: 'replace' },
        calculation: { kind: 'workbook' },
      });
      expect(RecordingKernelWorker.sessionRequests[1]).toMatchObject({
        update: {
          kind: 'patch',
          baseDocumentRevision: 1,
          changes: [
            {
              kind: 'upsert',
              sheetId: 'sheet-1',
              row: 0,
              column: 0,
              value: { kind: 'number', value: 3 },
            },
          ],
        },
        calculation: { kind: 'dirty' },
      });
      expect('sheets' in RecordingKernelWorker.sessionRequests[1].update).toBe(
        false,
      );
      expect(workbook.ops.at(-1)?.[0]).toMatchObject({
        value: { f: '=A1*2', m: '6', v: 6 },
      });
      unmount();
    } finally {
      if (workerDescriptor) {
        Object.defineProperty(globalThis, 'Worker', workerDescriptor);
      } else {
        Reflect.deleteProperty(globalThis, 'Worker');
      }
    }
  });

  test('keeps session revisions aligned when an earlier Worker request is cancelled', async () => {
    const workerDescriptor = Object.getOwnPropertyDescriptor(
      globalThis,
      'Worker',
    );
    DeferredKernelWorker.reset();
    Object.defineProperty(globalThis, 'Worker', {
      configurable: true,
      value: DeferredKernelWorker as unknown as typeof Worker,
    });
    const initial = content(2);
    const workbook = new RecordingWorkbook(initial.sheets);
    const workbookRef = { current: workbook as unknown as WorkbookInstance };

    try {
      const { rerender, unmount } = renderHook(
        ({ value }) =>
          useSpreadsheetCalculation({
            content: value,
            workbookRef,
          }),
        { initialProps: { value: initial } },
      );
      await waitFor(() =>
        expect(DeferredKernelWorker.sessionRequests).toHaveLength(1),
      );

      const latest = content(3);
      workbook.sheets = latest.sheets;
      rerender({ value: latest });
      await waitFor(() =>
        expect(DeferredKernelWorker.sessionRequests).toHaveLength(2),
      );

      const [initialRequest, patchedRequest] =
        DeferredKernelWorker.sessionRequests;
      expect(DeferredKernelWorker.cancelledRequestIds).toContain(
        initialRequest?.requestId,
      );
      expect(patchedRequest).toMatchObject({
        documentRevision: 2,
        update: {
          kind: 'patch',
          baseDocumentRevision: 1,
          changes: [
            {
              kind: 'upsert',
              sheetId: 'sheet-1',
              row: 0,
              column: 0,
              value: { kind: 'number', value: 3 },
            },
          ],
        },
        calculation: { kind: 'dirty' },
      });

      act(() => DeferredKernelWorker.flush());
      await waitFor(() => expect(workbook.ops).toHaveLength(1));
      expect(DeferredKernelWorker.revisionErrors).toEqual([]);
      expect(workbook.ops[0]?.[0]).toMatchObject({
        value: { f: '=A1*2', m: '6', v: 6 },
      });
      unmount();
    } finally {
      if (workerDescriptor) {
        Object.defineProperty(globalThis, 'Worker', workerDescriptor);
      } else {
        Reflect.deleteProperty(globalThis, 'Worker');
      }
    }
  });

  test('calculates the latest controlled workbook and applies no-history patches', async () => {
    const workbook = new RecordingWorkbook(content(2).sheets);
    const workbookRef = { current: workbook as unknown as WorkbookInstance };
    const initial = content(2);
    const { rerender, unmount } = renderHook(
      ({ value }) =>
        useSpreadsheetCalculation({
          content: value,
          workbookRef,
        }),
      { initialProps: { value: initial } },
    );

    const latest = content(3);
    workbook.sheets = latest.sheets;
    rerender({ value: latest });

    await waitFor(() => expect(workbook.ops).toHaveLength(1));
    expect(workbook.ops[0]?.[0]).toMatchObject({
      id: 'sheet-1',
      op: 'replace',
      path: ['data', 0, 1],
      value: { f: '=A1*2', m: '6', v: 6 },
    });
    unmount();
  });

  test('honors manual mode and exposes explicit selection recalculation', async () => {
    const value = content(4);
    value.calculation = {
      mode: 'manual',
      fullCalculationOnLoad: false,
      forceFullCalculation: false,
      iterativeCalculation: false,
      maximumIterations: 100,
      maximumChange: 0.001,
      fullPrecision: true,
    };
    const workbook = new RecordingWorkbook(value.sheets);
    const workbookRef = { current: workbook as unknown as WorkbookInstance };
    const { result, unmount } = renderHook(() =>
      useSpreadsheetCalculation({ content: value, workbookRef }),
    );

    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(workbook.ops).toEqual([]);

    act(() => {
      result.current.recalculate({
        scope: 'selection',
        sheetId: 'sheet-1',
        range: { row: [0, 0], column: [1, 1] },
      });
    });

    await waitFor(() => expect(workbook.ops).toHaveLength(1));
    expect(workbook.ops[0]?.[0]).toMatchObject({
      value: { f: '=A1*2', m: '8', v: 8 },
    });
    unmount();
  });

  test('falls back only the unsupported formula cell', async () => {
    const value = content(2);
    const formula = value.sheets[0]?.data?.[0]?.[1];
    if (!formula) throw new Error('Formula fixture is missing.');
    formula.f = '=A3S_UNKNOWN(A1)';
    formula.v = 41;
    formula.m = '41';
    const workbook = new RecordingWorkbook(value.sheets);
    const workbookRef = { current: workbook as unknown as WorkbookInstance };
    const { unmount } = renderHook(() =>
      useSpreadsheetCalculation({ content: value, workbookRef }),
    );

    await waitFor(() =>
      expect(workbook.fallbacks).toEqual([
        {
          sheetId: 'sheet-1',
          range: { row: [0, 0], column: [1, 1] },
        },
      ]),
    );
    expect(workbook.ops).toEqual([]);
    unmount();
  });

  test('falls back unresolved dependents in dependency order', async () => {
    const value = content(2);
    const input = value.sheets[0]?.data?.[0]?.[0];
    const dependent = value.sheets[0]?.data?.[0]?.[1];
    if (!input || !dependent) throw new Error('Formula fixture is missing.');
    input.f = '=A3S_UNKNOWN(1)';
    input.v = 41;
    input.m = '41';
    dependent.f = '=A1+1';
    dependent.v = 42;
    dependent.m = '42';
    const workbook = new RecordingWorkbook(value.sheets);
    const workbookRef = { current: workbook as unknown as WorkbookInstance };
    const { unmount } = renderHook(() =>
      useSpreadsheetCalculation({ content: value, workbookRef }),
    );

    await waitFor(() =>
      expect(workbook.fallbacks).toEqual([
        {
          sheetId: 'sheet-1',
          range: { row: [0, 0], column: [0, 0] },
        },
        {
          sheetId: 'sheet-1',
          range: { row: [0, 0], column: [1, 1] },
        },
      ]),
    );
    expect(workbook.ops).toEqual([]);
    unmount();
  });

  test('refreshes grouped formulas before calculating their dependents', async () => {
    const value = content(3);
    const row = value.sheets[0]?.data?.[0];
    if (!row) throw new Error('Formula fixture is missing.');
    row[1] = { f: '=A1*2', v: 4, m: '4' };
    row[2] = { f: '=B1+1', v: 5, m: '5' };
    value.sheets[0].formulaMetadata = {
      ranges: [
        {
          type: 'array',
          anchor: 'B1',
          reference: 'B1',
          formula: 'A1*2',
        },
      ],
    };
    const workbook = new RecordingWorkbook(value.sheets);
    workbook.onCalculateFormula = (_sheetId, range) => {
      if (range?.column[0] !== 1) return;
      const grouped = workbook.sheets[0]?.data?.[0]?.[1];
      if (!grouped) return;
      grouped.v = 6;
      grouped.m = '6';
    };
    const workbookRef = { current: workbook as unknown as WorkbookInstance };
    const { unmount } = renderHook(() =>
      useSpreadsheetCalculation({ content: value, workbookRef }),
    );

    await waitFor(() => expect(workbook.ops).toHaveLength(1));
    expect(workbook.fallbacks).toEqual([
      {
        sheetId: 'sheet-1',
        range: { row: [0, 0], column: [1, 1] },
      },
    ]);
    expect(workbook.ops[0]?.[0]).toMatchObject({
      path: ['data', 0, 2],
      value: { f: '=B1+1', m: '7', v: 7 },
    });
    unmount();
  });

  test('leaves data tables unchanged in automatic-except-data-tables mode', async () => {
    const value = content(2);
    value.calculation = {
      mode: 'automatic-except-data-tables',
      fullCalculationOnLoad: false,
      forceFullCalculation: false,
      iterativeCalculation: false,
      maximumIterations: 100,
      maximumChange: 0.001,
      fullPrecision: true,
    };
    value.sheets[0].formulaMetadata = {
      ranges: [
        {
          type: 'data-table',
          anchor: 'B1',
          reference: 'B1:B2',
        },
      ],
    };
    const workbook = new RecordingWorkbook(value.sheets);
    const workbookRef = { current: workbook as unknown as WorkbookInstance };
    const { result, unmount } = renderHook(() =>
      useSpreadsheetCalculation({ content: value, workbookRef }),
    );

    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(workbook.ops).toEqual([]);
    expect(workbook.fallbacks).toEqual([]);

    act(() => result.current.recalculate({ scope: 'workbook' }));
    await waitFor(() => expect(workbook.fallbacks).toHaveLength(1));
    expect(workbook.fallbacks[0]).toEqual({
      sheetId: 'sheet-1',
      range: { row: [0, 0], column: [1, 1] },
    });
    unmount();
  });
});

class RecordingWorkbook {
  fallbacks: Array<{
    sheetId: string | undefined;
    range: { row: number[]; column: number[] } | undefined;
  }> = [];
  ops: Op[][] = [];
  onCalculateFormula?: (
    sheetId?: string,
    range?: { row: number[]; column: number[] },
  ) => void;

  constructor(public sheets: Sheet[]) {}

  applyOp(ops: Op[]): void {
    this.ops.push(ops);
  }

  calculateFormula(
    sheetId?: string,
    range?: { row: number[]; column: number[] },
  ): void {
    this.fallbacks.push({ sheetId, range });
    this.onCalculateFormula?.(sheetId, range);
  }

  getAllSheets(): Sheet[] {
    return this.sheets;
  }
}

class RecordingKernelWorker {
  static completedResponses = 0;
  static sessionRequests: OfficeKernelSpreadsheetSessionCalculationRequest[] =
    [];

  static reset(): void {
    RecordingKernelWorker.completedResponses = 0;
    RecordingKernelWorker.sessionRequests = [];
  }

  private input = 0;
  private listeners = new Set<
    (event: MessageEvent<OfficeKernelWorkerResponse>) => void
  >();

  addEventListener(
    type: string,
    listener: (event: MessageEvent<OfficeKernelWorkerResponse>) => void,
  ): void {
    if (type === 'message') this.listeners.add(listener);
  }

  removeEventListener(
    type: string,
    listener: (event: MessageEvent<OfficeKernelWorkerResponse>) => void,
  ): void {
    if (type === 'message') this.listeners.delete(listener);
  }

  postMessage(message: OfficeKernelWorkerRequest): void {
    if (message.kind !== 'spreadsheetSessionCalculation') return;
    const request = message.request;
    RecordingKernelWorker.sessionRequests.push(request);
    if (request.update.kind === 'replace') {
      this.input = numberValue(
        request.update.sheets[0]?.cells.find(
          (cell) => cell.row === 0 && cell.column === 0,
        )?.value,
      );
    } else {
      const change = request.update.changes.find(
        (candidate) =>
          candidate.kind === 'upsert' &&
          candidate.row === 0 &&
          candidate.column === 0,
      );
      if (change?.kind === 'upsert') {
        this.input = numberValue(change.value);
      }
    }
    queueMicrotask(() => {
      const event = new MessageEvent<OfficeKernelWorkerResponse>('message', {
        data: {
          kind: 'response',
          response: {
            protocol: request.protocol,
            kind: 'spreadsheetSessionCalculationResult',
            requestId: request.requestId,
            revision: request.revision,
            documentRevision: request.documentRevision,
            engine: 'wasm',
            cells: [
              {
                sheetId: 'sheet-1',
                row: 0,
                column: 1,
                value: { kind: 'number', value: this.input * 2 },
              },
            ],
            calculationOrder: [{ sheetId: 'sheet-1', row: 0, column: 1 }],
            issues: [],
            stats: {
              updateKind: request.update.kind,
              calculationScope: request.calculation.kind,
              formulaCellCount: 1,
              dirtyFormulaCellCount: 1,
              evaluatedFormulaCellCount: 1,
              reusedFormulaCellCount: 0,
              dependencyEdgeCount: 1,
            },
          },
        },
      });
      for (const listener of this.listeners) listener(event);
      queueMicrotask(() => {
        RecordingKernelWorker.completedResponses += 1;
      });
    });
  }

  terminate(): void {}
}

class DeferredKernelWorker {
  static cancelledRequestIds: number[] = [];
  static instance: DeferredKernelWorker | null = null;
  static revisionErrors: string[] = [];
  static sessionRequests: OfficeKernelSpreadsheetSessionCalculationRequest[] =
    [];

  static reset(): void {
    DeferredKernelWorker.cancelledRequestIds = [];
    DeferredKernelWorker.instance = null;
    DeferredKernelWorker.revisionErrors = [];
    DeferredKernelWorker.sessionRequests = [];
  }

  static flush(): void {
    DeferredKernelWorker.instance?.flush();
  }

  private cancelled = new Set<number>();
  private documentRevision = 0;
  private input = 0;
  private listeners = new Set<
    (event: MessageEvent<OfficeKernelWorkerResponse>) => void
  >();
  private queue: OfficeKernelSpreadsheetSessionCalculationRequest[] = [];

  constructor() {
    DeferredKernelWorker.instance = this;
  }

  addEventListener(
    type: string,
    listener: (event: MessageEvent<OfficeKernelWorkerResponse>) => void,
  ): void {
    if (type === 'message') this.listeners.add(listener);
  }

  removeEventListener(
    type: string,
    listener: (event: MessageEvent<OfficeKernelWorkerResponse>) => void,
  ): void {
    if (type === 'message') this.listeners.delete(listener);
  }

  postMessage(message: OfficeKernelWorkerRequest): void {
    if (message.kind === 'cancel') {
      this.cancelled.add(message.requestId);
      DeferredKernelWorker.cancelledRequestIds.push(message.requestId);
      return;
    }
    if (message.kind !== 'spreadsheetSessionCalculation') return;
    DeferredKernelWorker.sessionRequests.push(message.request);
    this.queue.push(message.request);
  }

  terminate(): void {}

  private flush(): void {
    for (const request of this.queue.splice(0)) {
      if (request.update.kind === 'replace') {
        this.input = numberValue(
          request.update.sheets[0]?.cells.find(
            (cell) => cell.row === 0 && cell.column === 0,
          )?.value,
        );
      } else {
        if (request.update.baseDocumentRevision !== this.documentRevision) {
          DeferredKernelWorker.revisionErrors.push(
            `${request.update.baseDocumentRevision}:${this.documentRevision}`,
          );
          continue;
        }
        const change = request.update.changes.find(
          (candidate) =>
            candidate.kind === 'upsert' &&
            candidate.row === 0 &&
            candidate.column === 0,
        );
        if (change?.kind === 'upsert') {
          this.input = numberValue(change.value);
        }
      }
      this.documentRevision = request.documentRevision;
      if (this.cancelled.has(request.requestId)) continue;
      const event = new MessageEvent<OfficeKernelWorkerResponse>('message', {
        data: {
          kind: 'response',
          response: sessionResult(request, this.input),
        },
      });
      for (const listener of this.listeners) listener(event);
    }
  }
}

function sessionResult(
  request: OfficeKernelSpreadsheetSessionCalculationRequest,
  input: number,
): Extract<
  OfficeKernelWorkerResponse['response'],
  { kind: 'spreadsheetSessionCalculationResult' }
> {
  return {
    protocol: request.protocol,
    kind: 'spreadsheetSessionCalculationResult',
    requestId: request.requestId,
    revision: request.revision,
    documentRevision: request.documentRevision,
    engine: 'wasm',
    cells: [
      {
        sheetId: 'sheet-1',
        row: 0,
        column: 1,
        value: { kind: 'number', value: input * 2 },
      },
    ],
    calculationOrder: [{ sheetId: 'sheet-1', row: 0, column: 1 }],
    issues: [],
    stats: {
      updateKind: request.update.kind,
      calculationScope: request.calculation.kind,
      formulaCellCount: 1,
      dirtyFormulaCellCount: 1,
      evaluatedFormulaCellCount: 1,
      reusedFormulaCellCount: 0,
      dependencyEdgeCount: 1,
    },
  };
}

function numberValue(
  value:
    | {
        kind: string;
        value?: unknown;
      }
    | undefined,
): number {
  return value?.kind === 'number' && typeof value.value === 'number'
    ? value.value
    : 0;
}

function content(input: number): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        data: [
          [
            { v: input, m: String(input) },
            { f: '=A1*2', v: input * 2, m: String(input * 2) },
          ],
        ],
      },
    ],
  };
}
