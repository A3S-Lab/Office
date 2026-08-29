import { describe, expect, test } from '@rstest/core';
import { createOfficeKernelClient } from '../src/internal/kernel/office-kernel-client';
import {
  isOfficeKernelResponse,
  OFFICE_KERNEL_PROTOCOL_VERSION,
  type OfficeKernelSpreadsheetCalculatedCell,
  type OfficeKernelSpreadsheetCalculationRequest,
  type OfficeKernelSpreadsheetCoordinate,
  type OfficeKernelSpreadsheetInputSheet,
  type OfficeKernelSpreadsheetInputTable,
} from '../src/internal/kernel/office-kernel-protocol';
import { calculateSpreadsheetInJavaScript } from '../src/internal/kernel/office-kernel-spreadsheet-fallback';
import { JavaScriptSpreadsheetCalculationSession } from '../src/internal/kernel/office-kernel-spreadsheet-session-fallback';
import parityFixtures from './fixtures/spreadsheet-kernel-parity.json';

describe('Spreadsheet calculation kernel', () => {
  test('calculates sparse formula dependencies in deterministic order', async () => {
    const result = await calculateSpreadsheetInJavaScript(request());

    expect(result.engine).toBe('javascript');
    expect(result.cells).toEqual([
      {
        sheetId: 'sheet-1',
        row: 0,
        column: 1,
        value: { kind: 'number', value: 5 },
      },
      {
        sheetId: 'sheet-1',
        row: 1,
        column: 1,
        value: { kind: 'number', value: 10 },
      },
    ]);
    expect(result.calculationOrder).toEqual([
      { sheetId: 'sheet-1', row: 0, column: 1 },
      { sheetId: 'sheet-1', row: 1, column: 1 },
    ]);
    expect(result.issues).toEqual([]);
  });

  test('calculates table structured references in the JavaScript fallback', async () => {
    const table: OfficeKernelSpreadsheetInputTable = {
      name: 'Sales',
      displayName: 'SalesView',
      startRow: 0,
      endRow: 2,
      startColumn: 0,
      endColumn: 2,
      columns: ['Quantity', 'Unit Price', 'Total'],
      headerRow: true,
      totalsRow: false,
    };
    const sheets: OfficeKernelSpreadsheetInputSheet[] = [
      {
        id: 'sales',
        name: 'Sales',
        tables: [table],
        cells: [
          { row: 0, column: 0, value: { kind: 'text', value: 'Quantity' } },
          { row: 0, column: 1, value: { kind: 'text', value: 'Unit Price' } },
          { row: 0, column: 2, value: { kind: 'text', value: 'Total' } },
          { row: 1, column: 0, value: { kind: 'number', value: 2 } },
          { row: 1, column: 1, value: { kind: 'number', value: 10 } },
          {
            row: 1,
            column: 2,
            formula: '=[@Quantity]*[@[Unit Price]]',
            value: { kind: 'blank' },
          },
          { row: 2, column: 0, value: { kind: 'number', value: 3 } },
          { row: 2, column: 1, value: { kind: 'number', value: 20 } },
          {
            row: 2,
            column: 2,
            formula: '=[@Quantity]*[@[Unit Price]]',
            value: { kind: 'blank' },
          },
        ],
      },
      {
        id: 'summary',
        name: 'Summary',
        cells: [
          {
            row: 0,
            column: 0,
            formula: '=SUM(Sales[Quantity])',
            value: { kind: 'blank' },
          },
          {
            row: 1,
            column: 0,
            formula: '=SUM(SalesView[[Quantity]:[Unit Price]])',
            value: { kind: 'blank' },
          },
          {
            row: 2,
            column: 0,
            formula: '=COUNTA(Sales[#Headers])',
            value: { kind: 'blank' },
          },
        ],
      },
    ];
    const result = await calculateSpreadsheetInJavaScript({
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'spreadsheetCalculation',
      requestId: 12,
      revision: 1,
      documentRevision: 1,
      sheets,
    });

    expect(result.cells).toEqual([
      {
        sheetId: 'sales',
        row: 1,
        column: 2,
        value: { kind: 'number', value: 20 },
      },
      {
        sheetId: 'sales',
        row: 2,
        column: 2,
        value: { kind: 'number', value: 60 },
      },
      {
        sheetId: 'summary',
        row: 0,
        column: 0,
        value: { kind: 'number', value: 5 },
      },
      {
        sheetId: 'summary',
        row: 1,
        column: 0,
        value: { kind: 'number', value: 35 },
      },
      {
        sheetId: 'summary',
        row: 2,
        column: 0,
        value: { kind: 'number', value: 3 },
      },
    ]);
    expect(result.issues).toEqual([]);
  });

  test('validates table aliases and keeps oversized structured ranges sparse', async () => {
    const duplicateTable: OfficeKernelSpreadsheetInputTable = {
      name: 'Sales',
      displayName: 'Sales',
      startRow: 0,
      endRow: 0,
      startColumn: 0,
      endColumn: 0,
      columns: ['Value'],
      headerRow: false,
      totalsRow: false,
    };
    const duplicate = request();
    const duplicateSheet = duplicate.sheets[0];
    if (!duplicateSheet) throw new Error('Spreadsheet fixture is incomplete.');
    duplicateSheet.tables = [duplicateTable];
    duplicate.sheets.push({
      id: 'sheet-2',
      name: 'Sheet 2',
      tables: [{ ...duplicateTable }],
      cells: [],
    });
    await expect(
      calculateSpreadsheetInJavaScript(duplicate),
    ).rejects.toMatchObject({
      code: 'office.kernel.spreadsheet.table_invalid',
    });

    const sparse = request();
    const sparseSheet = sparse.sheets[0];
    if (!sparseSheet) throw new Error('Spreadsheet fixture is incomplete.');
    sparseSheet.tables = [
      {
        name: 'Huge',
        startRow: 0,
        endRow: 1_048_575,
        startColumn: 0,
        endColumn: 0,
        columns: ['Value'],
        headerRow: false,
        totalsRow: false,
      },
    ];
    sparseSheet.cells.push({
      row: 0,
      column: 2,
      formula: '=SUM(Huge[Value])',
      value: { kind: 'blank' },
    });
    const result = await calculateSpreadsheetInJavaScript(sparse);
    expect(
      result.cells.some(
        (cell) =>
          cell.sheetId === 'sheet-1' && cell.row === 0 && cell.column === 2,
      ),
    ).toBe(false);
    expect(result.issues).toContainEqual(
      expect.objectContaining({
        cell: { sheetId: 'sheet-1', row: 0, column: 2 },
        code: 'office.kernel.spreadsheet.formula_unsupported',
      }),
    );
  });

  test('fails closed for disjoint structured row areas in the JavaScript fallback', async () => {
    const input = request();
    const sheet = input.sheets[0];
    if (!sheet) throw new Error('Spreadsheet fixture is incomplete.');
    sheet.tables = [
      {
        name: 'Sales',
        startRow: 0,
        endRow: 3,
        startColumn: 0,
        endColumn: 0,
        columns: ['Value'],
        headerRow: true,
        totalsRow: true,
      },
    ];
    sheet.cells.push({
      row: 4,
      column: 0,
      formula: '=SUM(Sales[[#Headers],[#Totals]])',
      value: { kind: 'blank' },
    });

    const result = await calculateSpreadsheetInJavaScript(input);
    expect(
      result.cells.some(
        (cell) =>
          cell.sheetId === 'sheet-1' && cell.row === 4 && cell.column === 0,
      ),
    ).toBe(false);
    expect(result.issues).toContainEqual(
      expect.objectContaining({
        cell: { sheetId: 'sheet-1', row: 4, column: 0 },
        code: 'office.kernel.spreadsheet.formula_unsupported',
        message: expect.stringContaining('disjoint row areas'),
      }),
    );
  });

  test('sorts implicit formula targets independently of sparse input order', async () => {
    const input = request();
    input.sheets[0].cells = [
      {
        row: 4,
        column: 2,
        formula: '=5',
        value: { kind: 'blank' },
      },
      {
        row: 1,
        column: 3,
        formula: '=2',
        value: { kind: 'blank' },
      },
      {
        row: 1,
        column: 1,
        formula: '=1',
        value: { kind: 'blank' },
      },
    ];

    const result = await calculateSpreadsheetInJavaScript(input);

    expect(result.calculationOrder).toEqual([
      { sheetId: 'sheet-1', row: 1, column: 1 },
      { sheetId: 'sheet-1', row: 1, column: 3 },
      { sheetId: 'sheet-1', row: 4, column: 2 },
    ]);
  });

  test('recalculates formulas at the maximum worksheet row without dense allocation', async () => {
    const input = request();
    input.sheets[0].cells = [
      {
        row: 1_048_575,
        column: 0,
        value: { kind: 'number', value: 21 },
      },
      {
        row: 1_048_575,
        column: 1,
        formula: '=A1048576*2',
        value: { kind: 'blank' },
      },
    ];
    input.targets = [{ sheetId: 'sheet-1', row: 1_048_575, column: 1 }];

    const result = await calculateSpreadsheetInJavaScript(input);

    expect(result.cells).toEqual([
      {
        sheetId: 'sheet-1',
        row: 1_048_575,
        column: 1,
        value: { kind: 'number', value: 42 },
      },
    ]);
    expect(result.issues).toEqual([]);
  });

  test('isolates unsupported formulas without discarding valid results', async () => {
    const input = request();
    input.sheets[0]?.cells.push({
      row: 2,
      column: 1,
      formula: '=A3S_UNKNOWN(A1)',
      value: { kind: 'number', value: 41 },
    });

    const result = await calculateSpreadsheetInJavaScript(input);

    expect(result.cells).toHaveLength(2);
    expect(result.issues).toEqual([
      {
        cell: { sheetId: 'sheet-1', row: 2, column: 1 },
        code: 'office.kernel.spreadsheet.formula_unsupported',
        message: "Formula function 'A3S_UNKNOWN' is not supported.",
      },
    ]);
  });

  for (const fixture of parityFixtures) {
    test(`matches the WASM scalar contract for ${fixture.name}`, async () => {
      const input: OfficeKernelSpreadsheetCalculationRequest = {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: 'spreadsheetCalculation',
        requestId: 11,
        revision: 4,
        documentRevision: 9,
        sheets: fixture.sheets as OfficeKernelSpreadsheetInputSheet[],
        targets: fixture.targets as OfficeKernelSpreadsheetCoordinate[],
      };

      const result = await calculateSpreadsheetInJavaScript(input);

      expect(result.cells).toEqual(
        fixture.expectedCells as OfficeKernelSpreadsheetCalculatedCell[],
      );
      expect(result.calculationOrder).toEqual(
        fixture.expectedOrder as OfficeKernelSpreadsheetCoordinate[],
      );
      expect(result.issues.map(({ cell, code }) => ({ cell, code }))).toEqual(
        fixture.expectedIssues,
      );
    });
  }

  test('rejects malformed targets and oversized formulas at the boundary', async () => {
    const malformed = request();
    malformed.targets = [{ sheetId: 'missing', row: 0, column: 0 }];
    await expect(
      calculateSpreadsheetInJavaScript(malformed),
    ).rejects.toMatchObject({
      code: 'office.kernel.spreadsheet.target_invalid',
    });

    const oversized = request();
    const formulaCell = oversized.sheets[0]?.cells.find((cell) => cell.formula);
    if (!formulaCell) throw new Error('Formula fixture is missing.');
    formulaCell.formula = `=${'1'.repeat(8_192)}`;
    await expect(
      calculateSpreadsheetInJavaScript(oversized),
    ).rejects.toMatchObject({
      code: 'office.kernel.spreadsheet.formula_invalid',
    });

    const invalidRevision = request();
    invalidRevision.documentRevision = -1;
    await expect(
      calculateSpreadsheetInJavaScript(invalidRevision),
    ).rejects.toMatchObject({
      code: 'office.kernel.revision_invalid',
    });

    const oversizedRequestId = request();
    oversizedRequestId.requestId = 2 ** 32;
    await expect(
      calculateSpreadsheetInJavaScript(oversizedRequestId),
    ).rejects.toMatchObject({
      code: 'office.kernel.revision_invalid',
    });

    const oversizedTargets = request();
    oversizedTargets.targets = Array.from({ length: 100_001 }, () => ({
      sheetId: 'sheet-1',
      row: 0,
      column: 0,
    }));
    await expect(
      calculateSpreadsheetInJavaScript(oversizedTargets),
    ).rejects.toMatchObject({
      code: 'office.kernel.spreadsheet.target_limit_exceeded',
    });
  });

  test('bounds deeply nested dependency chains before the JavaScript stack', async () => {
    const chainLength = 300;
    const input = request();
    input.sheets[0].cells = Array.from({ length: chainLength }, (_, row) => ({
      row,
      column: 0,
      formula: row === 0 ? '=1' : `=A${row}+1`,
      value: { kind: 'number' as const, value: row + 1 },
    }));
    input.targets = [{ sheetId: 'sheet-1', row: chainLength - 1, column: 0 }];

    const result = await calculateSpreadsheetInJavaScript(input);

    expect(result.cells).toEqual([]);
    expect(
      result.issues.some(
        (issue) =>
          issue.code === 'office.kernel.spreadsheet.dependency_depth_exceeded',
      ),
    ).toBe(true);
  });

  test('bounds text produced by fallback formula calculation', async () => {
    const input = request();
    input.sheets[0].cells = [
      {
        row: 0,
        column: 0,
        value: { kind: 'text', value: 'a'.repeat(20_000) },
      },
      {
        row: 0,
        column: 1,
        value: { kind: 'text', value: 'b'.repeat(20_000) },
      },
      {
        row: 0,
        column: 2,
        formula: '=CONCAT(A1,B1)',
        value: { kind: 'blank' },
      },
    ];

    const result = await calculateSpreadsheetInJavaScript(input);

    expect(result.cells).toEqual([
      {
        sheetId: 'sheet-1',
        row: 0,
        column: 2,
        value: { kind: 'error', value: '#VALUE!' },
      },
    ]);
  });

  test('validates result coordinates and scalar values at the Worker boundary', () => {
    expect(
      isOfficeKernelResponse({
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: 'spreadsheetCalculationResult',
        requestId: 7,
        revision: 3,
        documentRevision: 2,
        engine: 'wasm',
        cells: [
          {
            sheetId: 'sheet-1',
            row: 0,
            column: 1,
            value: { kind: 'number', value: 5 },
          },
        ],
        calculationOrder: [{ sheetId: 'sheet-1', row: 0, column: 1 }],
        issues: [],
      }),
    ).toBe(true);

    expect(
      isOfficeKernelResponse({
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: 'spreadsheetCalculationResult',
        requestId: 7,
        revision: 3,
        documentRevision: 2,
        engine: 'wasm',
        cells: [
          {
            sheetId: '',
            row: -1,
            column: 1,
            value: { kind: 'number', value: Number.NaN },
          },
        ],
        calculationOrder: [],
        issues: [],
      }),
    ).toBe(false);

    expect(
      isOfficeKernelResponse({
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: 'spreadsheetCalculationResult',
        requestId: 7,
        revision: 3,
        documentRevision: 2,
        engine: 'wasm',
        cells: [
          {
            sheetId: 'sheet-1',
            row: 0,
            column: 1,
            value: { kind: 'text', value: 'x'.repeat(32_768) },
          },
        ],
        calculationOrder: [{ sheetId: 'sheet-1', row: 0, column: 1 }],
        issues: [],
      }),
    ).toBe(false);
  });

  test('validates persistent session statistics at the Worker boundary', () => {
    const response = {
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'spreadsheetSessionCalculationResult',
      requestId: 7,
      revision: 3,
      documentRevision: 2,
      engine: 'wasm',
      cells: [],
      calculationOrder: [],
      issues: [],
      stats: {
        updateKind: 'patch',
        calculationScope: 'dirty',
        formulaCellCount: 10,
        dirtyFormulaCellCount: 2,
        evaluatedFormulaCellCount: 2,
        reusedFormulaCellCount: 1,
        dependencyEdgeCount: 12,
      },
    };

    expect(isOfficeKernelResponse(response)).toBe(true);
    expect(
      isOfficeKernelResponse({
        ...response,
        stats: { ...response.stats, dependencyEdgeCount: 1_000_001 },
      }),
    ).toBe(false);
  });

  test('uses the JavaScript calculation boundary when Worker is unavailable', async () => {
    const workerDescriptor = Object.getOwnPropertyDescriptor(
      globalThis,
      'Worker',
    );
    Object.defineProperty(globalThis, 'Worker', {
      configurable: true,
      value: undefined,
    });
    const client = createOfficeKernelClient();

    try {
      const result = await client.spreadsheetCalculation({
        revision: 3,
        documentRevision: 2,
        sheets: request().sheets,
      });
      expect(result.engine).toBe('javascript');
      expect(result.cells.at(-1)?.value).toEqual({
        kind: 'number',
        value: 10,
      });
    } finally {
      client.dispose();
      if (workerDescriptor) {
        Object.defineProperty(globalThis, 'Worker', workerDescriptor);
      } else {
        Reflect.deleteProperty(globalThis, 'Worker');
      }
    }
  });

  test('keeps session requests recoverable when Worker is unavailable', async () => {
    const workerDescriptor = Object.getOwnPropertyDescriptor(
      globalThis,
      'Worker',
    );
    Object.defineProperty(globalThis, 'Worker', {
      configurable: true,
      value: undefined,
    });
    const client = createOfficeKernelClient();
    const sheets = request().sheets;

    try {
      const initial = await client.spreadsheetSessionCalculation({
        revision: 1,
        documentRevision: 1,
        update: { kind: 'replace', sheets },
        calculation: { kind: 'workbook' },
        fallbackSheets: sheets,
      });
      expect(initial.engine).toBe('javascript');
      expect(initial.stats).toMatchObject({
        updateKind: 'replace',
        calculationScope: 'workbook',
        formulaCellCount: 2,
      });

      const changed = structuredClone(sheets);
      const input = changed[0]?.cells.find(
        (cell) => cell.row === 0 && cell.column === 0,
      );
      if (!input) throw new Error('Spreadsheet input fixture is missing.');
      input.value = { kind: 'number', value: 4 };
      const patched = await client.spreadsheetSessionCalculation({
        revision: 2,
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
              value: { kind: 'number', value: 4 },
            },
          ],
        },
        calculation: { kind: 'dirty' },
        fallbackSheets: changed,
      });
      expect(patched.cells.at(-1)?.value).toEqual({
        kind: 'number',
        value: 14,
      });
    } finally {
      client.dispose();
      if (workerDescriptor) {
        Object.defineProperty(globalThis, 'Worker', workerDescriptor);
      } else {
        Reflect.deleteProperty(globalThis, 'Worker');
      }
    }
  });

  test('keeps the Worker JavaScript session revisioned and atomic', async () => {
    const session = new JavaScriptSpreadsheetCalculationSession();
    const sheets = request().sheets;
    await session.calculate({
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'spreadsheetSessionCalculation',
      requestId: 1,
      revision: 1,
      documentRevision: 1,
      update: { kind: 'replace', sheets },
      calculation: { kind: 'workbook' },
    });

    await expect(
      session.calculate({
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: 'spreadsheetSessionCalculation',
        requestId: 2,
        revision: 2,
        documentRevision: 2,
        update: {
          kind: 'patch',
          baseDocumentRevision: 0,
          changes: [
            {
              kind: 'upsert',
              sheetId: 'sheet-1',
              row: 0,
              column: 0,
              value: { kind: 'number', value: 99 },
            },
          ],
        },
        calculation: { kind: 'dirty' },
      }),
    ).rejects.toMatchObject({
      code: 'office.kernel.spreadsheet.session_revision_mismatch',
    });

    await expect(
      session.calculate({
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: 'spreadsheetSessionCalculation',
        requestId: 3,
        revision: 3,
        documentRevision: 0,
        update: {
          kind: 'patch',
          baseDocumentRevision: 1,
          changes: [],
        },
        calculation: { kind: 'dirty' },
      }),
    ).rejects.toMatchObject({
      code: 'office.kernel.spreadsheet.session_revision_invalid',
    });

    const result = await session.calculate({
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'spreadsheetSessionCalculation',
      requestId: 4,
      revision: 4,
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
            value: { kind: 'number', value: 4 },
          },
        ],
      },
      calculation: { kind: 'dirty' },
    });
    expect(result.cells.at(-1)?.value).toEqual({
      kind: 'number',
      value: 14,
    });
  });

  test('rejects an already-cancelled fallback calculation', async () => {
    const workerDescriptor = Object.getOwnPropertyDescriptor(
      globalThis,
      'Worker',
    );
    Object.defineProperty(globalThis, 'Worker', {
      configurable: true,
      value: undefined,
    });
    const client = createOfficeKernelClient();
    const controller = new AbortController();
    controller.abort();

    try {
      await expect(
        client.spreadsheetCalculation(
          {
            revision: 3,
            documentRevision: 2,
            sheets: request().sheets,
          },
          controller.signal,
        ),
      ).rejects.toMatchObject({ name: 'AbortError' });
    } finally {
      client.dispose();
      if (workerDescriptor) {
        Object.defineProperty(globalThis, 'Worker', workerDescriptor);
      } else {
        Reflect.deleteProperty(globalThis, 'Worker');
      }
    }
  });
});

function request(): OfficeKernelSpreadsheetCalculationRequest {
  return {
    protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
    kind: 'spreadsheetCalculation',
    requestId: 7,
    revision: 3,
    documentRevision: 2,
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        cells: [
          {
            row: 0,
            column: 0,
            value: { kind: 'number', value: 2 },
          },
          {
            row: 1,
            column: 0,
            value: { kind: 'number', value: 3 },
          },
          {
            row: 0,
            column: 1,
            formula: '=SUM(A1:A2)',
            value: { kind: 'blank' },
          },
          {
            row: 1,
            column: 1,
            formula: '=B1*2',
            value: { kind: 'blank' },
          },
        ],
      },
    ],
  };
}
