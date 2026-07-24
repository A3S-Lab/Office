import { describe, expect, test } from '@rstest/core';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';
import {
  createSpreadsheetKernelWorkbook,
  spreadsheetCalculationFallbackCells,
  spreadsheetCalculationOps,
  spreadsheetCalculationSessionUpdate,
  spreadsheetCalculationTargets,
} from '../src/internal/features/work/editors/spreadsheet-calculation-model';

describe('Spreadsheet calculation model', () => {
  test('creates a deterministic sparse workbook and ignores presentation-only cells', () => {
    const content = workbook();
    const compiled = createSpreadsheetKernelWorkbook(content);

    expect(compiled?.sheets).toEqual([
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
            row: 0,
            column: 1,
            formula: '=A1*2',
            value: { kind: 'number', value: 4 },
          },
          {
            row: 1,
            column: 0,
            value: { kind: 'text', value: 'North' },
          },
        ],
      },
    ]);

    const changedCache = structuredClone(content);
    const formula = changedCache.sheets[0]?.data?.[0]?.[1];
    if (formula) {
      formula.v = 999;
      formula.m = '999';
    }
    expect(createSpreadsheetKernelWorkbook(changedCache)?.sourceKey).toBe(
      compiled?.sourceKey,
    );

    const changedInput = structuredClone(content);
    const input = changedInput.sheets[0]?.data?.[0]?.[0];
    if (input) input.v = 3;
    expect(createSpreadsheetKernelWorkbook(changedInput)?.sourceKey).not.toBe(
      compiled?.sourceKey,
    );
  });

  test('selects only formulas inside an explicit range', () => {
    const content = workbook();
    content.sheets[0]?.data?.[2]?.splice(0, 1, { f: '=A1+10', v: 12, m: '12' });
    const compiled = createSpreadsheetKernelWorkbook(content);
    if (!compiled) throw new Error('Workbook did not compile.');

    expect(
      spreadsheetCalculationTargets(compiled, {
        scope: 'selection',
        sheetId: 'sheet-1',
        range: { row: [2, 0], column: [2, 0] },
      }),
    ).toEqual([
      { sheetId: 'sheet-1', row: 0, column: 1 },
      { sheetId: 'sheet-1', row: 2, column: 0 },
    ]);
    expect(
      spreadsheetCalculationTargets(compiled, { scope: 'workbook' }),
    ).toBeUndefined();
  });

  test('creates no-history cell patches that preserve formulas and formatting', () => {
    const content = workbook();
    const formula = content.sheets[0]?.data?.[0]?.[1];
    if (!formula) throw new Error('Formula fixture is missing.');
    formula.bg = '#f0f4ff';
    formula.ct = { fa: '0%', t: 'n' };

    const ops = spreadsheetCalculationOps(content.sheets, [
      {
        sheetId: 'sheet-1',
        row: 0,
        column: 1,
        value: { kind: 'number', value: 0.5 },
      },
      {
        sheetId: 'sheet-1',
        row: 2,
        column: 0,
        value: { kind: 'error', value: '#DIV/0!' },
      },
    ]);

    expect(ops).toEqual([
      {
        id: 'sheet-1',
        op: 'replace',
        path: ['data', 0, 1],
        value: {
          bg: '#f0f4ff',
          ct: { fa: '0%', t: 'n' },
          f: '=A1*2',
          m: '50%',
          v: 0.5,
        },
      },
    ]);
  });

  test('requires stable sheet IDs before using the kernel', () => {
    const content = workbook();
    delete content.sheets[0]?.id;
    expect(createSpreadsheetKernelWorkbook(content)).toBeNull();
  });

  test('rejects invalid sparse inputs before posting a Worker request', () => {
    const nonFinite = workbook();
    const numeric = nonFinite.sheets[0]?.data?.[0]?.[0];
    if (numeric) numeric.v = Number.POSITIVE_INFINITY;
    expect(createSpreadsheetKernelWorkbook(nonFinite)).toBeNull();

    const oversizedText = workbook();
    const text = oversizedText.sheets[0]?.data?.[1]?.[0];
    if (text) {
      text.v = 'x'.repeat(32_768);
      text.m = text.v;
    }
    expect(createSpreadsheetKernelWorkbook(oversizedText)).toBeNull();

    const invalidCoordinate = workbook();
    delete invalidCoordinate.sheets[0]?.data;
    invalidCoordinate.sheets[0].celldata = [
      {
        r: 1_048_576,
        c: 0,
        v: { v: 1, m: '1' },
      },
    ];
    expect(createSpreadsheetKernelWorkbook(invalidCoordinate)).toBeNull();
  });

  test('keeps grouped formula ranges on the compatibility path', () => {
    const content = workbook();
    content.sheets[0].formulaMetadata = {
      ranges: [
        {
          type: 'array',
          anchor: 'B1',
          reference: 'B1:B2',
          formula: 'A1*2',
        },
        {
          type: 'data-table',
          anchor: 'A3',
          reference: 'A3:B4',
        },
      ],
    };
    const dataTableRow = content.sheets[0]?.data?.[2];
    if (dataTableRow) {
      dataTableRow[0] = { f: '=A1', v: 2, m: '2' };
    }
    const compiled = createSpreadsheetKernelWorkbook(content);
    if (!compiled) throw new Error('Workbook did not compile.');

    expect(compiled.sheets[0]?.cells.find((cell) => cell.column === 1)).toEqual(
      {
        row: 0,
        column: 1,
        value: { kind: 'number', value: 4 },
      },
    );
    expect(
      spreadsheetCalculationFallbackCells(compiled, { scope: 'workbook' }),
    ).toEqual([
      {
        sheetId: 'sheet-1',
        row: 0,
        column: 1,
        type: 'array',
      },
      {
        sheetId: 'sheet-1',
        row: 2,
        column: 0,
        type: 'data-table',
      },
    ]);
    expect(
      spreadsheetCalculationFallbackCells(
        compiled,
        { scope: 'workbook' },
        false,
      ),
    ).toHaveLength(1);
  });

  test('emits bounded cell patches for stable workbook sessions', () => {
    const previous = createSpreadsheetKernelWorkbook(workbook());
    const changed = workbook();
    const input = changed.sheets[0]?.data?.[0]?.[0];
    const formula = changed.sheets[0]?.data?.[0]?.[1];
    const removed = changed.sheets[0]?.data?.[1];
    if (!input || !formula || !removed) {
      throw new Error('Workbook fixture is incomplete.');
    }
    input.v = 3;
    input.m = '3';
    formula.f = '=A1*3';
    formula.v = 9;
    formula.m = '9';
    removed[0] = null;
    const current = createSpreadsheetKernelWorkbook(changed);
    if (!previous || !current) throw new Error('Workbook did not compile.');

    expect(spreadsheetCalculationSessionUpdate(previous, current, 7)).toEqual({
      kind: 'patch',
      baseDocumentRevision: 7,
      changes: [
        {
          kind: 'upsert',
          sheetId: 'sheet-1',
          row: 0,
          column: 0,
          value: { kind: 'number', value: 3 },
        },
        {
          kind: 'upsert',
          sheetId: 'sheet-1',
          row: 0,
          column: 1,
          formula: '=A1*3',
          value: { kind: 'number', value: 9 },
        },
        {
          kind: 'remove',
          sheetId: 'sheet-1',
          row: 1,
          column: 0,
        },
      ],
    });
  });

  test('does not send formula result caches back to the calculation session', () => {
    const previous = createSpreadsheetKernelWorkbook(workbook());
    const recalculated = workbook();
    const formula = recalculated.sheets[0]?.data?.[0]?.[1];
    if (!formula) throw new Error('Formula fixture is missing.');
    formula.v = 999;
    formula.m = '999';
    const current = createSpreadsheetKernelWorkbook(recalculated);
    if (!previous || !current) throw new Error('Workbook did not compile.');

    expect(spreadsheetCalculationSessionUpdate(previous, current, 4)).toEqual({
      kind: 'patch',
      baseDocumentRevision: 4,
      changes: [],
    });
  });

  test('replaces the session when worksheet structure changes', () => {
    const previous = createSpreadsheetKernelWorkbook(workbook());
    const renamed = workbook();
    renamed.sheets[0].name = 'Renamed';
    const current = createSpreadsheetKernelWorkbook(renamed);
    if (!previous || !current) throw new Error('Workbook did not compile.');

    expect(spreadsheetCalculationSessionUpdate(previous, current, 3)).toEqual({
      kind: 'replace',
      sheets: current.sheets,
    });
  });
});

function workbook(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        data: [
          [{ v: 2, m: '2' }, { f: '=A1*2', v: 4, m: '4' }, { bg: '#f8fafc' }],
          [{ v: 'North', m: 'North' }, null, null],
          [null, null, null],
        ],
      },
    ],
  };
}
