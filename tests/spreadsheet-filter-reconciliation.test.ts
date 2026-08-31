import type { Op } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import {
  reconcileSpreadsheetFiltersAfterFortune,
  reconcileSpreadsheetFiltersAfterSort,
} from '../src/internal/features/work/editors/spreadsheet-filter-reconciliation';
import { workSpreadsheetSheetWithAutoFilterCriteria } from '../src/internal/features/work/work-spreadsheet-auto-filter';
import type { WorkSpreadsheetSheet } from '../src/internal/features/work/work-types';

test('recomputes AutoFilter-owned rows after a sort and preserves manual hiding', () => {
  const source = workSpreadsheetSheetWithAutoFilterCriteria(
    {
      id: 'sheet-1',
      name: 'Filtered',
      data: [
        [{ v: 'Status' }],
        [{ v: 'Keep' }],
        [{ v: 'Drop' }],
        [{ v: 'Keep' }],
      ],
      filter_select: { row: [0, 3], column: [0, 0] },
      config: { rowhidden: { '3': 0 } },
    },
    0,
    { type: 'equals', value: 'Keep' },
  );
  if (!source) throw new Error('Expected an AutoFilter fixture.');
  const changed: WorkSpreadsheetSheet = {
    ...source,
    data: [
      [{ v: 'Status' }],
      [{ v: 'Drop' }],
      [{ v: 'Keep' }],
      [{ v: 'Keep' }],
    ],
  };

  const [reconciled] = reconcileSpreadsheetFiltersAfterFortune(
    [changed],
    [source],
    cellOperations('sheet-1', [1, 2]),
  );

  expect(reconciled?.config?.rowhidden).toEqual({ '1': 0, '3': 0 });
  expect(reconciled?.filter?.['0']).toMatchObject({ rowhidden: { '1': 0 } });
});

test('recomputes table filters without moving the totals row or losing manual hiding', () => {
  const source: WorkSpreadsheetSheet = {
    id: 'sheet-1',
    name: 'Table',
    data: [
      [{ v: 'Status' }],
      [{ v: 'Keep' }],
      [{ v: 'Drop' }],
      [{ v: 'Keep' }],
      [{ v: 'Total' }],
    ],
    config: { rowhidden: { '2': 0, '3': 0 } },
    tables: [
      {
        id: 'table-1',
        name: 'Table1',
        range: { row: [0, 4], column: [0, 0] },
        columns: [{ name: 'Status' }],
        filters: [{ column: 0, criteria: { type: 'equals', value: 'Keep' } }],
        headerRow: true,
        totalsRow: true,
        style: { family: 'medium', number: 2 },
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    ],
  };
  const changed: WorkSpreadsheetSheet = {
    ...source,
    data: [
      [{ v: 'Status' }],
      [{ v: 'Drop' }],
      [{ v: 'Keep' }],
      [{ v: 'Keep' }],
      [{ v: 'Total' }],
    ],
  };

  const [reconciled] = reconcileSpreadsheetFiltersAfterFortune(
    [changed],
    [source],
    cellOperations('sheet-1', [1, 2]),
  );

  expect(reconciled?.config?.rowhidden).toEqual({ '1': 0, '3': 0 });
  expect(reconciled?.config?.rowhidden).not.toHaveProperty('4');
  expect(reconciled?.tables?.[0]?.filters).toEqual(source.tables?.[0]?.filters);
});

test('reapplies dynamic filters with the workbook 1904 date system', () => {
  const source = workSpreadsheetSheetWithAutoFilterCriteria(
    {
      id: 'sheet-1',
      name: '1904 dates',
      data: [
        [{ v: 'Date' }],
        [{ v: 0, ct: { fa: 'yyyy-MM-dd', t: 'd' } }],
        [{ v: 31, ct: { fa: 'yyyy-MM-dd', t: 'd' } }],
      ],
      filter_select: { row: [0, 2], column: [0, 0] },
    },
    0,
    { type: 'dynamic', kind: 'month-1' },
    { dateSystem: '1904', now: new Date(2026, 5, 17) },
  );
  if (!source) throw new Error('Expected a 1904 AutoFilter fixture.');
  const changed: WorkSpreadsheetSheet = {
    ...source,
    data: [
      [{ v: 'Date' }],
      [{ v: 31, ct: { fa: 'yyyy-MM-dd', t: 'd' } }],
      [{ v: 0, ct: { fa: 'yyyy-MM-dd', t: 'd' } }],
    ],
  };

  const [reconciled] = reconcileSpreadsheetFiltersAfterFortune(
    [changed],
    [source],
    cellOperations('sheet-1', [1, 2]),
    { dateSystem: '1904', now: new Date(2026, 5, 17) },
  );

  expect(reconciled?.filter?.['0']?.rowhidden).toEqual({ '1': 0 });
  expect(reconciled?.config?.rowhidden).toEqual({ '1': 0 });
});

test('remaps opaque native value-filter ownership through the sorted row order', () => {
  const source: WorkSpreadsheetSheet = {
    id: 'sheet-1',
    name: 'Native filter',
    data: [
      [{ v: 'Status' }],
      [{ v: 'Risk' }],
      [{ v: 'Keep' }],
      [{ v: 'Keep' }],
    ],
    filter_select: { row: [0, 3], column: [0, 0] },
    filter: {
      '0': {
        caljs: {},
        cindex: 0,
        edc: 0,
        edr: 3,
        optionstate: true,
        rowhidden: { '2': 0, '3': 0 },
        stc: 0,
        str: 0,
      },
    },
    config: { rowhidden: { '2': 0, '3': 0 } },
  };
  const sorted: WorkSpreadsheetSheet = {
    ...source,
    data: [
      [{ v: 'Status' }],
      [{ v: 'Keep' }],
      [{ v: 'Keep' }],
      [{ v: 'Risk' }],
    ],
  };

  const reconciled = reconcileSpreadsheetFiltersAfterSort(
    sorted,
    source,
    { row: [0, 3], column: [0, 0] },
    [1, 2, 0],
  );

  expect(reconciled?.filter?.['0']).toMatchObject({
    caljs: {},
    rowhidden: { '1': 0, '2': 0 },
  });
  expect(reconciled?.config?.rowhidden).toEqual({ '1': 0, '2': 0 });
});

test('skips unrelated edits and refuses unbounded filter rescans', () => {
  const filtered = workSpreadsheetSheetWithAutoFilterCriteria(
    {
      id: 'sheet-1',
      name: 'Filtered',
      data: [[{ v: 'Status' }], [{ v: 'Keep' }], [{ v: 'Drop' }]],
      filter_select: { row: [0, 2], column: [0, 0] },
    },
    0,
    { type: 'equals', value: 'Keep' },
  );
  if (!filtered) throw new Error('Expected an AutoFilter fixture.');
  const unrelated = { ...filtered, zoomRatio: 1.25 };
  const [unchanged] = reconcileSpreadsheetFiltersAfterFortune(
    [unrelated],
    [filtered],
    [
      {
        id: 'sheet-1',
        op: 'replace',
        path: ['data', 1, 2],
        value: { v: 'Outside' },
      },
    ],
  );
  expect(unchanged).toBe(unrelated);

  const oversized: WorkSpreadsheetSheet = {
    id: 'large',
    name: 'Large',
    row: 600_002,
    column: 2,
    tables: [
      {
        id: 'table-large',
        name: 'TableLarge',
        range: { row: [0, 600_001], column: [0, 1] },
        columns: [{ name: 'A' }, { name: 'B' }],
        filters: [
          { column: 0, criteria: { type: 'equals', value: 'A' } },
          { column: 1, criteria: { type: 'equals', value: 'B' } },
        ],
        headerRow: true,
        totalsRow: false,
        style: { family: 'medium', number: 2 },
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    ],
  };
  const [bounded] = reconcileSpreadsheetFiltersAfterFortune(
    [oversized],
    [oversized],
    cellOperations('large', [1]),
  );
  expect(bounded).toBe(oversized);
});

function cellOperations(sheetId: string, rows: number[]): Op[] {
  return rows.map((row) => ({
    id: sheetId,
    op: 'replace',
    path: ['data', row, 0],
    value: { v: row },
  }));
}
