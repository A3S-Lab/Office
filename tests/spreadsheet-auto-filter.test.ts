import type { Selection } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import {
  applySpreadsheetAutoFilterCriteria,
  clearSpreadsheetAutoFilterCriteria,
  spreadsheetAutoFilterCriteria,
  spreadsheetAutoFilterRange,
  toggleSpreadsheetAutoFilter,
} from '../src/internal/features/work/editors/spreadsheet-auto-filter';
import {
  createWorkArtifactBlob,
  importWorkFile,
} from '../src/internal/features/work/work-file-io';
import { createWorkArtifact } from '../src/internal/features/work/work-templates';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../src/internal/features/work/work-types';

describe('spreadsheet AutoFilter model', () => {
  test('expands a cell in a dense table to the WPS current region', () => {
    const sheet = quarterlySheet();

    expect(spreadsheetAutoFilterRange(sheet, selection(4, 3))).toEqual({
      row: [2, 6],
      column: [0, 6],
    });
    expect(spreadsheetAutoFilterRange(sheet, selection(2, 6))).toEqual({
      row: [2, 6],
      column: [0, 6],
    });
  });

  test('reads maximum sparse worksheet bounds without spreading logical rows', () => {
    const data: NonNullable<WorkSpreadsheetSheet['data']> = [];
    data.length = 1_048_576;
    data[0] = [];
    data[0].length = 16_384;
    data[0][0] = { v: 'Header', m: 'Header' };
    data[1] = [];
    data[1].length = 16_384;
    data[1][0] = { v: 'Value', m: 'Value' };
    data[1_048_575] = [];
    data[1_048_575].length = 16_384;
    data[1_048_575][16_383] = { v: 'Tail', m: 'Tail' };

    expect(
      spreadsheetAutoFilterRange(
        {
          id: 'maximum-sparse',
          name: 'Maximum sparse',
          data,
          row: 1_048_576,
          column: 16_384,
          config: {},
        },
        selection(0, 0),
      ),
    ).toEqual({ row: [0, 1], column: [0, 0] });
    expect(Object.keys(data)).toEqual(['0', '1', '1048575']);
  });

  test('keeps an explicit multi-row selection as the filter range', () => {
    const sheet = quarterlySheet();
    expect(
      spreadsheetAutoFilterRange(sheet, {
        row: [2, 4],
        column: [1, 3],
        row_focus: 4,
        column_focus: 3,
      }),
    ).toEqual({ row: [2, 4], column: [1, 3] });
  });

  test('rejects empty data, merged ranges, tables, and pivot output', () => {
    const sheet = quarterlySheet();
    expect(
      spreadsheetAutoFilterRange(
        {
          ...sheet,
          data: [[{ v: 'Header' }], []],
          config: {},
        },
        { row: [0, 1], column: [0, 0] },
      ),
    ).toBeNull();
    expect(
      spreadsheetAutoFilterRange(
        {
          ...sheet,
          config: {
            ...sheet.config,
            merge: {
              ...sheet.config?.merge,
              '2_0': { r: 2, c: 0, rs: 1, cs: 2 },
            },
          },
        },
        selection(3, 1),
      ),
    ).toBeNull();
    expect(
      spreadsheetAutoFilterRange(
        { ...sheet, isPivotTable: true },
        selection(3, 1),
      ),
    ).toBeNull();
    expect(
      spreadsheetAutoFilterRange(
        {
          ...sheet,
          tables: [
            {
              id: 'table-1',
              name: 'Table1',
              range: { row: [2, 6], column: [0, 6] },
              columns: Array.from({ length: 7 }, (_, index) => ({
                name: `Column${index + 1}`,
              })),
              filters: [],
              headerRow: true,
              totalsRow: false,
              style: { family: 'medium', number: 2 },
              showFirstColumn: false,
              showLastColumn: false,
              showRowStripes: true,
              showColumnStripes: false,
            },
          ],
        },
        selection(3, 1),
      ),
    ).toBeNull();
  });

  test('toggles controlled filter state without mutating the workbook', () => {
    const content = workbook(quarterlySheet());
    const enabled = toggleSpreadsheetAutoFilter(
      content,
      'sheet-1',
      selection(3, 4),
    );

    expect(enabled).not.toBeNull();
    expect(enabled).not.toBe(content);
    expect(content.sheets[0]?.filter_select).toBeUndefined();
    expect(enabled?.sheets[0]?.filter_select).toEqual({
      row: [2, 6],
      column: [0, 6],
    });
    expect(enabled?.sheets[0]?.filter).toEqual({});
    expect(enabled?.sheets[0]?.luckysheet_select_save).toEqual([
      selection(3, 4),
    ]);
  });

  test('clears only rows hidden by active filter criteria', () => {
    const sheet: WorkSpreadsheetSheet = {
      ...quarterlySheet(),
      filter_select: { row: [2, 6], column: [0, 6] },
      filter: {
        '0': { rowhidden: { '4': 0, '6': 0 } },
        '2': { rowhidden: { '4': 0 } },
      },
      config: {
        ...quarterlySheet().config,
        rowhidden: { '4': 0, '6': 0, '9': 0 },
      },
    };

    const disabled = toggleSpreadsheetAutoFilter(
      workbook(sheet),
      'sheet-1',
      selection(2, 0),
    );

    expect(disabled?.sheets[0]?.filter_select).toBeUndefined();
    expect(disabled?.sheets[0]?.filter).toBeUndefined();
    expect(disabled?.sheets[0]?.config?.rowhidden).toEqual({ '9': 0 });
  });

  test('preserves imported hidden rows when criteria metadata is unavailable', () => {
    const sheet: WorkSpreadsheetSheet = {
      ...quarterlySheet(),
      filter_select: { row: [2, 6], column: [0, 6] },
      filter: {},
      config: {
        ...quarterlySheet().config,
        rowhidden: { '5': 0, '9': 0 },
      },
    };

    const disabled = toggleSpreadsheetAutoFilter(
      workbook(sheet),
      'sheet-1',
      selection(2, 0),
    );
    expect(disabled?.sheets[0]?.config?.rowhidden).toEqual({
      '5': 0,
      '9': 0,
    });
  });

  test('applies typed text criteria while preserving manually hidden rows', () => {
    const sheet: WorkSpreadsheetSheet = {
      ...quarterlySheet(),
      filter: {},
      filter_select: { row: [2, 6], column: [0, 6] },
      config: { ...quarterlySheet().config, rowhidden: { '9': 0 } },
    };

    const filtered = applySpreadsheetAutoFilterCriteria(
      workbook(sheet),
      'sheet-1',
      6,
      { type: 'contains', value: '风险' },
    );

    expect(filtered).not.toBeNull();
    expect(spreadsheetAutoFilterCriteria(filtered?.sheets[0], 6)).toEqual({
      type: 'contains',
      value: '风险',
    });
    expect(filtered?.sheets[0]?.filter?.['6']).toMatchObject({
      optionstate: true,
      rowhidden: { '3': 0, '5': 0 },
      str: 2,
      edr: 6,
      cindex: 6,
      stc: 0,
      edc: 6,
    });
    expect(filtered?.sheets[0]?.config?.rowhidden).toEqual({
      '3': 0,
      '5': 0,
      '9': 0,
    });
    expect(sheet.filter).toEqual({});
    expect(sheet.config?.rowhidden).toEqual({ '9': 0 });
  });

  test('preserves a manually hidden row that overlaps a replaced criterion', () => {
    const sheet: WorkSpreadsheetSheet = {
      ...quarterlySheet(),
      filter: {},
      filter_select: { row: [2, 6], column: [0, 6] },
      config: { ...quarterlySheet().config, rowhidden: { '3': 0 } },
    };
    const riskOnly = applySpreadsheetAutoFilterCriteria(
      workbook(sheet),
      'sheet-1',
      6,
      { type: 'contains', value: '风险' },
    );
    const normalOnly = riskOnly
      ? applySpreadsheetAutoFilterCriteria(riskOnly, 'sheet-1', 6, {
          type: 'contains',
          value: '正常',
        })
      : null;
    const cleared = normalOnly
      ? clearSpreadsheetAutoFilterCriteria(normalOnly, 'sheet-1', 6)
      : null;

    expect(riskOnly?.sheets[0]?.config?.rowhidden).toEqual({
      '3': 0,
      '5': 0,
    });
    expect(normalOnly?.sheets[0]?.config?.rowhidden).toEqual({
      '3': 0,
      '4': 0,
      '6': 0,
    });
    expect(cleared?.sheets[0]?.config?.rowhidden).toEqual({ '3': 0 });
  });

  test('combines numeric criteria across columns and clears one owner only', () => {
    const sheet: WorkSpreadsheetSheet = {
      ...quarterlySheet(),
      filter: {},
      filter_select: { row: [2, 6], column: [0, 6] },
      config: { ...quarterlySheet().config, rowhidden: { '9': 0 } },
    };
    const textFiltered = applySpreadsheetAutoFilterCriteria(
      workbook(sheet),
      'sheet-1',
      6,
      { type: 'contains', value: '风险' },
    );
    const numericFiltered = textFiltered
      ? applySpreadsheetAutoFilterCriteria(textFiltered, 'sheet-1', 1, {
          type: 'greater-than',
          value: '130',
        })
      : null;

    expect(numericFiltered?.sheets[0]?.filter?.['1']).toMatchObject({
      rowhidden: { '3': 0, '4': 0, '5': 0 },
    });
    expect(numericFiltered?.sheets[0]?.config?.rowhidden).toEqual({
      '3': 0,
      '4': 0,
      '5': 0,
      '9': 0,
    });

    const cleared = numericFiltered
      ? clearSpreadsheetAutoFilterCriteria(numericFiltered, 'sheet-1', 1)
      : null;
    expect(cleared?.sheets[0]?.filter?.['1']).toBeUndefined();
    expect(cleared?.sheets[0]?.filter?.['6']).toBeDefined();
    expect(cleared?.sheets[0]?.config?.rowhidden).toEqual({
      '3': 0,
      '5': 0,
      '9': 0,
    });
  });

  test('supports blanks, non-blanks, inclusive ranges, and selected values', () => {
    const sheet: WorkSpreadsheetSheet = {
      id: 'sheet-1',
      name: '筛选条件',
      data: [
        [{ v: '值' }],
        [{ v: 4 }],
        [{ v: 8 }],
        [{ v: 12 }],
        [{ v: '' }],
        [],
      ],
      filter: {},
      filter_select: { row: [0, 5], column: [0, 0] },
      config: {},
    };
    const scenarios = [
      [{ type: 'between', lower: '5', upper: '10' }, ['1', '3', '4', '5']],
      [{ type: 'blanks' }, ['1', '2', '3']],
      [{ type: 'non-blanks' }, ['4', '5']],
      [
        { type: 'values', values: ['4', '12'], includeBlanks: false },
        ['2', '4', '5'],
      ],
    ] as const;

    for (const [criteria, hiddenRows] of scenarios) {
      const filtered = applySpreadsheetAutoFilterCriteria(
        workbook(sheet),
        'sheet-1',
        0,
        criteria,
      );
      expect(Object.keys(filtered?.sheets[0]?.config?.rowhidden ?? {})).toEqual(
        hiddenRows,
      );
    }
  });

  test('fails closed for inactive, header, and out-of-range filter columns', () => {
    const active: WorkSpreadsheetSheet = {
      ...quarterlySheet(),
      filter: {},
      filter_select: { row: [2, 6], column: [1, 6] },
    };
    const criteria = { type: 'equals', value: '正常' } as const;

    expect(
      applySpreadsheetAutoFilterCriteria(
        workbook(quarterlySheet()),
        'sheet-1',
        6,
        criteria,
      ),
    ).toBeNull();
    expect(
      applySpreadsheetAutoFilterCriteria(
        workbook(active),
        'sheet-1',
        0,
        criteria,
      ),
    ).toBeNull();
    expect(
      applySpreadsheetAutoFilterCriteria(
        workbook(active),
        'sheet-1',
        7,
        criteria,
      ),
    ).toBeNull();
    expect(
      applySpreadsheetAutoFilterCriteria(
        workbook(active),
        'missing-sheet',
        6,
        criteria,
      ),
    ).toBeNull();
  });

  test('preserves the filter range, criteria, and hidden rows through XLSX export', async () => {
    const artifact = createWorkArtifact('quarterly-plan');
    if (artifact.content.type !== 'spreadsheet') {
      throw new Error('Expected the quarterly plan spreadsheet.');
    }
    const content: WorkSpreadsheetContent = {
      ...artifact.content,
      sheets: artifact.content.sheets.map((sheet, index) =>
        index === 0
          ? {
              ...sheet,
              config: { ...sheet.config, rowhidden: { '9': 0 } },
              filter: {},
              filter_select: { row: [2, 6], column: [0, 5] },
            }
          : sheet,
      ),
    };

    const filtered = applySpreadsheetAutoFilterCriteria(
      content,
      content.sheets[0]?.id ?? '',
      3,
      { type: 'greater-than', value: '0.95' },
    );
    if (!filtered) throw new Error('Expected the filter criteria to apply.');

    const blob = await createWorkArtifactBlob({
      ...artifact,
      content: filtered,
    });
    const imported = await importWorkFile(
      new File([blob], 'quarterly-plan-filtered.xlsx', { type: blob.type }),
    );
    if (imported.content.type !== 'spreadsheet') {
      throw new Error('Expected an imported spreadsheet.');
    }

    expect(imported.content.sheets[0]?.filter_select).toEqual({
      row: [2, 6],
      column: [0, 5],
    });
    expect(
      spreadsheetAutoFilterCriteria(imported.content.sheets[0], 3),
    ).toEqual({ type: 'greater-than', value: '0.95' });
    expect(imported.content.sheets[0]?.filter?.['3']?.rowhidden).toEqual({
      '4': 0,
      '5': 0,
    });
    expect(imported.content.sheets[0]?.config?.rowhidden).toMatchObject({
      '4': 0,
      '5': 0,
      '9': 0,
    });
  });
});

function selection(row: number, column: number): Selection {
  return {
    row: [row, row],
    column: [column, column],
    row_focus: row,
    column_focus: column,
  };
}

function workbook(sheet: WorkSpreadsheetSheet): WorkSpreadsheetContent {
  return { type: 'spreadsheet', sheets: [sheet] };
}

function quarterlySheet(): WorkSpreadsheetSheet {
  return {
    id: 'sheet-1',
    name: '季度经营',
    row: 40,
    column: 12,
    config: {
      merge: {
        '0_0': { r: 0, c: 0, rs: 1, cs: 7 },
      },
    },
    data: [
      [{ v: '2026 年度经营概览' }],
      [],
      [
        { v: '季度' },
        { v: '收入' },
        { v: '成本' },
        { v: '利润' },
        { v: '利润率' },
        { v: '负责人' },
        { v: '状态' },
      ],
      [
        { v: 'Q1' },
        { v: 120 },
        { v: 84 },
        { v: 36 },
        { v: 0.3 },
        { v: '林夏' },
        { v: '正常' },
      ],
      [
        { v: 'Q2' },
        { v: 128 },
        { v: 91 },
        { v: 37 },
        { v: 0.289 },
        { v: '周文' },
        { v: '有风险' },
      ],
      [
        { v: 'Q3' },
        { v: 0 },
        { v: 0 },
        { f: '=B6-C6', v: 0 },
        { v: false },
        { v: '陈明' },
        { v: '正常' },
      ],
      [
        { v: 'Q4' },
        { v: 146 },
        { v: 101 },
        { v: 45 },
        { v: 0.308 },
        { v: '宋杰' },
        { v: '有风险' },
      ],
    ],
  };
}
