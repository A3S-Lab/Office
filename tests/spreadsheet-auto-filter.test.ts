import type { Selection } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import {
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

  test('rejects empty data, merged ranges, and pivot output', () => {
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

  test('preserves the filter range and hidden rows through XLSX export', async () => {
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
              config: { ...sheet.config, rowhidden: { '4': 0 } },
              filter: {},
              filter_select: { row: [2, 6], column: [0, 5] },
            }
          : sheet,
      ),
    };

    const blob = await createWorkArtifactBlob({ ...artifact, content });
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
    expect(imported.content.sheets[0]?.config?.rowhidden).toMatchObject({
      '4': 0,
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
