import { expect, test } from '@rstest/core';
import { buildSpreadsheetPivotOutput } from '../src/internal/features/work/work-spreadsheet-pivot-engine';
import {
  spreadsheetPivotFields,
  spreadsheetPivotValidation,
} from '../src/internal/features/work/work-spreadsheet-pivots';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetPivotTable,
} from '../src/internal/features/work/work-types';

function workbook(
  reportFilters: WorkSpreadsheetPivotTable['reportFilters'],
): {
  content: WorkSpreadsheetContent;
  pivot: WorkSpreadsheetPivotTable;
} {
  const pivot: WorkSpreadsheetPivotTable = {
    id: 'pivot-1',
    name: '销售汇总',
    sourceSheetId: 'sheet-1',
    sourceReference: 'A1:C5',
    anchor: 'E1',
    rowFields: [0],
    columnFields: [],
    reportFilters,
    values: [{ fieldIndex: 2, aggregation: 'sum' }],
    rowGrandTotals: true,
    columnGrandTotals: true,
    styleName: 'PivotStyleLight16',
    refreshOnLoad: false,
  };
  const content: WorkSpreadsheetContent = {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: '数据',
        status: 1,
        data: [
          [{ v: '产品' }, { v: '区域' }, { v: '金额' }],
          [{ v: 'A' }, { v: '东' }, { v: 10 }],
          [{ v: 'B' }, { v: '西' }, { v: 20 }],
          [{ v: 'A' }, { v: '西' }, { v: 30 }],
          [{ v: 'C' }, { v: '东' }, { v: 40 }],
        ],
        pivotTables: [pivot],
      },
    ],
  };
  return { content, pivot };
}

function outputLabels(
  content: WorkSpreadsheetContent,
  pivot: WorkSpreadsheetPivotTable,
): string[] {
  const fields = spreadsheetPivotFields(content, pivot);
  const sourceSheet = content.sheets[0];
  const matrix = buildSpreadsheetPivotOutput(
    sourceSheet,
    {
      startRow: 0,
      endRow: 4,
      startColumn: 0,
      endColumn: 2,
    },
    fields,
    pivot,
  );
  return matrix.map((row) =>
    row
      .filter(Boolean)
      .map((cell) => String(cell?.m ?? cell?.v ?? ''))
      .join('|'),
  );
}

test('legacy single-select report filter still matches one item', () => {
  const { content, pivot } = workbook([
    { fieldIndex: 1, selectedItem: '东' },
  ]);
  const labels = outputLabels(content, pivot);
  expect(labels[0]).toContain('区域');
  expect(labels[0]).toContain('东');
  expect(labels.some((line) => line.startsWith('A|'))).toBe(true);
  expect(labels.some((line) => line.startsWith('C|'))).toBe(true);
  expect(labels.some((line) => line.startsWith('B|'))).toBe(false);
});

test('multi-select selectedItems OR-matches slicer-style report filters', () => {
  const { content, pivot } = workbook([
    { fieldIndex: 1, selectedItems: ['东', '西'] },
  ]);
  const labels = outputLabels(content, pivot);
  expect(labels[0]).toContain('东, 西');
  expect(labels.some((line) => line.startsWith('A|'))).toBe(true);
  expect(labels.some((line) => line.startsWith('B|'))).toBe(true);
  expect(labels.some((line) => line.startsWith('C|'))).toBe(true);
});

test('empty selectedItems admits no source rows', () => {
  const { content, pivot } = workbook([{ fieldIndex: 1, selectedItems: [] }]);
  expect(outputLabels(content, pivot)).toEqual([]);
});

test('selectedItems takes precedence over legacy selectedItem', () => {
  const { content, pivot } = workbook([
    { fieldIndex: 1, selectedItem: '东', selectedItems: ['西'] },
  ]);
  const labels = outputLabels(content, pivot);
  expect(labels[0]).toContain('西');
  expect(labels.some((line) => line.startsWith('B|'))).toBe(true);
  expect(labels.some((line) => line.startsWith('C|'))).toBe(false);
});

test('validation rejects selectedItems missing from the source', () => {
  const { content, pivot } = workbook([
    { fieldIndex: 1, selectedItems: ['北'] },
  ]);
  const result = spreadsheetPivotValidation(content, 'sheet-1', pivot);
  expect(result.valid).toBe(false);
  expect(result.code).toBe('pivot.filter-item-invalid');
});
