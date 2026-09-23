import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { SpreadsheetPivotPanel } from '../src/internal/features/work/editors/spreadsheet-pivot-panel';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

const content: WorkSpreadsheetContent = {
  type: 'spreadsheet',
  sheets: [
    {
      id: 'sheet-1',
      name: '数据',
      status: 1,
      data: [
        [{ v: '类别' }, { v: '金额' }],
        [{ v: '产品 A' }, { v: 12 }],
        [{ v: '产品 B' }, { v: 18 }],
      ],
      pivotTables: [
        {
          id: 'pivot-1',
          name: '销售汇总',
          sourceSheetId: 'sheet-1',
          sourceReference: 'A1:B3',
          anchor: 'D1',
          rowFields: [0],
          columnFields: [],
          values: [{ fieldIndex: 1, aggregation: 'sum' }],
          rowGrandTotals: true,
          columnGrandTotals: true,
          styleName: 'PivotStyleLight16',
          refreshOnLoad: false,
        },
        {
          id: 'pivot-2',
          name: '产品汇总',
          sourceSheetId: 'sheet-1',
          sourceReference: 'A1:B3',
          anchor: 'H1',
          rowFields: [0],
          columnFields: [],
          values: [{ fieldIndex: 1, aggregation: 'average' }],
          rowGrandTotals: true,
          columnGrandTotals: true,
          styleName: 'PivotStyleLight18',
          refreshOnLoad: false,
        },
      ],
    },
  ],
};

test('protects a dirty pivot draft from switching and refresh actions', () => {
  const changes: WorkSpreadsheetContent[] = [];
  render(
    <SpreadsheetPivotPanel
      content={content}
      activeSheetId="sheet-1"
      onChange={(next) => changes.push(next)}
    />,
  );
  const name = screen.getByRole('textbox', { name: '透视表名称' });
  fireEvent.change(name, { target: { value: '尚未保存的汇总' } });
  fireEvent.click(screen.getByRole('button', { name: /产品汇总/ }));

  expect(name).toHaveValue('尚未保存的汇总');
  expect(screen.getByRole('alert')).toHaveTextContent('请先保存或取消');
  fireEvent.click(screen.getByRole('button', { name: '刷新全部' }));
  expect(changes).toEqual([]);

  fireEvent.click(screen.getByRole('button', { name: '取消更改' }));
  expect(name).toHaveValue('销售汇总');
  fireEvent.click(screen.getByRole('button', { name: /产品汇总/ }));
  expect(name).toHaveValue('产品汇总');
});

test('exposes slicer-style multi-select checkboxes on report filter fields', () => {
  const filtered: WorkSpreadsheetContent = {
    ...content,
    sheets: [
      {
        ...content.sheets[0],
        data: [
          [{ v: '区域' }, { v: '产品' }, { v: '金额' }],
          [{ v: '东' }, { v: 'A' }, { v: 10 }],
          [{ v: '西' }, { v: 'B' }, { v: 20 }],
        ],
        pivotTables: [
          {
            id: 'pivot-filter',
            name: '区域汇总',
            sourceSheetId: 'sheet-1',
            sourceReference: 'A1:C3',
            anchor: 'E1',
            rowFields: [1],
            columnFields: [],
            reportFilters: [{ fieldIndex: 0 }],
            values: [{ fieldIndex: 2, aggregation: 'sum' }],
            rowGrandTotals: true,
            columnGrandTotals: true,
            styleName: 'PivotStyleLight16',
            refreshOnLoad: false,
          },
        ],
      },
    ],
  };
  const changes: WorkSpreadsheetContent[] = [];
  render(
    <SpreadsheetPivotPanel
      content={filtered}
      activeSheetId="sheet-1"
      onChange={(next) => changes.push(next)}
    />,
  );

  expect(screen.getByRole('group', { name: '区域 切片器筛选' })).toBeTruthy();
  expect(screen.getByText('切片器多选筛选')).toBeTruthy();
  // From the default "all" state, clearing one value leaves the rest selected.
  fireEvent.click(screen.getByRole('checkbox', { name: '区域 西' }));
  fireEvent.click(screen.getByRole('button', { name: '保存并刷新' }));

  expect(changes).toHaveLength(1);
  expect(changes[0].sheets[0].pivotTables?.[0].reportFilters).toEqual([
    { fieldIndex: 0, selectedItems: ['东'] },
  ]);
});
