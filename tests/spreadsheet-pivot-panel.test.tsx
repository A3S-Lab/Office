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
