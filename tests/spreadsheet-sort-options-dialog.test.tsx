import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, within } from '@testing-library/react';
import { SpreadsheetSortDialog } from '../src/internal/features/work/editors/spreadsheet-sort-dialog';
import type {
  SpreadsheetSortDialogSource,
  SpreadsheetSortDialogValue,
} from '../src/internal/features/work/editors/spreadsheet-sort';
import { SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS } from '../src/internal/features/work/editors/spreadsheet-sort-custom-list';

test('switches to WPS row sorting and authors horizontal appearance priorities', () => {
  const applied: SpreadsheetSortDialogValue[] = [];
  render(
    <SpreadsheetSortDialog
      source={sortSource()}
      restoreFocusTarget={() => null}
      onApply={(value) => {
        applied.push(value);
        return true;
      }}
      onClose={() => undefined}
    />,
  );

  const header = screen.getByRole('checkbox', { name: '数据包含标题' });
  expect(header).toBeChecked();
  expect(header).toBeEnabled();
  fireEvent.click(screen.getByRole('button', { name: '选项…' }));

  const options = screen.getByRole('dialog', { name: '排序选项' });
  expect(
    within(options).getByRole('radio', { name: /按列排序/ }),
  ).toBeChecked();
  fireEvent.click(within(options).getByRole('radio', { name: /按行排序/ }));
  fireEvent.click(within(options).getByRole('button', { name: '确定' }));

  expect(screen.queryByRole('dialog', { name: '排序选项' })).toBeNull();
  expect(header).not.toBeChecked();
  expect(header).toBeDisabled();
  expect(screen.getByRole('combobox', { name: '排序条件 1 行' })).toHaveValue(
    '1',
  );
  fireEvent.change(
    screen.getByRole('combobox', { name: '排序条件 1 排序依据' }),
    { target: { value: 'cell-color' } },
  );
  fireEvent.change(
    screen.getByRole('combobox', { name: '排序条件 1 目标外观' }),
    { target: { value: 'cell-color:#fce8e6' } },
  );
  expect(
    within(screen.getByRole('combobox', { name: '排序条件 1 位置' })).getByRole(
      'option',
      { name: '置于左侧' },
    ),
  ).toBeInTheDocument();
  fireEvent.change(screen.getByRole('combobox', { name: '排序条件 1 位置' }), {
    target: { value: 'last' },
  });
  expect(screen.getByText(/单元格颜色 #FCE8E6，置于右侧/)).toBeVisible();

  fireEvent.click(screen.getByRole('button', { name: '添加条件' }));
  expect(screen.getByRole('combobox', { name: '排序条件 2 行' })).toHaveValue(
    '0',
  );
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(applied).toEqual([
    {
      orientation: 'left-to-right',
      hasHeader: false,
      keys: [
        {
          index: 1,
          sortOn: 'cell-color',
          color: '#fce8e6',
          position: 'last',
        },
        { index: 0, direction: 'ascending' },
      ],
    },
  ]);
});

function sortSource(): SpreadsheetSortDialogSource {
  return {
    sheetId: 'sheet-1',
    sheetName: 'Quarterly plan',
    range: { row: [0, 2], column: [0, 3] },
    rangeReference: 'A1:D3',
    activeRow: 1,
    columns: [
      { index: 0, label: 'A（Metric）' },
      { index: 1, label: 'B（Q1）' },
      { index: 2, label: 'C（Q2）' },
      { index: 3, label: 'D（Q3）' },
    ],
    rows: [
      { index: 0, label: '行 1' },
      { index: 1, label: '行 2' },
      { index: 2, label: '行 3' },
    ],
    customLists: SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS,
    appearanceRows: [
      [
        { cellColor: '#4472c4', fontColor: '#ffffff', icon: null },
        { cellColor: '#4472c4', fontColor: '#ffffff', icon: null },
        { cellColor: '#4472c4', fontColor: '#ffffff', icon: null },
        { cellColor: '#4472c4', fontColor: '#ffffff', icon: null },
      ],
      [
        { cellColor: null, fontColor: null, icon: null },
        { cellColor: '#fce8e6', fontColor: null, icon: null },
        { cellColor: null, fontColor: null, icon: null },
        { cellColor: '#fce8e6', fontColor: null, icon: null },
      ],
      [
        { cellColor: null, fontColor: null, icon: null },
        { cellColor: null, fontColor: null, icon: null },
        { cellColor: null, fontColor: null, icon: null },
        { cellColor: null, fontColor: null, icon: null },
      ],
    ],
    value: {
      orientation: 'top-to-bottom',
      hasHeader: true,
      keys: [{ index: 0, direction: 'ascending' }],
    },
  };
}
