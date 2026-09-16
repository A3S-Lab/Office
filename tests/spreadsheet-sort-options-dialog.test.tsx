import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor, within } from '@testing-library/react';
import { useState } from 'react';
import type {
  SpreadsheetSortDialogSource,
  SpreadsheetSortDialogValue,
} from '../src/internal/features/work/editors/spreadsheet-sort';
import { SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS } from '../src/internal/features/work/editors/spreadsheet-sort-custom-list';
import { SpreadsheetSortDialog } from '../src/internal/features/work/editors/spreadsheet-sort-dialog';
import { SpreadsheetSortOptionsDialog } from '../src/internal/features/work/editors/spreadsheet-sort-options-dialog';

function chooseOfficeSelectOption(
  ariaLabel: string,
  optionName: string | RegExp,
) {
  fireEvent.click(screen.getByRole('combobox', { name: ariaLabel }));
  fireEvent.click(screen.getByRole('option', { name: optionName }));
}

function expectOfficeSelectValue(ariaLabel: string, value: string) {
  expect(screen.getByRole('combobox', { name: ariaLabel })).toHaveAttribute(
    'data-selected-value',
    value,
  );
}

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
  expect(
    within(options).getByRole('radio', { name: '拼音排序' }),
  ).toBeChecked();
  expect(
    within(options).getByRole('checkbox', { name: '区分大小写' }),
  ).not.toBeChecked();
  fireEvent.click(within(options).getByRole('radio', { name: '笔画排序' }));
  fireEvent.click(
    within(options).getByRole('checkbox', { name: '区分大小写' }),
  );
  fireEvent.click(within(options).getByRole('radio', { name: /按行排序/ }));
  fireEvent.click(within(options).getByRole('button', { name: '确定' }));

  expect(screen.queryByRole('dialog', { name: '排序选项' })).toBeNull();
  expect(header).not.toBeChecked();
  expect(header).toBeDisabled();
  expectOfficeSelectValue('排序条件 1 行', '1');
  chooseOfficeSelectOption('排序条件 1 排序依据', '单元格颜色');
  chooseOfficeSelectOption('排序条件 1 目标外观', '单元格颜色 #FCE8E6');
  fireEvent.click(screen.getByRole('combobox', { name: '排序条件 1 位置' }));
  expect(screen.getByRole('option', { name: '置于左侧' })).toBeInTheDocument();
  fireEvent.click(screen.getByRole('option', { name: '置于右侧' }));
  expect(screen.getByText(/单元格颜色 #FCE8E6，置于右侧/)).toBeVisible();

  fireEvent.click(screen.getByRole('button', { name: '添加条件' }));
  expectOfficeSelectValue('排序条件 2 行', '0');
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(applied).toEqual([
    {
      orientation: 'left-to-right',
      caseSensitive: true,
      textMethod: 'stroke',
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

test('changes text comparison without resetting existing sort levels', () => {
  const source = sortSource();
  const applied: SpreadsheetSortDialogValue[] = [];
  render(
    <SpreadsheetSortDialog
      source={{
        ...source,
        value: {
          ...source.value,
          keys: [
            { index: 0, direction: 'ascending' },
            { index: 1, direction: 'descending' },
          ],
        },
      }}
      restoreFocusTarget={() => null}
      onApply={(value) => {
        applied.push(value);
        return true;
      }}
      onClose={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '选项…' }));
  const options = screen.getByRole('dialog', { name: '排序选项' });
  fireEvent.click(within(options).getByRole('radio', { name: '笔画排序' }));
  fireEvent.click(
    within(options).getByRole('checkbox', { name: '区分大小写' }),
  );
  fireEvent.click(within(options).getByRole('button', { name: '确定' }));

  expectOfficeSelectValue('排序条件 1 列', '0');
  expectOfficeSelectValue('排序条件 2 列', '1');
  fireEvent.click(screen.getByRole('button', { name: '确定' }));
  expect(applied).toEqual([
    {
      orientation: 'top-to-bottom',
      caseSensitive: true,
      textMethod: 'stroke',
      hasHeader: true,
      keys: [
        { index: 0, direction: 'ascending' },
        { index: 1, direction: 'descending' },
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
      caseSensitive: false,
      textMethod: 'pinyin',
      hasHeader: true,
      keys: [{ index: 0, direction: 'ascending' }],
    },
  };
}

test('cancels a dirty Sort Options draft before Escape closes the dialog', async () => {
  const closes: string[] = [];
  const value = {
    caseSensitive: false,
    textMethod: 'pinyin' as const,
    orientation: 'top-to-bottom' as const,
  };

  function Harness() {
    const [open, setOpen] = useState(true);
    return (
      <>
        <button type="button" data-testid="sort-options-dialog-launcher">
          选项…
        </button>
        {open ? (
          <SpreadsheetSortOptionsDialog
            value={value}
            restoreFocusTarget={() =>
              document.querySelector<HTMLElement>(
                '[data-testid="sort-options-dialog-launcher"]',
              )
            }
            onApply={() => {
              setOpen(false);
            }}
            onClose={() => {
              closes.push('close');
              setOpen(false);
            }}
          />
        ) : null}
      </>
    );
  }

  render(<Harness />);
  const caseSensitive = screen.getByRole('checkbox', { name: '区分大小写' });
  fireEvent.click(caseSensitive);
  expect(caseSensitive).toBeChecked();

  fireEvent.keyDown(caseSensitive, { key: 'Escape' });
  expect(screen.getByRole('dialog', { name: '排序选项' })).toBeInTheDocument();
  expect(caseSensitive).not.toBeChecked();
  expect(closes).toEqual([]);

  fireEvent.keyDown(caseSensitive, { key: 'Escape' });
  expect(screen.queryByRole('dialog', { name: '排序选项' })).toBeNull();
  expect(closes).toEqual(['close']);
  await waitFor(() =>
    expect(screen.getByTestId('sort-options-dialog-launcher')).toHaveFocus(),
  );
});
