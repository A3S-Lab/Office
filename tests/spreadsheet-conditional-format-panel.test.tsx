import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { SpreadsheetConditionalFormatPanel } from '../src/internal/features/work/editors/spreadsheet-conditional-format-panel';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('uses product language in the empty conditional-format state', () => {
  render(
    <SpreadsheetConditionalFormatPanel
      content={{
        type: 'spreadsheet',
        sheets: [{ id: 'sheet-1', name: '工作表 1', data: [] }],
      }}
      onChange={() => undefined}
    />,
  );

  expect(screen.getByLabelText('条件格式规则')).toBeInTheDocument();
  expect(screen.getByRole('status')).toHaveTextContent('还没有条件格式规则。');
  expect(screen.queryByText(/Work/)).toBeNull();
});

const contentWithRules: WorkSpreadsheetContent = {
  type: 'spreadsheet',
  sheets: [
    {
      id: 'sheet-1',
      name: '工作表 1',
      status: 1,
      data: [],
      luckysheet_conditionformat_save: [
        {
          type: 'default',
          cellrange: [{ row: [0, 9], column: [0, 0] }],
          format: { textColor: '#9c0006', cellColor: '#ffc7ce' },
          conditionName: 'greaterThan',
          conditionRange: [],
          conditionValue: ['0'],
        },
        {
          type: 'default',
          cellrange: [{ row: [0, 9], column: [1, 1] }],
          format: { textColor: '#006100', cellColor: '#c6efce' },
          conditionName: 'lessThan',
          conditionRange: [],
          conditionValue: ['100'],
        },
      ],
    },
  ],
};

test('preserves a dirty conditional-format draft across unrelated workbook updates', () => {
  const props = { onChange: () => undefined };
  const { rerender } = render(
    <SpreadsheetConditionalFormatPanel content={contentWithRules} {...props} />,
  );
  const reference = screen.getByRole('textbox', {
    name: '条件格式范围',
  });
  fireEvent.change(reference, { target: { value: 'A1:A20' } });

  rerender(
    <SpreadsheetConditionalFormatPanel
      content={{
        ...contentWithRules,
        sheets: [{ ...contentWithRules.sheets[0], data: [[{ v: '已更新' }]] }],
      }}
      {...props}
    />,
  );

  expect(reference).toHaveValue('A1:A20');
});

test('requires saving or cancelling before switching conditional-format rules', () => {
  const changes: WorkSpreadsheetContent[] = [];
  render(
    <SpreadsheetConditionalFormatPanel
      content={contentWithRules}
      onChange={(next) => changes.push(next)}
    />,
  );
  const reference = screen.getByRole('textbox', {
    name: '条件格式范围',
  });
  fireEvent.change(reference, { target: { value: 'A1:A20' } });
  fireEvent.click(screen.getByRole('button', { name: /小于/ }));

  expect(reference).toHaveValue('A1:A20');
  expect(screen.getByRole('alert')).toHaveTextContent('请先保存或取消');
  expect(changes).toEqual([]);

  fireEvent.click(screen.getByRole('button', { name: '取消更改' }));
  expect(reference).toHaveValue('A1:A10');
  fireEvent.click(screen.getByRole('button', { name: /小于/ }));
  expect(reference).toHaveValue('B1:B10');
});

test('keeps a newly saved conditional rule stable until controlled content catches up', () => {
  const changes: WorkSpreadsheetContent[] = [];
  render(
    <SpreadsheetConditionalFormatPanel
      content={{
        type: 'spreadsheet',
        sheets: [{ id: 'sheet-1', name: '工作表 1', status: 1, data: [] }],
      }}
      onChange={(next) => changes.push(next)}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '保存规则' }));

  expect(changes).toHaveLength(1);
  expect(changes[0].sheets[0].luckysheet_conditionformat_save).toHaveLength(1);
  expect(screen.getByRole('textbox', { name: '条件格式范围' })).toHaveValue(
    'A1:A10',
  );
  expect(screen.getByRole('button', { name: '保存规则' })).toBeDisabled();
});
