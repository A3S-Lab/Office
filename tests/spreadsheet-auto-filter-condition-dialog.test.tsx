import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { SpreadsheetAutoFilterConditionDialog } from '../src/internal/features/work/editors/spreadsheet-auto-filter-condition-dialog';
import type { WorkSpreadsheetFilterCriteria } from '../src/internal/features/work/work-types';

test('validates and applies a bounded numeric condition', () => {
  const applied: WorkSpreadsheetFilterCriteria[] = [];
  render(
    <SpreadsheetAutoFilterConditionDialog
      source={{
        columnLabel: '收入',
        criteria: null,
        hasActiveFilter: false,
        numeric: true,
        sheetName: '季度经营',
      }}
      restoreFocusTarget={() => null}
      onApply={(criteria) => {
        applied.push(criteria);
        return true;
      }}
      onClear={() => false}
      onClose={() => undefined}
    />,
  );

  expect(
    screen.getByRole('dialog', { name: '自定义自动筛选' }),
  ).toHaveTextContent('季度经营!收入');
  fireEvent.change(screen.getByRole('combobox', { name: '筛选条件' }), {
    target: { value: 'between' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '下限' }), {
    target: { value: '120' },
  });
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();

  fireEvent.change(screen.getByRole('textbox', { name: '下限' }), {
    target: { value: 'not-a-number' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '上限' }), {
    target: { value: '140' },
  });
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();

  fireEvent.change(screen.getByRole('textbox', { name: '下限' }), {
    target: { value: ' 150 ' },
  });
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();

  fireEvent.change(screen.getByRole('textbox', { name: '下限' }), {
    target: { value: ' 120 ' },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));
  expect(applied).toEqual([{ type: 'between', lower: '120', upper: '140' }]);
});

test('shows and invokes the owned clear action for an active condition', () => {
  let clears = 0;
  let closes = 0;
  render(
    <SpreadsheetAutoFilterConditionDialog
      source={{
        columnLabel: '状态',
        criteria: { type: 'contains', value: '风险' },
        hasActiveFilter: true,
        numeric: false,
        sheetName: '季度经营',
      }}
      restoreFocusTarget={() => null}
      onApply={() => false}
      onClear={() => {
        clears += 1;
        return true;
      }}
      onClose={() => {
        closes += 1;
      }}
    />,
  );

  expect(screen.getByRole('combobox', { name: '筛选条件' })).toHaveValue(
    'contains',
  );
  expect(screen.getByRole('textbox', { name: '筛选值' })).toHaveValue('风险');
  fireEvent.click(screen.getByRole('button', { name: '清除此列筛选' }));
  expect(clears).toBe(1);
  expect(closes).toBe(1);
});

test('authors two custom conditions with an explicit OR relationship', () => {
  const applied: WorkSpreadsheetFilterCriteria[] = [];
  render(
    <SpreadsheetAutoFilterConditionDialog
      source={{
        columnLabel: '状态',
        criteria: null,
        hasActiveFilter: false,
        numeric: false,
        sheetName: '季度经营',
      }}
      restoreFocusTarget={() => null}
      onApply={(criteria) => {
        applied.push(criteria);
        return true;
      }}
      onClear={() => false}
      onClose={() => undefined}
    />,
  );

  fireEvent.change(screen.getByRole('combobox', { name: '筛选条件' }), {
    target: { value: 'begins-with' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '筛选值' }), {
    target: { value: '待' },
  });
  fireEvent.click(screen.getByRole('button', { name: '添加第二个条件' }));
  fireEvent.click(screen.getByRole('radio', { name: '或者' }));
  fireEvent.change(screen.getByRole('combobox', { name: '第二个筛选条件' }), {
    target: { value: 'does-not-end-with' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '第二个筛选值' }), {
    target: { value: '归档' },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(applied).toEqual([
    {
      type: 'compound',
      conjunction: 'or',
      conditions: [
        { type: 'begins-with', value: '待' },
        { type: 'does-not-end-with', value: '归档' },
      ],
    },
  ]);
});

test('restores and validates a compound numeric condition', () => {
  const applied: WorkSpreadsheetFilterCriteria[] = [];
  render(
    <SpreadsheetAutoFilterConditionDialog
      source={{
        columnLabel: '收入',
        criteria: {
          type: 'compound',
          conjunction: 'and',
          conditions: [
            { type: 'greater-than', value: '100' },
            { type: 'less-than', value: '200' },
          ],
        },
        hasActiveFilter: true,
        numeric: true,
        sheetName: '季度经营',
      }}
      restoreFocusTarget={() => null}
      onApply={(criteria) => {
        applied.push(criteria);
        return true;
      }}
      onClear={() => false}
      onClose={() => undefined}
    />,
  );

  expect(screen.getByRole('combobox', { name: '筛选条件' })).toHaveValue(
    'greater-than',
  );
  expect(screen.getByRole('radio', { name: '并且' })).toBeChecked();
  expect(screen.getByRole('combobox', { name: '第二个筛选条件' })).toHaveValue(
    'less-than',
  );
  expect(screen.getByRole('textbox', { name: '第二个筛选值' })).toHaveValue(
    '200',
  );

  fireEvent.change(screen.getByRole('textbox', { name: '第二个筛选值' }), {
    target: { value: 'not-a-number' },
  });
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();
  fireEvent.change(screen.getByRole('textbox', { name: '第二个筛选值' }), {
    target: { value: ' 180 ' },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(applied).toEqual([
    {
      type: 'compound',
      conjunction: 'and',
      conditions: [
        { type: 'greater-than', value: '100' },
        { type: 'less-than', value: '180' },
      ],
    },
  ]);
});

test('authors bounded Top/Bottom item and percentage filters for numeric columns', () => {
  const applied: WorkSpreadsheetFilterCriteria[] = [];
  const { unmount } = render(
    <SpreadsheetAutoFilterConditionDialog
      source={{
        columnLabel: '收入',
        criteria: { type: 'top-percent', percent: 25 },
        hasActiveFilter: true,
        numeric: true,
        sheetName: '季度经营',
      }}
      restoreFocusTarget={() => null}
      onApply={(criteria) => {
        applied.push(criteria);
        return true;
      }}
      onClear={() => false}
      onClose={() => undefined}
    />,
  );

  expect(screen.getByRole('combobox', { name: '筛选条件' })).toHaveValue(
    'top-percent',
  );
  expect(screen.getByRole('textbox', { name: '百分比' })).toHaveValue('25');
  fireEvent.change(screen.getByRole('textbox', { name: '百分比' }), {
    target: { value: '101' },
  });
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();

  fireEvent.change(screen.getByRole('combobox', { name: '筛选条件' }), {
    target: { value: 'bottom' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '项目数' }), {
    target: { value: ' 2 ' },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));
  expect(applied).toEqual([{ type: 'bottom', count: 2 }]);

  unmount();
  render(
    <SpreadsheetAutoFilterConditionDialog
      source={{
        columnLabel: '状态',
        criteria: null,
        hasActiveFilter: false,
        numeric: false,
        sheetName: '季度经营',
      }}
      restoreFocusTarget={() => null}
      onApply={() => false}
      onClear={() => false}
      onClose={() => undefined}
    />,
  );
  expect(screen.queryByRole('option', { name: '前几项' })).toBeNull();
});
