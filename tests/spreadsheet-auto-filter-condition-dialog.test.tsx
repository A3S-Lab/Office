import { expect, test, afterEach } from '@rstest/core';
import { cleanup, fireEvent, render, screen } from '@testing-library/react';

afterEach(() => {
  cleanup();
});

function chooseOfficeSelectOption(
  ariaLabel: string,
  optionName: string | RegExp,
) {
  fireEvent.click(screen.getByRole('combobox', { name: ariaLabel }));
  fireEvent.click(screen.getByRole('option', { name: optionName }));
}

function chooseOfficeSelectOptionByValue(ariaLabel: string, value: string) {
  fireEvent.click(screen.getByRole('combobox', { name: ariaLabel }));
  fireEvent.click(
    document.querySelector(
      `.work-office-select-menu [role='option'][data-value='${value}']`,
    ) as HTMLElement,
  );
}

function expectOfficeSelectValue(ariaLabel: string, value: string) {
  expect(screen.getByRole('combobox', { name: ariaLabel })).toHaveAttribute(
    'data-selected-value',
    value,
  );
}

import { SpreadsheetAutoFilterConditionDialog } from '../src/internal/features/work/editors/spreadsheet-auto-filter-condition-dialog';
import type { WorkSpreadsheetFilterCriteria } from '../src/internal/features/work/work-types';

test('validates and applies a bounded numeric condition', () => {
  const applied: WorkSpreadsheetFilterCriteria[] = [];
  render(
    <SpreadsheetAutoFilterConditionDialog
      source={{
        columnLabel: '收入',
        criteria: null,
        date: false,
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
  chooseOfficeSelectOption('筛选条件', '介于');
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
        date: false,
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

  expectOfficeSelectValue('筛选条件', 'contains');
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
        date: false,
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

  chooseOfficeSelectOption('筛选条件', '开头是');
  fireEvent.change(screen.getByRole('textbox', { name: '筛选值' }), {
    target: { value: '待' },
  });
  fireEvent.click(screen.getByRole('button', { name: '添加第二个条件' }));
  fireEvent.click(screen.getByRole('radio', { name: '或者' }));
  chooseOfficeSelectOptionByValue('第二个筛选条件', 'does-not-end-with');
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

test('restores and authors WPS wildcard expressions in compound conditions', () => {
  const applied: WorkSpreadsheetFilterCriteria[] = [];
  render(
    <SpreadsheetAutoFilterConditionDialog
      source={{
        columnLabel: '名称',
        criteria: { type: 'matches-wildcard', value: 'K?ng*' },
        date: false,
        hasActiveFilter: true,
        numeric: false,
        sheetName: '客户清单',
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

  expectOfficeSelectValue('筛选条件', 'matches-wildcard');
  expect(screen.getByRole('textbox', { name: '通配符表达式' })).toHaveValue(
    'K?ng*',
  );
  expect(screen.getByText(/\* 匹配任意多个字符/)).toBeVisible();
  fireEvent.change(screen.getByRole('textbox', { name: '通配符表达式' }), {
    target: { value: '*'.repeat(32_768) },
  });
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();
  fireEvent.change(screen.getByRole('textbox', { name: '通配符表达式' }), {
    target: { value: 'K?ng*' },
  });
  fireEvent.click(screen.getByRole('button', { name: '添加第二个条件' }));
  fireEvent.click(screen.getByRole('radio', { name: '并且' }));
  chooseOfficeSelectOption('第二个筛选条件', '通配符不匹配');
  fireEvent.change(screen.getByRole('textbox', { name: '第二个筛选值' }), {
    target: { value: 'King~*' },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(applied).toEqual([
    {
      type: 'compound',
      conjunction: 'and',
      conditions: [
        { type: 'matches-wildcard', value: 'K?ng*' },
        { type: 'does-not-match-wildcard', value: 'King~*' },
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
        date: false,
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

  expectOfficeSelectValue('筛选条件', 'greater-than');
  expect(screen.getByRole('radio', { name: '并且' })).toBeChecked();
  expectOfficeSelectValue('第二个筛选条件', 'less-than');
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
        date: false,
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

  expectOfficeSelectValue('筛选条件', 'top-percent');
  expect(screen.getByRole('textbox', { name: '百分比' })).toHaveValue('25');
  fireEvent.change(screen.getByRole('textbox', { name: '百分比' }), {
    target: { value: '101' },
  });
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();

  chooseOfficeSelectOption('筛选条件', '后几项');
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
        date: false,
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
  fireEvent.click(screen.getByRole('combobox', { name: '筛选条件' }));
  expect(screen.queryByRole('option', { name: '前几项' })).toBeNull();
  fireEvent.click(screen.getByRole('combobox', { name: '筛选条件' }));
});

test('restores and authors value-free dynamic date conditions for date columns', () => {
  const applied: WorkSpreadsheetFilterCriteria[] = [];
  render(
    <SpreadsheetAutoFilterConditionDialog
      source={{
        columnLabel: '交付日期',
        criteria: { type: 'dynamic', kind: 'this-month' },
        date: true,
        hasActiveFilter: true,
        numeric: false,
        sheetName: '项目计划',
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

  expectOfficeSelectValue('筛选条件', 'this-month');
  fireEvent.click(screen.getByRole('combobox', { name: '筛选条件' }));
  expect(screen.getByRole('option', { name: '今天' })).toBeVisible();
  expect(screen.getByRole('option', { name: '六月' })).toBeVisible();
  fireEvent.click(screen.getByRole('combobox', { name: '筛选条件' }));
  expect(screen.queryByRole('textbox')).toBeNull();
  expect(screen.queryByRole('button', { name: '添加第二个条件' })).toBeNull();

  chooseOfficeSelectOptionByValue('筛选条件', 'year-to-date');
  fireEvent.click(screen.getByRole('button', { name: '确定' }));
  expect(applied).toEqual([{ type: 'dynamic', kind: 'year-to-date' }]);
});

test('offers and restores strict average conditions for numeric columns', () => {
  const applied: WorkSpreadsheetFilterCriteria[] = [];
  render(
    <SpreadsheetAutoFilterConditionDialog
      source={{
        columnLabel: '收入',
        criteria: { type: 'dynamic', kind: 'above-average' },
        date: false,
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

  expectOfficeSelectValue('筛选条件', 'above-average');
  fireEvent.click(screen.getByRole('combobox', { name: '筛选条件' }));
  expect(screen.getByRole('option', { name: '低于平均值' })).toBeVisible();
  expect(screen.queryByRole('option', { name: '今天' })).toBeNull();
  fireEvent.click(screen.getByRole('combobox', { name: '筛选条件' }));
  expect(screen.queryByRole('textbox')).toBeNull();
  chooseOfficeSelectOption('筛选条件', '低于平均值');
  fireEvent.click(screen.getByRole('button', { name: '确定' }));
  expect(applied).toEqual([{ type: 'dynamic', kind: 'below-average' }]);
});
