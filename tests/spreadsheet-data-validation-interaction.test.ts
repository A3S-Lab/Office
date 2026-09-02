import { describe, expect, test } from '@rstest/core';
import type { WorkSpreadsheetDataValidationItem } from '../src/internal/features/work/work-types';
import {
  cloneSpreadsheetDataValidationItem,
  spreadsheetDataValidationConfirmLabels,
  spreadsheetDataValidationDialogDescription,
  spreadsheetDataValidationDialogTitle,
  spreadsheetDataValidationErrorStyle,
  spreadsheetDataValidationInteraction,
  spreadsheetDataValidationValueText,
} from '../src/internal/features/work/editors/spreadsheet-data-validation-interaction';

describe('spreadsheet data-validation interaction policy', () => {
  test('maps Office error styles to the right browser interaction', () => {
    expect(spreadsheetDataValidationErrorStyle({ errorStyle: 'stop' })).toBe(
      'stop',
    );
    expect(spreadsheetDataValidationErrorStyle({ errorStyle: 'warning' })).toBe(
      'warning',
    );
    expect(
      spreadsheetDataValidationErrorStyle({ errorStyle: 'information' }),
    ).toBe('information');
    expect(spreadsheetDataValidationErrorStyle({})).toBe('stop');

    expect(spreadsheetDataValidationInteraction('stop')).toBe('notice');
    expect(spreadsheetDataValidationInteraction('warning')).toBe('confirm');
    expect(spreadsheetDataValidationInteraction('information')).toBe('confirm');
  });

  test('keeps custom copy readable without repeating the title', () => {
    const item = {
      errorStyle: 'information',
      errorTitle: 'Date outside 2026',
      errorMessage: 'Enter a date in calendar year 2026.',
    } satisfies Pick<
      WorkSpreadsheetDataValidationItem,
      'errorStyle' | 'errorTitle' | 'errorMessage'
    >;

    expect(spreadsheetDataValidationDialogTitle(item)).toBe(
      'Date outside 2026',
    );
    expect(
      spreadsheetDataValidationDialogDescription(
        'Date outside 2026\nEnter a date in calendar year 2026.',
        '2027-01-01',
        item,
      ),
    ).toBe('Enter a date in calendar year 2026.\n\n当前输入：2027-01-01');
    expect(
      spreadsheetDataValidationDialogDescription('', '', {
        errorStyle: 'stop',
      }),
    ).toBe('当前输入不符合此单元格的数据验证规则。');
  });

  test('uses Traditional Office confirmation labels and safe value text', () => {
    expect(spreadsheetDataValidationConfirmLabels('warning')).toEqual({
      cancelLabel: '取消',
      confirmLabel: '继续输入',
    });
    expect(spreadsheetDataValidationConfirmLabels('information')).toEqual({
      cancelLabel: '返回修改',
      confirmLabel: '保留输入',
    });
    expect(spreadsheetDataValidationValueText('  invalid  ')).toBe('invalid');
    expect(spreadsheetDataValidationValueText(0)).toBe('0');
    expect(spreadsheetDataValidationValueText(false)).toBe('false');
    expect(spreadsheetDataValidationValueText({ value: 'invalid' })).toBe('');
  });

  test('clones validation metadata before an asynchronous dialog', () => {
    const source = {
      type: 'dropdown',
      type2: '',
      rangeTxt: 'A1',
      value1: 'Ready',
      value2: '',
      validity: '',
      remote: false,
      prohibitInput: true,
      errorStyle: 'warning',
      errorTitle: 'Invalid state',
      errorMessage: 'Choose Ready.',
      hintShow: false,
      hintValue: '',
    } satisfies WorkSpreadsheetDataValidationItem;
    const clone = cloneSpreadsheetDataValidationItem(source);
    expect(clone).toEqual(source);
    expect(clone).not.toBe(source);
  });
});
