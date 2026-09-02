import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { useRef, useState } from 'react';
import { SpreadsheetDataValidationDialog } from '../src/internal/features/work/editors/spreadsheet-data-validation-dialog';
import {
  type SpreadsheetDataValidationDialogSource,
  type SpreadsheetDataValidationDialogValue,
  spreadsheetDataValidationFailureMessage,
} from '../src/internal/features/work/editors/spreadsheet-data-validation';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('applies an accessible dropdown rule to every selected range', () => {
  const source = dialogSource();
  const values: SpreadsheetDataValidationDialogValue[] = [];
  render(
    <SpreadsheetDataValidationDialog
      source={source}
      restoreFocusTarget={() => null}
      onApply={(value) => {
        values.push(value);
        return true;
      }}
      onClose={() => undefined}
      onRemove={() => false}
      onValidate={(value) => validationMessage(source, value)}
    />,
  );

  const dialog = screen.getByRole('dialog', { name: '数据验证' });
  expect(dialog).toHaveTextContent('Inputs!B2:B3,D5:E5');
  expect(dialog).toHaveTextContent('2 个选定区域');
  expect(screen.getByRole('combobox', { name: '允许' })).toHaveValue(
    'dropdown',
  );
  expect(
    screen.getByRole('checkbox', { name: '输入无效数据时显示错误警告' }),
  ).toBeChecked();
  expect(screen.getByRole('checkbox', { name: '忽略空值' })).toBeChecked();
  expect(
    screen.getByRole('checkbox', { name: '在单元格内显示下拉箭头' }),
  ).toBeChecked();
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();

  const listSource = screen.getByRole('textbox', { name: '来源' });
  fireEvent.change(listSource, { target: { value: 'Ready,,Blocked' } });
  expect(screen.getByRole('alert')).toHaveTextContent(
    '请输入不超过 255 个字符的逗号分隔列表',
  );
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();

  fireEvent.change(listSource, { target: { value: 'Ready,Blocked' } });
  fireEvent.click(
    screen.getByRole('checkbox', { name: '选中单元格时显示输入信息' }),
  );
  fireEvent.change(screen.getByRole('textbox', { name: '输入信息标题' }), {
    target: { value: 'Workflow state' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '输入信息' }), {
    target: { value: 'Choose a workflow state.' },
  });
  fireEvent.change(screen.getByRole('combobox', { name: '错误警告样式' }), {
    target: { value: 'warning' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '错误警告标题' }), {
    target: { value: 'Invalid state' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '错误警告消息' }), {
    target: { value: 'Choose Ready or Blocked.' },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(values).toEqual([
    {
      type: 'dropdown',
      type2: '',
      value1: 'Ready,Blocked',
      value2: '',
      allowBlank: true,
      showDropdownArrow: true,
      prohibitInput: true,
      errorStyle: 'warning',
      errorTitle: 'Invalid state',
      errorMessage: 'Choose Ready or Blocked.',
      hintShow: true,
      hintTitle: 'Workflow state',
      hintValue: 'Choose a workflow state.',
    },
  ]);
});

test('switches to date validation and normalizes its visible error state', () => {
  const source = dialogSource();
  render(
    <SpreadsheetDataValidationDialog
      source={source}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onClose={() => undefined}
      onRemove={() => false}
      onValidate={(value) => validationMessage(source, value)}
    />,
  );

  fireEvent.change(screen.getByRole('combobox', { name: '允许' }), {
    target: { value: 'date' },
  });
  expect(screen.getByRole('combobox', { name: '数据' })).toHaveValue('between');
  fireEvent.change(screen.getByRole('textbox', { name: '开始日期' }), {
    target: { value: '2026-13-01' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '结束日期' }), {
    target: { value: '2026-12-31' },
  });
  expect(screen.getByRole('alert')).toHaveTextContent('请输入有效的日期');

  fireEvent.change(screen.getByRole('textbox', { name: '开始日期' }), {
    target: { value: 'DATE(2026,8,21)' },
  });
  expect(screen.queryByRole('alert')).toBeNull();
  expect(screen.getByRole('button', { name: '确定' })).toBeEnabled();
});

test('authors an accessible custom formula without exposing numeric operators', () => {
  const source = dialogSource();
  const values: SpreadsheetDataValidationDialogValue[] = [];
  render(
    <SpreadsheetDataValidationDialog
      source={source}
      restoreFocusTarget={() => null}
      onApply={(value) => {
        values.push(value);
        return true;
      }}
      onClose={() => undefined}
      onRemove={() => false}
      onValidate={(value) => validationMessage(source, value)}
    />,
  );

  fireEvent.change(screen.getByRole('combobox', { name: '允许' }), {
    target: { value: 'custom' },
  });
  expect(screen.queryByRole('combobox', { name: '数据' })).toBeNull();
  expect(screen.getByRole('textbox', { name: '公式' })).toBeVisible();
  fireEvent.change(screen.getByRole('textbox', { name: '公式' }), {
    target: { value: '=AND(A1<>"",A1<=100)' },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(values[0]).toMatchObject({
    type: 'custom',
    type2: '',
    value1: '=AND(A1<>"",A1<=100)',
    value2: '',
  });
});

test('clears existing rules and restores the exact trigger', () => {
  render(<DataValidationDialogHarness />);
  const trigger = screen.getByRole('button', { name: '打开数据验证' });
  trigger.focus();
  fireEvent.click(trigger);

  expect(screen.getByRole('dialog', { name: '数据验证' })).toBeVisible();
  fireEvent.click(screen.getByRole('button', { name: '全部清除' }));

  expect(screen.queryByRole('dialog', { name: '数据验证' })).toBeNull();
  expect(screen.getByTestId('data-validation-remove-count')).toHaveTextContent(
    '1',
  );
  expect(trigger).toHaveFocus();
});

function DataValidationDialogHarness() {
  const [open, setOpen] = useState(false);
  const [removeCount, setRemoveCount] = useState(0);
  const triggerRef = useRef<HTMLButtonElement>(null);
  return (
    <>
      <button ref={triggerRef} type="button" onClick={() => setOpen(true)}>
        打开数据验证
      </button>
      <output data-testid="data-validation-remove-count">{removeCount}</output>
      {open && (
        <SpreadsheetDataValidationDialog
          source={{ ...dialogSource(), hasValidation: true }}
          restoreFocusTarget={() => triggerRef.current}
          onApply={() => true}
          onClose={() => setOpen(false)}
          onRemove={() => {
            setRemoveCount((current) => current + 1);
            return true;
          }}
          onValidate={() => null}
        />
      )}
    </>
  );
}

function validationMessage(
  source: SpreadsheetDataValidationDialogSource,
  value: SpreadsheetDataValidationDialogValue,
): string | null {
  return spreadsheetDataValidationFailureMessage(dialogContent(), {
    sheetId: source.sheetId,
    ranges: source.ranges,
    activeCell: source.activeCell,
    value,
  });
}

function dialogSource(): SpreadsheetDataValidationDialogSource {
  return {
    sheetId: 'sheet-1',
    sheetName: 'Inputs',
    ranges: [
      { row: [1, 2], column: [1, 1] },
      { row: [4, 4], column: [3, 4] },
    ],
    activeCell: { row: 1, column: 1 },
    rangeReference: 'B2:B3,D5:E5',
    hasValidation: false,
    mixed: false,
    value: {
      type: 'dropdown',
      type2: '',
      value1: '',
      value2: '',
      allowBlank: true,
      showDropdownArrow: true,
      prohibitInput: true,
      errorStyle: 'stop',
      errorTitle: '',
      errorMessage: '',
      hintShow: false,
      hintTitle: '',
      hintValue: '',
    },
  };
}

function dialogContent(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Inputs',
        row: 40,
        column: 12,
      },
    ],
  };
}
