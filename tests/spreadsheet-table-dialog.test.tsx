import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { SpreadsheetTableDialog } from '../src/internal/features/work/editors/spreadsheet-table-dialog';

test('validates an accessible Create Table workflow before applying', () => {
  const applied: Array<{ headerRow: boolean; rangeReference: string }> = [];
  const closed: string[] = [];
  render(
    <SpreadsheetTableDialog
      source={{
        name: 'Table1',
        range: { row: [0, 2], column: [0, 2] },
        rangeReference: 'A1:C3',
        sheetId: 'sheet-1',
        sheetName: 'Sales',
        value: { headerRow: true, rangeReference: 'A1:C3' },
      }}
      restoreFocusTarget={() => null}
      onApply={(value) => {
        applied.push(value);
        return true;
      }}
      onClose={() => closed.push('close')}
      onValidate={(value) =>
        value.rangeReference === 'A1:C3' ? null : '请输入一个连续区域。'
      }
    />,
  );

  const dialog = screen.getByRole('dialog', { name: '创建表格' });
  expect(dialog).toHaveTextContent('Sales');
  expect(screen.getByRole('textbox', { name: '表格区域' })).toHaveValue(
    'A1:C3',
  );
  expect(screen.getByRole('checkbox', { name: '表包含标题' })).toBeChecked();

  fireEvent.change(screen.getByRole('textbox', { name: '表格区域' }), {
    target: { value: 'A1,B2' },
  });
  expect(screen.getByRole('alert')).toHaveTextContent('请输入一个连续区域。');
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();

  fireEvent.change(screen.getByRole('textbox', { name: '表格区域' }), {
    target: { value: 'A1:C3' },
  });
  fireEvent.click(screen.getByRole('checkbox', { name: '表包含标题' }));
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(applied).toEqual([{ headerRow: false, rangeReference: 'A1:C3' }]);
  expect(closed).toEqual(['close']);
});
