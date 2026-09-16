import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
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

test('cancels a dirty Create Table draft before Escape closes the dialog', async () => {
  const source = {
    name: 'Table1',
    range: { row: [0, 2], column: [0, 2] },
    rangeReference: 'A1:C3',
    sheetId: 'sheet-1',
    sheetName: 'Sales',
    value: { headerRow: true, rangeReference: 'A1:C3', totalsRow: false },
  } as const;
  const closes: string[] = [];

  function Harness() {
    const [open, setOpen] = useState(true);
    return (
      <>
        <button type="button" data-testid="create-table-dialog-launcher">
          创建表格
        </button>
        {open ? (
          <SpreadsheetTableDialog
            source={source}
            restoreFocusTarget={() =>
              document.querySelector<HTMLElement>(
                '[data-testid="create-table-dialog-launcher"]',
              )
            }
            onApply={() => true}
            onClose={() => {
              closes.push('close');
              setOpen(false);
            }}
            onValidate={() => null}
          />
        ) : null}
      </>
    );
  }

  render(<Harness />);
  const range = screen.getByRole('textbox', { name: '表格区域' });
  fireEvent.change(range, { target: { value: 'A1:D10' } });
  expect(range).toHaveValue('A1:D10');

  fireEvent.keyDown(range, { key: 'Escape' });
  expect(screen.getByRole('dialog', { name: '创建表格' })).toBeInTheDocument();
  expect(range).toHaveValue('A1:C3');
  expect(closes).toEqual([]);

  fireEvent.keyDown(range, { key: 'Escape' });
  expect(screen.queryByRole('dialog', { name: '创建表格' })).toBeNull();
  expect(closes).toEqual(['close']);
  await waitFor(() =>
    expect(screen.getByTestId('create-table-dialog-launcher')).toHaveFocus(),
  );
});
