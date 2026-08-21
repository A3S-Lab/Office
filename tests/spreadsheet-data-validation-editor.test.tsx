import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import type { SpreadsheetContent } from '../src/core';
import { SpreadsheetEditor } from '../src/react';

test('commits data validation once, restores focus, and undoes in one step', async () => {
  const changes: SpreadsheetContent[] = [];
  render(<ControlledSpreadsheet changes={changes} />);

  const dataTab = await waitFor(() =>
    screen.getByRole('tab', { name: '数据' }),
  );
  fireEvent.click(dataTab);
  const launcher = screen.getByRole('button', { name: '数据验证' });
  launcher.focus();
  fireEvent.click(launcher);
  expect(screen.getByRole('dialog', { name: '数据验证' })).toHaveTextContent(
    'Inputs!A1',
  );
  fireEvent.click(screen.getByRole('button', { name: '取消' }));
  await waitFor(() => expect(launcher).toHaveFocus());
  expect(changes).toHaveLength(0);

  fireEvent.click(launcher);
  fireEvent.change(screen.getByRole('textbox', { name: '来源' }), {
    target: { value: 'Ready,Blocked' },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  await waitFor(() =>
    expect(screen.getByTestId('data-validation-state')).toHaveTextContent(
      'Ready,Blocked',
    ),
  );
  expect(screen.getByTestId('data-validation-state')).toHaveTextContent(
    '"rangeTxt":"A1"',
  );
  expect(screen.getByTestId('data-validation-selection')).toHaveTextContent(
    '0:0',
  );
  expect(changes).toHaveLength(1);
  await waitFor(() => expect(launcher).toHaveFocus());

  const undo = screen.getByRole('button', { name: '撤销' });
  await waitFor(() => expect(undo).toBeEnabled());
  fireEvent.click(undo);

  await waitFor(() =>
    expect(screen.getByTestId('data-validation-state')).toHaveTextContent(
      'none',
    ),
  );
  expect(screen.getByTestId('data-validation-selection')).toHaveTextContent(
    '0:0',
  );
  expect(changes).toHaveLength(2);
  await waitFor(() => expect(undo).toBeDisabled());
});

function ControlledSpreadsheet({ changes }: { changes: SpreadsheetContent[] }) {
  const [content, setContent] = useState<SpreadsheetContent>(() => ({
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Inputs',
        status: 1,
        row: 4,
        column: 4,
        data: [[{ v: 'State', m: 'State', bg: '#fff2cc', bl: 1 }]],
        luckysheet_select_save: [
          {
            row: [0, 0],
            column: [0, 0],
            row_focus: 0,
            column_focus: 0,
          },
        ],
      },
    ],
  }));
  const sheet = content.sheets[0];
  const selection = sheet?.luckysheet_select_save?.at(-1);
  return (
    <>
      <SpreadsheetEditor
        content={content}
        onChange={(next) => {
          changes.push(next);
          setContent(next);
        }}
        theme="light"
      />
      <output data-testid="data-validation-state">
        {sheet?.dataValidationRanges?.length
          ? JSON.stringify(sheet.dataValidationRanges)
          : 'none'}
      </output>
      <output data-testid="data-validation-selection">
        {selection
          ? `${selection.row[0] ?? 0}:${selection.column[0] ?? 0}`
          : 'none'}
      </output>
    </>
  );
}
