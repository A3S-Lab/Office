import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { SpreadsheetProtectionPanel } from '../src/internal/features/work/editors/spreadsheet-protection-panel';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

const content: WorkSpreadsheetContent = {
  type: 'spreadsheet',
  sheets: [
    {
      id: 'sheet-1',
      name: '工作表 1',
      status: 1,
      row: 30,
      column: 10,
      data: [],
      config: {
        authority: {
          allowRangeList: [
            { name: '输入区', sqref: 'B2:B10' },
            { name: '备注区', sqref: 'D2:D10' },
          ],
        },
      },
    },
  ],
};

test('keeps a dirty editable-range draft while protection settings update', () => {
  const changes: WorkSpreadsheetContent[] = [];
  const { rerender } = render(
    <SpreadsheetProtectionPanel
      content={content}
      onChange={(next) => changes.push(next)}
    />,
  );
  fireEvent.click(screen.getByRole('button', { name: /输入区/ }));
  const name = screen.getByRole('textbox', { name: '可编辑区域名称' });
  fireEvent.change(name, { target: { value: '尚未保存的输入区' } });

  fireEvent.click(screen.getByRole('button', { name: /备注区/ }));
  expect(name).toHaveValue('尚未保存的输入区');
  expect(screen.getByRole('alert')).toHaveTextContent('请先保存或取消');

  fireEvent.click(screen.getByRole('checkbox', { name: '启用工作表保护' }));
  expect(changes).toHaveLength(1);
  rerender(
    <SpreadsheetProtectionPanel
      content={changes[0]}
      onChange={(next) => changes.push(next)}
    />,
  );
  expect(name).toHaveValue('尚未保存的输入区');

  fireEvent.click(screen.getByRole('button', { name: '取消更改' }));
  expect(name).toHaveValue('输入区');
  fireEvent.click(screen.getByRole('button', { name: /备注区/ }));
  expect(name).toHaveValue('备注区');
});

test('keeps a newly saved editable range stable until controlled content catches up', () => {
  const changes: WorkSpreadsheetContent[] = [];
  render(
    <SpreadsheetProtectionPanel
      content={{
        type: 'spreadsheet',
        sheets: [
          {
            id: 'sheet-1',
            name: '工作表 1',
            status: 1,
            row: 30,
            column: 10,
            data: [],
          },
        ],
      }}
      onChange={(next) => changes.push(next)}
    />,
  );
  fireEvent.change(screen.getByRole('textbox', { name: '可编辑区域名称' }), {
    target: { value: 'InputCells' },
  });

  fireEvent.click(screen.getByRole('button', { name: '保存区域' }));

  expect(changes).toHaveLength(1);
  expect(screen.getByRole('textbox', { name: '可编辑区域名称' })).toHaveValue(
    'InputCells',
  );
  expect(screen.getByRole('button', { name: '保存区域' })).toBeDisabled();
});
