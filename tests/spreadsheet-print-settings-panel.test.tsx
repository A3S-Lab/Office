import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { SpreadsheetPrintSettingsPanel } from '../src/internal/features/work/editors/spreadsheet-print-settings-panel';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

const content: WorkSpreadsheetContent = {
  type: 'spreadsheet',
  sheets: [{ id: 'sheet-1', name: '工作表 1', status: 1, data: [] }],
};

test('keeps incomplete print measurements out of the page setup draft', () => {
  const changes: WorkSpreadsheetContent[] = [];
  render(
    <SpreadsheetPrintSettingsPanel
      content={content}
      onChange={(next) => changes.push(next)}
    />,
  );

  const scale = screen.getByRole('textbox', { name: '缩放比例' });
  fireEvent.change(scale, { target: { value: '' } });
  expect(scale).toHaveValue('');
  expect(changes).toEqual([]);
  fireEvent.blur(scale);
  expect(scale).toHaveValue('100');

  fireEvent.change(scale, { target: { value: '125.' } });
  fireEvent.keyDown(scale, { key: 'Enter' });
  expect(scale).toHaveValue('125');
  expect(changes).toEqual([]);

  const topMargin = screen.getByRole('textbox', {
    name: '上边距（毫米）',
  });
  fireEvent.change(topMargin, { target: { value: '18.456' } });
  fireEvent.blur(topMargin);
  expect(topMargin).toHaveValue('18.46');

  fireEvent.click(screen.getByRole('button', { name: '保存打印设置' }));
  expect(changes).toHaveLength(1);
  expect(changes[0].pageSetups?.[0]).toMatchObject({
    scale: 125,
    margins: expect.objectContaining({ top: 18.46 }),
  });
});

test('preserves dirty print settings and guards worksheet switching', () => {
  const twoSheetContent: WorkSpreadsheetContent = {
    type: 'spreadsheet',
    sheets: [
      { id: 'sheet-1', name: '工作表 1', status: 1, data: [] },
      { id: 'sheet-2', name: '工作表 2', status: 0, data: [] },
    ],
    printAreas: [
      { sheetId: 'sheet-1', reference: '$A$1:$A$10' },
      { sheetId: 'sheet-2', reference: '$B$1:$B$10' },
    ],
  };
  const props = { onChange: () => undefined };
  const { rerender } = render(
    <SpreadsheetPrintSettingsPanel content={twoSheetContent} {...props} />,
  );
  const reference = screen.getByRole('textbox', { name: '打印范围' });
  fireEvent.change(reference, { target: { value: 'A1:A20' } });
  rerender(
    <SpreadsheetPrintSettingsPanel
      content={{
        ...twoSheetContent,
        sheets: [
          { ...twoSheetContent.sheets[0], data: [[{ v: '已更新' }]] },
          twoSheetContent.sheets[1],
        ],
      }}
      {...props}
    />,
  );
  expect(reference).toHaveValue('A1:A20');

  fireEvent.click(screen.getByRole('combobox', { name: '打印设置工作表' }));
  fireEvent.click(screen.getByRole('option', { name: '工作表 2' }));
  expect(reference).toHaveValue('A1:A20');
  expect(screen.getByRole('alert')).toHaveTextContent('请先保存或取消');

  fireEvent.click(screen.getByRole('button', { name: '取消更改' }));
  expect(reference).toHaveValue('$A$1:$A$10');
  fireEvent.click(screen.getByRole('combobox', { name: '打印设置工作表' }));
  fireEvent.click(screen.getByRole('option', { name: '工作表 2' }));
  expect(reference).toHaveValue('$B$1:$B$10');
});
