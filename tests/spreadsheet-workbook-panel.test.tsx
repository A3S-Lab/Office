import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { SpreadsheetWorkbookPanel } from '../src/internal/features/work/editors/spreadsheet-workbook-panel';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

const content: WorkSpreadsheetContent = {
  type: 'spreadsheet',
  sheets: [{ id: 'sheet-1', name: '工作表 1', status: 1, data: [] }],
};

const can = {
  recalculateFormula: () => true,
};

const commands = {
  recalculateFormula: () => true,
  setSpreadsheetContent: () => true,
};

test('keeps workbook panel controls outside its scrollable body', () => {
  const closed: boolean[] = [];
  render(
    <SpreadsheetWorkbookPanel
      content={content}
      view="print-area"
      activeSheetId="sheet-1"
      can={can}
      commands={commands}
      onClose={() => closed.push(true)}
    />,
  );

  const panel = screen.getByRole('region', { name: '打印设置' });
  const body = screen.getByRole('region', { name: '打印设置内容' });
  const close = screen.getByRole('button', { name: '关闭工作簿设置' });

  expect(panel).toContainElement(body);
  expect(body).not.toContainElement(close);

  fireEvent.keyDown(body, { key: 'Escape' });
  expect(closed).toEqual([true]);
});

test('starts each workbook panel view at the top of a fresh scroll body', () => {
  const props = {
    content,
    activeSheetId: 'sheet-1',
    can,
    commands,
    onClose: () => undefined,
  };
  const { rerender } = render(
    <SpreadsheetWorkbookPanel {...props} view="print-area" />,
  );

  const printBody = screen.getByRole('region', { name: '打印设置内容' });
  printBody.scrollTop = 120;

  rerender(<SpreadsheetWorkbookPanel {...props} view="names" />);

  const namesBody = screen.getByRole('region', { name: '名称管理器内容' });
  expect(namesBody).not.toBe(printBody);
  expect(namesBody.scrollTop).toBe(0);
});
