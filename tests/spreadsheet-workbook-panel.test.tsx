import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
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
  const close = screen.getByRole('button', { name: '关闭打印设置' });

  expect(panel).toContainElement(body);
  expect(body).not.toContainElement(close);

  fireEvent.keyDown(body, { key: 'Escape' });
  expect(closed).toEqual([true]);
});

test('dismisses an open workbook panel with Escape outside the panel', () => {
  const closed: boolean[] = [];
  render(
    <SpreadsheetWorkbookPanel
      content={content}
      view="charts"
      activeSheetId="sheet-1"
      can={can}
      commands={commands}
      onClose={() => closed.push(true)}
    />,
  );

  fireEvent.keyDown(document, { key: 'Escape' });

  expect(closed).toEqual([true]);
});

test('lets an active overlay consume Escape before the workbook panel', () => {
  const closed: boolean[] = [];
  render(
    <SpreadsheetWorkbookPanel
      content={content}
      view="names"
      activeSheetId="sheet-1"
      can={can}
      commands={commands}
      onClose={() => closed.push(true)}
    />,
  );
  const menu = document.createElement('div');
  menu.setAttribute('role', 'menu');
  document.body.append(menu);

  fireEvent.keyDown(document, { key: 'Escape' });
  menu.remove();
  expect(closed).toEqual([]);

  fireEvent.keyDown(document, { key: 'Escape' });
  expect(closed).toEqual([true]);
});

test('dismisses the workbook panel when an editor stops Escape propagation', async () => {
  const closed: boolean[] = [];
  render(
    <SpreadsheetWorkbookPanel
      content={content}
      view="charts"
      activeSheetId="sheet-1"
      can={can}
      commands={commands}
      onClose={() => closed.push(true)}
    />,
  );
  const editor = document.createElement('div');
  editor.addEventListener('keydown', (event) => event.stopPropagation());
  document.body.append(editor);

  fireEvent.keyDown(editor, { key: 'Escape' });

  await waitFor(() => expect(closed).toEqual([true]));
  editor.remove();
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

test('cancels a dirty formula draft before Escape closes the workbook panel', () => {
  const closed: boolean[] = [];
  render(
    <SpreadsheetWorkbookPanel
      content={{
        ...content,
        calculation: {
          mode: 'automatic',
          fullCalculationOnLoad: false,
          forceFullCalculation: false,
          iterativeCalculation: true,
          maximumIterations: 100,
          maximumChange: 0.001,
          fullPrecision: true,
        },
      }}
      view="formulas"
      activeSheetId="sheet-1"
      can={can}
      commands={commands}
      onClose={() => closed.push(true)}
    />,
  );
  const fullPrecision = screen.getByRole('checkbox', {
    name: '使用完整精度',
  });
  fireEvent.click(fullPrecision);

  fireEvent.keyDown(fullPrecision, { key: 'Escape' });
  expect(fullPrecision).toBeChecked();
  expect(closed).toEqual([]);

  fireEvent.keyDown(fullPrecision, { key: 'Escape' });
  expect(closed).toEqual([true]);
});

test('protects a dirty defined-name draft from object switching', () => {
  const changes: WorkSpreadsheetContent[] = [];
  render(
    <SpreadsheetWorkbookPanel
      content={{
        ...content,
        namedRanges: [
          {
            id: 'name-1',
            name: 'Revenue',
            reference: "'工作表 1'!$A$1:$A$10",
          },
          {
            id: 'name-2',
            name: 'Costs',
            reference: "'工作表 1'!$B$1:$B$10",
          },
        ],
      }}
      view="names"
      activeSheetId="sheet-1"
      can={can}
      commands={{
        ...commands,
        setSpreadsheetContent: (next) => {
          changes.push(next);
          return true;
        },
      }}
      onClose={() => undefined}
    />,
  );
  const name = screen.getByRole('textbox', { name: '名称' });
  fireEvent.change(name, { target: { value: 'RevenueDraft' } });
  fireEvent.click(screen.getByRole('button', { name: /Costs/ }));

  expect(name).toHaveValue('RevenueDraft');
  expect(screen.getByRole('alert')).toHaveTextContent('请先保存或取消');
  expect(changes).toEqual([]);

  fireEvent.click(screen.getByRole('button', { name: '取消更改' }));
  expect(name).toHaveValue('Revenue');
  fireEvent.click(screen.getByRole('button', { name: /Costs/ }));
  expect(name).toHaveValue('Costs');
});

test('keeps a newly saved defined name stable until controlled content catches up', () => {
  const changes: WorkSpreadsheetContent[] = [];
  render(
    <SpreadsheetWorkbookPanel
      content={content}
      view="names"
      activeSheetId="sheet-1"
      can={can}
      commands={{
        ...commands,
        setSpreadsheetContent: (next) => {
          changes.push(next);
          return true;
        },
      }}
      onClose={() => undefined}
    />,
  );
  fireEvent.change(screen.getByRole('textbox', { name: '名称' }), {
    target: { value: 'Revenue' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '名称引用位置' }), {
    target: { value: "'工作表 1'!$A$1:$A$10" },
  });

  fireEvent.click(screen.getByRole('button', { name: '保存名称' }));

  expect(changes).toHaveLength(1);
  expect(screen.getByRole('textbox', { name: '名称' })).toHaveValue('Revenue');
  expect(screen.getByRole('button', { name: '保存名称' })).toBeDisabled();
});
