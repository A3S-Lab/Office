import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, within } from '@testing-library/react';
import { useRef, useState } from 'react';
import { SpreadsheetFormatCellsDialog } from '../src/internal/features/work/editors/spreadsheet-format-cells-dialog';
import {
  createSpreadsheetFormatCellsDialogSource,
  type SpreadsheetFormatCellsDialogSource,
} from '../src/internal/features/work/editors/spreadsheet-format-cells-dialog-model';
import type { SpreadsheetCellFormatPatch } from '../src/internal/features/work/editors/spreadsheet-cell-format';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('applies explicitly touched settings from every Format Cells tab', () => {
  const patches: SpreadsheetCellFormatPatch[] = [];
  const closes: string[] = [];
  render(
    <SpreadsheetFormatCellsDialog
      source={mixedSource()}
      restoreFocusTarget={() => null}
      onApply={(patch) => {
        patches.push(patch);
        return true;
      }}
      onClose={() => closes.push('close')}
    />,
  );

  const tabs = screen.getByRole('tablist', { name: '单元格格式分类' });
  expect(within(tabs).getAllByRole('tab')).toHaveLength(6);
  expect(screen.getByRole('tab', { name: '数字' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  fireEvent.change(screen.getByRole('textbox', { name: '数字格式代码' }), {
    target: { value: '0.000' },
  });

  fireEvent.click(screen.getByRole('tab', { name: '对齐' }));
  fireEvent.click(screen.getByRole('checkbox', { name: '自动换行' }));

  fireEvent.click(screen.getByRole('tab', { name: '字体' }));
  const bold = screen.getByRole('checkbox', { name: '加粗' });
  expect(bold).toHaveAttribute('aria-checked', 'mixed');
  fireEvent.click(bold);

  fireEvent.click(screen.getByRole('tab', { name: '边框' }));
  fireEvent.click(screen.getByRole('button', { name: '上框线' }));

  fireEvent.click(screen.getByRole('tab', { name: '填充' }));
  fireEvent.click(screen.getByRole('button', { name: '无填充' }));

  fireEvent.click(screen.getByRole('tab', { name: '保护' }));
  fireEvent.click(screen.getByRole('checkbox', { name: '隐藏公式' }));
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  expect(patches).toEqual([
    {
      numberFormat: '0.000',
      wrapText: true,
      bold: false,
      borders: [{ target: 'top', color: '#172033', style: 'thin' }],
      fillColor: null,
      hidden: true,
    },
  ]);
  expect(patches[0]).not.toHaveProperty('horizontalAlignment');
  expect(closes).toEqual(['close']);
});

test('cancels without changes and restores the exact trigger', () => {
  render(<FormatCellsHarness source={mixedSource()} />);
  const trigger = screen.getByRole('button', { name: '打开单元格格式' });
  trigger.focus();
  fireEvent.click(trigger);

  const dialog = screen.getByRole('dialog', { name: '设置单元格格式' });
  fireEvent.keyDown(dialog, { key: 'Escape' });

  expect(screen.queryByRole('dialog', { name: '设置单元格格式' })).toBeNull();
  expect(screen.getByTestId('format-cells-apply-count')).toHaveTextContent('0');
  expect(trigger).toHaveFocus();
});

test('blocks Apply while a custom number code is invalid', () => {
  render(
    <SpreadsheetFormatCellsDialog
      source={mixedSource()}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onClose={() => undefined}
    />,
  );

  fireEvent.change(screen.getByRole('textbox', { name: '数字格式代码' }), {
    target: { value: ' ' },
  });

  expect(screen.getByRole('alert')).toHaveTextContent('请输入数字格式代码。');
  expect(screen.getByRole('button', { name: '应用' })).toBeDisabled();
});

test('explains an imported font size that blocks Apply', () => {
  const source = mixedSource();
  render(
    <SpreadsheetFormatCellsDialog
      source={{
        ...source,
        fields: {
          ...source.fields,
          fontSize: { value: 500, mixed: false },
        },
      }}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onClose={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('tab', { name: '字体' }));
  expect(screen.getByRole('alert')).toHaveTextContent(
    '字号需为 1–409 之间的数字。',
  );
  expect(screen.getByRole('button', { name: '应用' })).toBeDisabled();
});

function FormatCellsHarness({
  source,
}: {
  source: SpreadsheetFormatCellsDialogSource;
}) {
  const [open, setOpen] = useState(false);
  const [applyCount, setApplyCount] = useState(0);
  const triggerRef = useRef<HTMLButtonElement>(null);
  return (
    <>
      <button ref={triggerRef} type="button" onClick={() => setOpen(true)}>
        打开单元格格式
      </button>
      <output data-testid="format-cells-apply-count">{applyCount}</output>
      {open && (
        <SpreadsheetFormatCellsDialog
          source={source}
          restoreFocusTarget={() => triggerRef.current}
          onApply={() => {
            setApplyCount((current) => current + 1);
            return true;
          }}
          onClose={() => setOpen(false)}
        />
      )}
    </>
  );
}

function mixedSource(): SpreadsheetFormatCellsDialogSource {
  const content = {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        data: [
          [
            {
              v: 12,
              ct: { fa: '#,##0', t: 'n' },
              ht: '1',
              bl: 1,
              bg: '#fff2cc',
            },
            {
              v: 0.25,
              ct: { fa: '0.00%', t: 'n' },
              ht: '2',
              bl: 0,
            },
          ],
        ],
      },
    ],
  } satisfies WorkSpreadsheetContent;
  const source = createSpreadsheetFormatCellsDialogSource(
    content,
    'sheet-1',
    { row: [0, 0], column: [0, 1] },
    content.sheets[0]?.data ?? [],
    { row: 0, column: 0 },
  );
  if (!source) throw new Error('Expected a Format Cells source.');
  return source;
}
