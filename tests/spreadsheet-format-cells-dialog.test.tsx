import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { useRef, useState } from 'react';
import { SpreadsheetFormatCellsDialog } from '../src/internal/features/work/editors/spreadsheet-format-cells-dialog';
import {
  createSpreadsheetFormatCellsDialogSource,
  type SpreadsheetFormatCellsDialogSource,
} from '../src/internal/features/work/editors/spreadsheet-format-cells-dialog-model';
import type { SpreadsheetCellFormatPatch } from '../src/internal/features/work/editors/spreadsheet-cell-format';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';
import { withXlsxGradientFill } from '../src/internal/features/work/work-xlsx-gradient-fill';

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
  fireEvent.click(screen.getByRole('combobox', { name: '下划线样式' }));
  fireEvent.click(screen.getByRole('option', { name: '双下划线' }));

  fireEvent.click(screen.getByRole('tab', { name: '边框' }));
  fireEvent.click(screen.getByRole('button', { name: '上框线' }));
  fireEvent.click(screen.getByRole('button', { name: '斜下框线' }));
  fireEvent.click(screen.getByRole('button', { name: '斜上框线' }));

  fireEvent.click(screen.getByRole('tab', { name: '填充' }));
  fireEvent.click(screen.getByRole('radio', { name: '无填充' }));

  fireEvent.click(screen.getByRole('tab', { name: '保护' }));
  fireEvent.click(screen.getByRole('checkbox', { name: '隐藏公式' }));
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  expect(patches).toEqual([
    {
      numberFormat: '0.000',
      wrapText: true,
      bold: false,
      underline: 'double',
      borders: [
        { target: 'top', color: '#172033', style: 'thin' },
        { target: 'diagonalDown', color: '#172033', style: 'thin' },
        { target: 'diagonalUp', color: '#172033', style: 'thin' },
      ],
      fill: { kind: 'none' },
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

test('opens the Font tab on the shortcut-specific family or size control', async () => {
  const lightFontSource = mixedSource();
  lightFontSource.fields.fontColor = {
    value: '#FFFFFF',
    mixed: false,
  };
  const first = render(
    <SpreadsheetFormatCellsDialog
      source={lightFontSource}
      openIntent={{ tab: 'font', focus: 'fontFamily' }}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onClose={() => undefined}
    />,
  );

  expect(screen.getByRole('tab', { name: '字体' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  const family = screen.getByRole('combobox', { name: '单元格字体' });
  expect(family).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+F Meta+Shift+F',
  );
  expect(screen.getByText('A3S Office 字体预览')).toHaveStyle({
    backgroundColor: '#172033',
    color: '#FFFFFF',
  });
  await waitFor(() => expect(family).toHaveFocus());
  first.unmount();

  render(
    <SpreadsheetFormatCellsDialog
      source={mixedSource()}
      openIntent={{ tab: 'font', focus: 'fontSize' }}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onClose={() => undefined}
    />,
  );
  const size = screen.getByRole('combobox', { name: '单元格字号' });
  expect(size).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+P Meta+Shift+P',
  );
  await waitFor(() => expect(size).toHaveFocus());
});

test('authors a multi-stop linear gradient as one typed Fill patch', () => {
  const patches: SpreadsheetCellFormatPatch[] = [];
  render(
    <SpreadsheetFormatCellsDialog
      source={mixedSource()}
      openIntent={{ tab: 'fill' }}
      restoreFocusTarget={() => null}
      onApply={(patch) => {
        patches.push(patch);
        return true;
      }}
      onClose={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('radio', { name: '渐变' }));
  fireEvent.change(screen.getByRole('textbox', { name: '线性渐变角度' }), {
    target: { value: '45' },
  });
  fireEvent.click(screen.getByRole('button', { name: '添加色标' }));
  expect(screen.getAllByRole('button', { name: /色标 \d+ 颜色/ })).toHaveLength(
    3,
  );
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  expect(patches).toEqual([
    {
      fill: {
        kind: 'gradient',
        value: {
          degree: 45,
          stops: [
            { color: '#fff2cc', position: 0 },
            { color: '#fff9e6', position: 0.5 },
            { color: '#ffffff', position: 1 },
          ],
          type: 'linear',
        },
      },
    },
  ]);
});

test('edits imported path geometry and invalidates only the changed stop color origin', () => {
  const patches: SpreadsheetCellFormatPatch[] = [];
  render(
    <SpreadsheetFormatCellsDialog
      source={gradientSource()}
      openIntent={{ tab: 'fill' }}
      restoreFocusTarget={() => null}
      onApply={(patch) => {
        patches.push(patch);
        return true;
      }}
      onClose={() => undefined}
    />,
  );

  expect(screen.getByRole('radio', { name: '渐变' })).toBeChecked();
  expect(screen.getByRole('combobox', { name: '渐变类型' })).toHaveTextContent(
    '路径',
  );
  fireEvent.change(screen.getByRole('textbox', { name: '路径渐变左边界' }), {
    target: { value: '20' },
  });
  fireEvent.click(screen.getByRole('button', { name: '色标 1 颜色' }));
  fireEvent.click(screen.getByRole('option', { name: '颜色 #6d9eeb' }));
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  expect(patches).toEqual([
    {
      fill: {
        kind: 'gradient',
        value: {
          bottom: 0.8,
          left: 0.2,
          right: 0.75,
          stops: [
            { color: '#6d9eeb', position: 0 },
            {
              color: '#ffffff',
              colorOrigin: {
                baseColor: '#ffffff',
                index: 1,
                kind: 'theme',
                renderedColor: '#ffffff',
              },
              position: 1,
            },
          ],
          top: 0.2,
          type: 'path',
        },
      },
    },
  ]);
});

test('authors a native pattern and blocks crossed path geometry', () => {
  const patches: SpreadsheetCellFormatPatch[] = [];
  const first = render(
    <SpreadsheetFormatCellsDialog
      source={mixedSource()}
      openIntent={{ tab: 'fill' }}
      restoreFocusTarget={() => null}
      onApply={(patch) => {
        patches.push(patch);
        return true;
      }}
      onClose={() => undefined}
    />,
  );
  fireEvent.click(screen.getByRole('radio', { name: '图案' }));
  fireEvent.click(screen.getByRole('combobox', { name: '填充图案样式' }));
  fireEvent.click(screen.getByRole('option', { name: '深色网格' }));
  fireEvent.click(screen.getByRole('button', { name: '应用' }));
  expect(patches).toEqual([
    {
      fill: {
        kind: 'pattern',
        value: {
          backgroundColor: '#fff2cc',
          foregroundColor: '#ffffff',
          patternType: 'darkGrid',
        },
      },
    },
  ]);
  first.unmount();

  render(
    <SpreadsheetFormatCellsDialog
      source={gradientSource()}
      openIntent={{ tab: 'fill' }}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onClose={() => undefined}
    />,
  );
  fireEvent.change(screen.getByRole('textbox', { name: '路径渐变左边界' }), {
    target: { value: '90' },
  });
  expect(screen.getByRole('alert')).toHaveTextContent('请检查填充设置');
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

function gradientSource(): SpreadsheetFormatCellsDialogSource {
  const themeStart = {
    baseColor: '#4472c4',
    index: 4,
    kind: 'theme',
    renderedColor: '#4472c4',
  } as const;
  const themeEnd = {
    baseColor: '#ffffff',
    index: 1,
    kind: 'theme',
    renderedColor: '#ffffff',
  } as const;
  const content = {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        data: [
          [
            withXlsxGradientFill(
              { bg: '#4472c4', v: 'Gradient' },
              {
                bottom: 0.8,
                left: 0.25,
                right: 0.75,
                stops: [
                  { color: '#4472c4', colorOrigin: themeStart, position: 0 },
                  { color: '#ffffff', colorOrigin: themeEnd, position: 1 },
                ],
                top: 0.2,
                type: 'path',
              },
            ),
          ],
        ],
      },
    ],
  } satisfies WorkSpreadsheetContent;
  const source = createSpreadsheetFormatCellsDialogSource(
    content,
    'sheet-1',
    { row: [0, 0], column: [0, 0] },
    content.sheets[0]?.data ?? [],
    { row: 0, column: 0 },
  );
  if (!source) throw new Error('Expected a gradient Format Cells source.');
  return source;
}
