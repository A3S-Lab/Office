import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { SpreadsheetPasteSpecialDialog } from '../src/internal/features/work/editors/spreadsheet-paste-special-dialog';
import type {
  SpreadsheetClipboardSnapshot,
  SpreadsheetPasteSpecialOptions,
} from '../src/internal/features/work/editors/spreadsheet-paste-special';

test('submits WPS paste content, arithmetic, and layout options', () => {
  const applied: SpreadsheetPasteSpecialOptions[] = [];
  let closeCount = 0;
  render(
    <SpreadsheetPasteSpecialDialog
      source={dialogSource(richSnapshot())}
      restoreFocusTarget={() => null}
      onApply={(options) => {
        applied.push(options);
        return true;
      }}
      onClose={() => {
        closeCount += 1;
      }}
      onValidate={() => null}
    />,
  );

  expect(screen.getByRole('dialog', { name: '选择性粘贴' })).toBeVisible();
  expect(screen.getByLabelText('剪贴板摘要')).toHaveTextContent(
    'A3S 富剪贴板2 行 × 2 列',
  );
  fireEvent.click(screen.getByRole('radio', { name: '值' }));
  fireEvent.click(screen.getByRole('radio', { name: '加' }));
  fireEvent.click(screen.getByRole('checkbox', { name: '跳过空白单元格' }));
  fireEvent.click(screen.getByRole('checkbox', { name: '转置行列' }));
  fireEvent.click(screen.getByRole('button', { name: '粘贴' }));

  expect(applied).toEqual([
    {
      content: 'values',
      operation: 'add',
      skipBlanks: true,
      transpose: true,
    },
  ]);
  expect(closeCount).toBe(1);
});

test('disables rich-only modes for an external text clipboard', () => {
  render(
    <SpreadsheetPasteSpecialDialog
      source={dialogSource(textSnapshot())}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onClose={() => undefined}
      onValidate={() => null}
    />,
  );

  expect(screen.getByRole('radio', { name: /^格式/ })).toBeDisabled();
  expect(screen.getByRole('radio', { name: /^数据验证/ })).toBeDisabled();
  expect(screen.getByRole('radio', { name: /^列宽/ })).toBeDisabled();
  expect(screen.getByRole('radio', { name: '值' })).toBeEnabled();
  expect(screen.getByLabelText('剪贴板摘要')).toHaveTextContent(
    '纯文本1 行 × 2 列',
  );
});

test('resets incompatible options for column widths and exposes validation errors', () => {
  const validation = (options: SpreadsheetPasteSpecialOptions) =>
    options.content === 'column-widths' ? '目标列受保护。' : null;
  render(
    <SpreadsheetPasteSpecialDialog
      source={dialogSource(richSnapshot())}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onClose={() => undefined}
      onValidate={validation}
    />,
  );

  fireEvent.click(screen.getByRole('checkbox', { name: '转置行列' }));
  fireEvent.click(screen.getByRole('radio', { name: '列宽' }));

  expect(screen.getByRole('checkbox', { name: '转置行列' })).toBeDisabled();
  expect(screen.getByRole('checkbox', { name: '转置行列' })).not.toBeChecked();
  expect(screen.getByRole('radio', { name: '加' })).toBeDisabled();
  expect(screen.getByRole('alert')).toHaveTextContent('目标列受保护。');
  expect(screen.getByRole('button', { name: '粘贴' })).toBeDisabled();
});

function dialogSource(snapshot: SpreadsheetClipboardSnapshot) {
  return {
    snapshot,
    targetSheetId: 'sheet-1',
    targetRange: { row: [2, 2], column: [2, 2] } as {
      row: [number, number];
      column: [number, number];
    },
    invoker: null,
  };
}

function richSnapshot(): SpreadsheetClipboardSnapshot {
  return {
    version: 1,
    kind: 'rich',
    plainText: '1\t2\n3\t4',
    sourceSheetId: 'sheet-1',
    sourceRange: { row: [0, 1], column: [0, 1] },
    rowCount: 2,
    columnCount: 2,
    cells: [
      [clipboardCell(1), clipboardCell(2)],
      [clipboardCell(3), clipboardCell(4)],
    ],
    columnWidths: [96, 96],
    merges: [],
    containsUnsupportedFormulaState: false,
  };
}

function textSnapshot(): SpreadsheetClipboardSnapshot {
  return {
    version: 1,
    kind: 'text',
    plainText: '1\t2',
    sourceRange: { row: [0, 0], column: [0, 1] },
    rowCount: 1,
    columnCount: 2,
    cells: [[clipboardCell('1'), clipboardCell('2')]],
    merges: [],
    containsUnsupportedFormulaState: false,
  };
}

function clipboardCell(value: number | string) {
  return { cell: { v: value, m: String(value) }, borders: {} };
}
