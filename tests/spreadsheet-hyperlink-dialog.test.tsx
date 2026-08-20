import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { useRef, useState } from 'react';
import { SpreadsheetHyperlinkDialog } from '../src/internal/features/work/editors/spreadsheet-hyperlink-dialog';
import {
  createSpreadsheetHyperlinkDialogSource,
  type SpreadsheetHyperlinkDialogSource,
  type SpreadsheetHyperlinkDialogValue,
  validateSpreadsheetHyperlinkRequest,
} from '../src/internal/features/work/editors/spreadsheet-hyperlink';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('inserts a validated webpage hyperlink from the accessible dialog', () => {
  const content = dialogContent();
  const source = createSpreadsheetHyperlinkDialogSource(content, {
    sheetId: 'sheet-1',
    row: 0,
    column: 0,
  });
  if (!source) throw new Error('Expected a hyperlink dialog source.');
  const values: SpreadsheetHyperlinkDialogValue[] = [];

  render(
    <SpreadsheetHyperlinkDialog
      source={source}
      restoreFocusTarget={() => null}
      onApply={(value) => {
        values.push(value);
        return true;
      }}
      onClose={() => undefined}
      onRemove={() => false}
      onValidate={(value) => {
        const result = validateSpreadsheetHyperlinkRequest(content, {
          sheetId: source.sheetId,
          row: source.row,
          column: source.column,
          ...value,
        });
        return result.ok ? null : result.message;
      }}
    />,
  );

  expect(
    screen.getByRole('dialog', { name: '插入超链接' }),
  ).toBeInTheDocument();
  expect(screen.getByRole('radiogroup', { name: '链接类型' })).toBeVisible();
  expect(screen.getByRole('radio', { name: '网页' })).toBeChecked();
  expect(screen.getByRole('radio', { name: '单元格区域' })).not.toBeChecked();
  expect(screen.getByRole('radio', { name: '工作表' })).not.toBeChecked();
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();

  fireEvent.change(screen.getByRole('textbox', { name: '显示文本' }), {
    target: { value: 'A3S Office docs' },
  });
  const address = screen.getByRole('textbox', { name: '地址' });
  fireEvent.change(address, { target: { value: 'javascript:alert(1)' } });
  expect(screen.getByRole('alert')).toHaveTextContent(
    '请输入有效的 HTTP 或 HTTPS 地址。',
  );
  expect(screen.getByRole('button', { name: '确定' })).toBeDisabled();

  fireEvent.change(address, { target: { value: 'a3s.dev/office' } });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(values).toEqual([
    {
      linkType: 'webpage',
      linkAddress: 'a3s.dev/office',
      displayText: 'A3S Office docs',
    },
  ]);
});

test('switches between native cell-range and worksheet targets', () => {
  const source = blankDialogSource();
  const values: SpreadsheetHyperlinkDialogValue[] = [];
  render(
    <SpreadsheetHyperlinkDialog
      source={source}
      restoreFocusTarget={() => null}
      onApply={(value) => {
        values.push(value);
        return true;
      }}
      onClose={() => undefined}
      onRemove={() => false}
      onValidate={() => null}
    />,
  );

  fireEvent.click(screen.getByRole('radio', { name: '单元格区域' }));
  fireEvent.change(screen.getByRole('textbox', { name: '单元格或区域' }), {
    target: { value: "'Archive 2025'!C9:E12" },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));
  expect(values[0]).toEqual({
    linkType: 'cellrange',
    linkAddress: "'Archive 2025'!C9:E12",
    displayText: "'Archive 2025'!C9:E12",
  });

  values.length = 0;
  fireEvent.click(screen.getByRole('radio', { name: '工作表' }));
  fireEvent.change(screen.getByRole('combobox', { name: '工作表' }), {
    target: { value: 'Archive 2025' },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));
  expect(values[0]).toEqual({
    linkType: 'sheet',
    linkAddress: 'Archive 2025',
    displayText: 'Archive 2025',
  });
});

test('edits and removes an existing link while restoring the exact trigger', () => {
  render(<HyperlinkDialogHarness source={existingDialogSource()} />);
  const trigger = screen.getByRole('button', { name: '打开超链接' });
  trigger.focus();
  fireEvent.click(trigger);

  expect(
    screen.getByRole('dialog', { name: '编辑超链接' }),
  ).toBeInTheDocument();
  expect(screen.getByRole('textbox', { name: '显示文本' })).toBeDisabled();
  expect(
    screen.getByText('公式单元格的显示文本由公式结果决定。'),
  ).toBeVisible();
  fireEvent.click(screen.getByRole('button', { name: '移除超链接' }));

  expect(screen.queryByRole('dialog', { name: '编辑超链接' })).toBeNull();
  expect(screen.getByTestId('hyperlink-remove-count')).toHaveTextContent('1');
  expect(trigger).toHaveFocus();
});

function HyperlinkDialogHarness({
  source,
}: {
  source: SpreadsheetHyperlinkDialogSource;
}) {
  const [open, setOpen] = useState(false);
  const [removeCount, setRemoveCount] = useState(0);
  const triggerRef = useRef<HTMLButtonElement>(null);
  return (
    <>
      <button ref={triggerRef} type="button" onClick={() => setOpen(true)}>
        打开超链接
      </button>
      <output data-testid="hyperlink-remove-count">{removeCount}</output>
      {open && (
        <SpreadsheetHyperlinkDialog
          source={source}
          restoreFocusTarget={() => triggerRef.current}
          onApply={() => true}
          onClose={() => setOpen(false)}
          onRemove={() => {
            setRemoveCount((current) => current + 1);
            return true;
          }}
          onValidate={() => null}
        />
      )}
    </>
  );
}

function blankDialogSource(): SpreadsheetHyperlinkDialogSource {
  return {
    sheetId: 'sheet-1',
    sheetName: 'Sheet 1',
    row: 0,
    column: 0,
    cellReference: 'A1',
    displayText: '',
    displayTextEditable: true,
    hasHyperlink: false,
    link: null,
    sheetOptions: [
      { id: 'sheet-1', name: 'Sheet 1' },
      { id: 'archive', name: 'Archive 2025' },
    ],
  };
}

function existingDialogSource(): SpreadsheetHyperlinkDialogSource {
  return {
    ...blankDialogSource(),
    row: 1,
    column: 1,
    cellReference: 'B2',
    displayText: 'A3S Office',
    displayTextEditable: false,
    hasHyperlink: true,
    link: { linkType: 'webpage', linkAddress: 'https://a3s.dev' },
  };
}

function dialogContent(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        row: 40,
        column: 12,
        status: 1,
      },
      {
        id: 'archive',
        name: 'Archive 2025',
        row: 100,
        column: 20,
      },
    ],
  };
}
