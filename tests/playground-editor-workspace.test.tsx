import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { useState } from 'react';
import { createArtifact } from '../src/core';
import {
  AssistantPanel,
  EditorExportButton,
} from '../playground/src/editor-workspace';

test('keeps compact export actions named when their text is hidden', () => {
  const view = render(
    <EditorExportButton
      kind="markdown"
      exporting={false}
      onExport={() => undefined}
    />,
  );

  expect(screen.getByRole('button', { name: '导出' })).toHaveAttribute(
    'aria-label',
    '导出',
  );

  view.rerender(
    <EditorExportButton
      kind="pdf"
      exporting={false}
      onExport={() => undefined}
    />,
  );
  expect(screen.getByRole('button', { name: '下载 PDF' })).toHaveAttribute(
    'title',
    '下载 PDF',
  );
});

test('keeps Word export formats behind one compact header action', () => {
  let docxExports = 0;
  let pdfExports = 0;
  render(
    <EditorExportButton
      kind="document"
      exporting={false}
      onExport={() => {
        docxExports += 1;
      }}
      onExportPdf={() => {
        pdfExports += 1;
      }}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '导出' }));
  expect(screen.getByRole('menu', { name: '导出格式' })).toBeVisible();
  fireEvent.click(screen.getByRole('menuitem', { name: '导出 PDF' }));
  expect(pdfExports).toBe(1);
  expect(docxExports).toBe(0);

  fireEvent.click(screen.getByRole('button', { name: '导出' }));
  fireEvent.click(screen.getByRole('menuitem', { name: '下载 DOCX' }));
  expect(docxExports).toBe(1);
});

test('contains compact assistant focus and restores its exact trigger', () => {
  render(<ModalAssistantHarness />);

  const trigger = screen.getByRole('button', { name: '打开 AI 助手' });
  trigger.focus();
  fireEvent.click(trigger);

  const assistant = screen.getByRole('dialog', { name: 'AI 助手' });
  const close = screen.getByRole('button', { name: '关闭 AI 助手' });
  const background = screen.getByRole('button', { name: '编辑器操作' });
  expect(assistant).toHaveAttribute('aria-modal', 'true');
  expect(background.closest<HTMLElement>('[inert]')).toBeInTheDocument();
  expect(close).toHaveFocus();

  fireEvent.keyDown(close, { key: 'Tab' });
  expect(
    screen.getByRole('separator', { name: '调整 AI 助手宽度' }),
  ).toHaveFocus();
  fireEvent.keyDown(
    screen.getByRole('separator', { name: '调整 AI 助手宽度' }),
    { key: 'Tab', shiftKey: true },
  );
  expect(close).toHaveFocus();

  fireEvent.keyDown(close, { key: 'Escape' });
  expect(screen.queryByRole('dialog', { name: 'AI 助手' })).toBeNull();
  expect(trigger).toHaveFocus();
});

function ModalAssistantHarness() {
  const [open, setOpen] = useState(false);
  const artifact = createArtifact('blank-document');
  return (
    <div>
      <button type="button" onClick={() => setOpen(true)}>
        打开 AI 助手
      </button>
      {open && (
        <AssistantPanel
          artifact={artifact}
          lastRequest={null}
          modal
          width={460}
          onClose={() => setOpen(false)}
          onWidthChange={() => undefined}
        />
      )}
      <section>
        <button type="button">编辑器操作</button>
      </section>
    </div>
  );
}
