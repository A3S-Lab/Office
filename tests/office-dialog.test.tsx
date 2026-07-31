import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useRef, useState } from 'react';
import { Dialog } from '../src/internal/design-system/primitives';
import { useOfficeDialog } from '../src/internal/features/work/editors/office-dialog';
import {
  DOCUMENT_LINK_VALIDATION_MESSAGE,
  normalizeDocumentHref,
} from '../src/internal/features/work/work-document-links';

test('validates prompt input, traps focus, and restores the editing target', async () => {
  render(<OfficeDialogHarness />);

  fireEvent.click(screen.getByRole('button', { name: '打开链接对话框' }));

  const dialog = screen.getByRole('dialog', { name: '添加链接' });
  const input = screen.getByRole('textbox', { name: '链接地址' });
  const submit = screen.getByRole('button', { name: '应用链接' });
  expect(
    screen
      .getByRole('button', { name: '打开删除确认' })
      .closest<HTMLElement>('[inert]'),
  ).toBeInTheDocument();
  expect(input.closest('.ds-dialog-backdrop')).not.toHaveAttribute('inert');
  expect(dialog).toHaveAccessibleDescription('为选中的文字设置跳转地址。');
  expect(input).toHaveFocus();
  expect(input).not.toHaveAttribute('aria-invalid');
  expect(submit).toBeDisabled();
  expect(screen.queryByRole('alert')).not.toBeInTheDocument();

  fireEvent.blur(input);
  expect(input).toHaveAttribute('aria-invalid', 'true');
  expect(screen.getByRole('alert')).toHaveTextContent(
    DOCUMENT_LINK_VALIDATION_MESSAGE,
  );

  input.focus();
  fireEvent.change(input, { target: { value: 'https://a3s.dev/docs' } });
  expect(submit).toBeEnabled();
  fireEvent.keyDown(input, { key: 'Tab' });
  expect(screen.getByRole('button', { name: '取消' })).toHaveFocus();

  fireEvent.click(submit);
  await waitFor(() =>
    expect(screen.getByRole('status')).toHaveTextContent(
      'https://a3s.dev/docs',
    ),
  );
  await waitFor(() =>
    expect(screen.getByRole('textbox', { name: '文档正文' })).toHaveFocus(),
  );
});

test('uses a safe default action and explicit styling for destructive confirms', async () => {
  render(<OfficeDialogHarness />);

  const trigger = screen.getByRole('button', { name: '打开删除确认' });
  trigger.focus();
  fireEvent.click(trigger);

  expect(screen.getByRole('button', { name: '取消' })).toHaveFocus();
  const remove = screen.getByRole('button', { name: '删除' });
  expect(remove).toHaveClass('danger');
  fireEvent.keyDown(screen.getByRole('dialog'), { key: 'Escape' });

  expect(screen.queryByRole('dialog')).not.toBeInTheDocument();
  await waitFor(() => expect(trigger).toHaveFocus());
  expect(screen.getByRole('status')).toHaveTextContent('保留');
});

test('keeps portal dialogs inside their Office theme boundary', async () => {
  function Fixture() {
    const [open, setOpen] = useState(false);
    return (
      <section data-a3s-office data-theme="dark">
        <button type="button" onClick={() => setOpen(true)}>
          打开主题弹窗
        </button>
        {open && (
          <Dialog title="主题弹窗" onClose={() => setOpen(false)}>
            <p>主题内容</p>
          </Dialog>
        )}
      </section>
    );
  }

  render(<Fixture />);
  const trigger = screen.getByRole('button', { name: '打开主题弹窗' });
  trigger.focus();
  fireEvent.click(trigger);

  const dialog = screen.getByRole('dialog', { name: '主题弹窗' });
  expect(dialog.closest('[data-a3s-office]')).toHaveAttribute(
    'data-theme',
    'dark',
  );
  expect(trigger).toHaveAttribute('inert');
  expect(dialog.closest('.ds-dialog-backdrop')).not.toHaveAttribute('inert');
  fireEvent.keyDown(dialog, { key: 'Escape' });
  await waitFor(() => expect(trigger).toHaveFocus());
});

function OfficeDialogHarness() {
  const dialog = useOfficeDialog();
  const editorRef = useRef<HTMLTextAreaElement>(null);
  const [result, setResult] = useState('');

  return (
    <>
      <button
        type="button"
        onClick={() => {
          void dialog
            .prompt({
              title: '添加链接',
              description: '为选中的文字设置跳转地址。',
              fieldLabel: '链接地址',
              initialValue: 'https://',
              confirmLabel: '应用链接',
              required: true,
              validate: (value) =>
                normalizeDocumentHref(value)
                  ? null
                  : DOCUMENT_LINK_VALIDATION_MESSAGE,
              restoreFocusTarget: () => editorRef.current,
            })
            .then((value) => setResult(value ?? '取消'));
        }}
      >
        打开链接对话框
      </button>
      <button
        type="button"
        onClick={() => {
          void dialog
            .confirm({
              title: '删除内容？',
              confirmLabel: '删除',
              confirmTone: 'danger',
            })
            .then((confirmed) => setResult(confirmed ? '删除' : '保留'));
        }}
      >
        打开删除确认
      </button>
      <textarea ref={editorRef} aria-label="文档正文" />
      <output>{result}</output>
      {dialog.dialog}
    </>
  );
}
