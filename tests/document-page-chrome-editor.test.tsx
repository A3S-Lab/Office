import { Editor } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import {
  applyDocumentPageChromeEditorCommand,
  createDocumentPageChromeEditorExtensions,
  DocumentPageChromeRichTextEditor,
  documentPageChromeEditorState,
  normalizeDocumentPageChromeHref,
} from '../src/internal/features/work/editors/document-page-chrome-editor';
import { sanitizeDocumentPageChromeHtml } from '../src/internal/features/work/work-document-page-chrome';

const pixelPng =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC' +
  'AAAAC0lEQVR42mP8/x8AAusB9Y9Z9WQAAAAASUVORK5CYII=';

test('applies typed page-chrome commands to a TipTap document', () => {
  const editor = new Editor({
    extensions: createDocumentPageChromeEditorExtensions(),
    content: '<p>Quarterly header</p>',
  });
  editor.commands.setTextSelection({ from: 1, to: 10 });

  expect(
    applyDocumentPageChromeEditorCommand(editor, { type: 'toggleBold' }),
  ).toBe(true);
  expect(
    applyDocumentPageChromeEditorCommand(editor, {
      type: 'toggleUnderline',
    }),
  ).toBe(true);
  expect(
    applyDocumentPageChromeEditorCommand(editor, {
      type: 'setColor',
      color: '#175cd3',
    }),
  ).toBe(true);
  expect(
    applyDocumentPageChromeEditorCommand(editor, {
      type: 'setLink',
      href: 'https://a3s.dev/office',
    }),
  ).toBe(true);
  expect(
    applyDocumentPageChromeEditorCommand(editor, {
      type: 'setAlignment',
      alignment: 'right',
    }),
  ).toBe(true);

  const state = documentPageChromeEditorState(editor);
  const html = sanitizeDocumentPageChromeHtml(editor.getHTML());
  expect(state).toMatchObject({
    alignment: 'right',
    bold: true,
    color: '#175cd3',
    link: 'https://a3s.dev/office',
    underline: true,
  });
  expect(html).toContain('text-align: right');
  expect(html).toContain('color: #175cd3');
  expect(html).toContain('href="https://a3s.dev/office"');
  expect(html).toContain('<strong>');
  expect(html).toContain('<u>');

  expect(
    applyDocumentPageChromeEditorCommand(editor, {
      type: 'setLink',
      href: 'javascript:alert(1)',
    }),
  ).toBe(false);
  editor.commands.setTextSelection(editor.state.doc.content.size);
  expect(
    applyDocumentPageChromeEditorCommand(editor, {
      type: 'insertImage',
      alt: 'A3S mark',
      source: pixelPng,
    }),
  ).toBe(true);
  expect(editor.getHTML()).toContain('alt="A3S mark"');
  expect(editor.getHTML()).toContain('data:image/png;base64');
  editor.destroy();
});

test('keeps the page-chrome surface controlled and exposes active formatting', async () => {
  const changes: string[] = [];
  let editor: Editor | null = null;
  const view = render(
    <DocumentPageChromeRichTextEditor
      label="默认页眉"
      value="<p>Page title</p>"
      onChange={(html) => changes.push(html)}
      onEditorChange={(current) => {
        editor = current;
      }}
    />,
  );

  const textbox = await screen.findByRole('textbox', { name: '默认页眉' });
  expect(textbox).toHaveAttribute('data-document-page-chrome-engine', 'tiptap');
  expect(textbox).toHaveClass('work-document-page-chrome-content');
  expect(editor).not.toBeNull();
  const current = editor as Editor;
  current.commands.setTextSelection({ from: 1, to: 5 });
  fireEvent.click(screen.getByRole('button', { name: '默认页眉加粗' }));

  await waitFor(() => {
    expect(changes.at(-1)).toContain('<strong>Page</strong>');
  });
  expect(screen.getByRole('button', { name: '默认页眉加粗' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );

  view.rerender(
    <DocumentPageChromeRichTextEditor
      label="默认页眉"
      value="<p>Externally updated</p>"
      onChange={(html) => changes.push(html)}
      onEditorChange={(next) => {
        editor = next;
      }}
    />,
  );
  await waitFor(() => {
    expect(screen.getByRole('textbox', { name: '默认页眉' })).toHaveTextContent(
      'Externally updated',
    );
  });
  view.unmount();
});

test('supports a toolbar-free page surface and exits with Escape', async () => {
  let exits = 0;
  render(
    <DocumentPageChromeRichTextEditor
      autoFocus
      label="页内页眉"
      value="<p>Direct page editing</p>"
      showToolbar={false}
      onChange={() => undefined}
      onExit={() => {
        exits += 1;
      }}
    />,
  );

  const textbox = await screen.findByRole('textbox', { name: '页内页眉' });
  expect(
    screen.queryByRole('toolbar', { name: '页内页眉格式' }),
  ).not.toBeInTheDocument();
  await waitFor(() => expect(document.activeElement).toBe(textbox));

  fireEvent.keyDown(textbox, { key: 'Escape' });
  expect(exits).toBe(1);
});

test('rejects unsafe links and unsupported header images without native UI', async () => {
  expect(normalizeDocumentPageChromeHref(' javascript:alert(1) ')).toBeNull();
  expect(normalizeDocumentPageChromeHref(' #summary ')).toBe('#summary');
  expect(normalizeDocumentPageChromeHref(' mailto:office@a3s.dev ')).toBe(
    'mailto:office@a3s.dev',
  );

  render(
    <DocumentPageChromeRichTextEditor
      label="默认页脚"
      value=""
      onChange={() => undefined}
    />,
  );
  const input = await screen.findByLabelText('默认页脚图片文件');
  fireEvent.change(input, {
    target: {
      files: [new File(['not an image'], 'notes.txt', { type: 'text/plain' })],
    },
  });

  expect(
    await screen.findByRole('heading', { name: '无法插入图片' }),
  ).toBeInTheDocument();
  expect(screen.getByRole('button', { name: '知道了' })).toBeInTheDocument();
});
