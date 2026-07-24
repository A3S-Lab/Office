import { Editor } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { createDocumentPageChromeEditorExtensions } from '../src/internal/features/work/editors/document-page-chrome-editor';
import { DocumentPageChromeRibbon } from '../src/internal/features/work/editors/document-page-chrome-ribbon';

test('uses typed commands and explicit navigation in the page-chrome ribbon', async () => {
  const editor = new Editor({
    extensions: createDocumentPageChromeEditorExtensions(),
    content: '<p>Quarterly header</p>',
  });
  editor.commands.setTextSelection({ from: 1, to: 17 });
  const parts: string[] = [];
  let pageNumberToggles = 0;
  let closes = 0;

  render(
    <DocumentPageChromeRibbon
      editor={editor}
      editingPart="header"
      showPageNumber={false}
      onEditingPartChange={(part) => parts.push(part)}
      onTogglePageNumber={() => {
        pageNumberToggles += 1;
      }}
      onClose={() => {
        closes += 1;
      }}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '页眉页脚加粗' }));
  fireEvent.click(screen.getByRole('button', { name: '页眉页脚居中' }));
  await waitFor(() => {
    expect(editor.getHTML()).toContain('<strong>Quarterly header</strong>');
    expect(editor.getHTML()).toContain('text-align: center');
  });
  expect(screen.getByRole('button', { name: '页眉页脚加粗' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  expect(screen.getByRole('button', { name: '页眉页脚居中' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );

  fireEvent.click(screen.getByRole('button', { name: '切换到页脚' }));
  fireEvent.click(screen.getByRole('button', { name: '显示页码' }));
  fireEvent.click(screen.getByRole('button', { name: '关闭页眉和页脚' }));

  expect(parts).toEqual(['footer']);
  expect(pageNumberToggles).toBe(1);
  expect(closes).toBe(1);
  editor.destroy();
});
