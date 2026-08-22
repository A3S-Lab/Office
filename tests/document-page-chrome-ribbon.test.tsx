import { Editor } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
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
  expect(screen.getByRole('button', { name: '页眉页脚加粗' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+B Meta+B',
  );
  expect(screen.getByRole('button', { name: '页眉页脚下标' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+= Meta+=',
  );
  expect(screen.getByRole('button', { name: '页眉页脚上标' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+= Meta+Shift+=',
  );
  expect(screen.getByRole('button', { name: '页眉页脚居中' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+E Meta+E',
  );
  fireEvent.click(screen.getByRole('button', { name: '页眉页脚斜体' }));
  fireEvent.click(screen.getByRole('button', { name: '页眉页脚下划线' }));
  fireEvent.click(screen.getByRole('button', { name: '更多页眉页脚下划线' }));
  const underlineMenu = await screen.findByRole('menu', {
    name: '页眉页脚下划线样式',
  });
  fireEvent.click(
    within(underlineMenu).getByRole('menuitemradio', {
      name: '双下划线',
    }),
  );
  expect(editor.getAttributes('underline').underlineStyle).toBe('double');
  for (const label of [
    '页眉页脚左对齐',
    '页眉页脚居中',
    '页眉页脚右对齐',
    '页眉页脚两端对齐',
  ]) {
    fireEvent.click(screen.getByRole('button', { name: label }));
  }
  fireEvent.click(screen.getByRole('button', { name: '页眉页脚上标' }));
  fireEvent.click(screen.getByRole('button', { name: '页眉页脚文字颜色' }));
  fireEvent.click(screen.getByRole('option', { name: '颜色 #0070c0' }));
  await waitFor(() => {
    expect(editor.getHTML()).toContain('<strong>');
    expect(editor.getHTML()).toContain('<em>');
    expect(editor.getHTML()).toContain('data-office-underline-style="double"');
    expect(editor.getHTML()).toContain('text-align: justify');
    expect(editor.getHTML()).toContain('color: #0070c0');
    expect(editor.getHTML()).toContain('<sup>');
  });
  expect(screen.getByRole('button', { name: '页眉页脚加粗' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  expect(
    screen.getByRole('button', { name: '页眉页脚两端对齐' }),
  ).toHaveAttribute('aria-pressed', 'true');
  expect(screen.getByRole('button', { name: '页眉页脚上标' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );

  fireEvent.click(screen.getByRole('button', { name: '页眉页脚下标' }));
  await waitFor(() => {
    expect(editor.getHTML()).toContain('<sub>');
    expect(editor.getHTML()).not.toContain('<sup>');
  });
  expect(screen.getByRole('button', { name: '页眉页脚下标' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  expect(screen.getByRole('button', { name: '页眉页脚上标' })).toHaveAttribute(
    'aria-pressed',
    'false',
  );

  const beforeUndo = editor.getHTML();
  fireEvent.click(screen.getByRole('button', { name: '撤销页眉页脚编辑' }));
  await waitFor(() => expect(editor.getHTML()).not.toBe(beforeUndo));
  fireEvent.click(screen.getByRole('button', { name: '重做页眉页脚编辑' }));
  await waitFor(() => expect(editor.getHTML()).toBe(beforeUndo));

  fireEvent.click(screen.getByRole('button', { name: '添加页眉页脚链接' }));
  const linkDialog = await screen.findByRole('dialog', { name: '添加链接' });
  fireEvent.change(
    within(linkDialog).getByRole('textbox', { name: '链接地址' }),
    {
      target: { value: 'https://a3s.dev/header' },
    },
  );
  fireEvent.click(within(linkDialog).getByRole('button', { name: '添加链接' }));
  await waitFor(() =>
    expect(editor.getHTML()).toContain('href="https://a3s.dev/header"'),
  );
  fireEvent.click(
    await screen.findByRole('button', { name: '移除页眉页脚链接' }),
  );
  await waitFor(() => expect(editor.getHTML()).not.toContain('href='));

  const imageInput = screen.getByLabelText('页眉页脚图片文件');
  let imagePickerRequests = 0;
  imageInput.addEventListener('click', () => {
    imagePickerRequests += 1;
  });
  fireEvent.click(screen.getByRole('button', { name: '插入页眉页脚图片' }));
  expect(imagePickerRequests).toBe(1);

  fireEvent.click(screen.getByRole('button', { name: '切换到页脚' }));
  fireEvent.click(screen.getByRole('button', { name: '切换到页眉' }));
  fireEvent.click(screen.getByRole('button', { name: '显示页码' }));
  fireEvent.click(screen.getByRole('button', { name: '关闭页眉和页脚' }));

  expect(parts).toEqual(['footer', 'header']);
  expect(pageNumberToggles).toBe(1);
  expect(closes).toBe(1);
  editor.destroy();
});
