import { afterEach, expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { Editor } from '@tiptap/core';
import { createArtifact, type DocumentContent } from '../src/core';
import { documentTextStatistics } from '../src/internal/features/work/editors/document-editor-support';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { DocumentEditor } from '../src/react';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('counts WPS-style document text statistics across writing systems', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<h1>项目 A</h1><p>你好 world 2</p><p></p>',
  });

  expect(documentTextStatistics(editor)).toEqual({
    characterCountWithSpaces: 14,
    characterCountWithoutSpaces: 11,
    paragraphCount: 3,
    wordCount: 7,
  });
});

test('opens word count from the status bar and WPS shortcut', async () => {
  const artifact = createArtifact('project-brief');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }

  render(
    <DocumentEditor
      content={artifact.content as DocumentContent}
      onChange={() => undefined}
      theme="light"
    />,
  );

  const body = await screen.findByRole('textbox', { name: '文档正文' });
  const wordCount = screen.getByRole('button', { name: /字数统计：/ });
  expect(wordCount).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+G Meta+Shift+G',
  );
  expect(wordCount.closest('.work-office-status')).toHaveAccessibleName(
    '文档状态栏',
  );
  const statusToolbar = screen.getByRole('toolbar', {
    name: '文档视图与缩放',
  });
  const pageView = screen.getByRole('button', { name: '页面视图' });
  pageView.focus();
  fireEvent.keyDown(statusToolbar, { key: 'ArrowRight' });
  expect(screen.getByRole('button', { name: '网页视图' })).toHaveFocus();

  wordCount.focus();
  fireEvent.click(wordCount);
  const clickedDialog = await screen.findByRole('dialog', {
    name: '字数统计',
  });
  expect(clickedDialog).toHaveTextContent('页数');
  expect(clickedDialog).toHaveTextContent('字符数（不计空格）');
  expect(clickedDialog).toHaveTextContent('段落数');
  fireEvent.click(screen.getByRole('button', { name: '关闭' }));
  await waitFor(() => expect(wordCount).toHaveFocus());

  body.focus();
  fireEvent.keyDown(body, {
    key: 'g',
    code: 'KeyG',
    ctrlKey: true,
    shiftKey: true,
  });
  expect(
    await screen.findByRole('dialog', { name: '字数统计' }),
  ).toBeInTheDocument();
  fireEvent.keyDown(screen.getByRole('dialog'), { key: 'Escape' });
  await waitFor(() => expect(body).toHaveFocus());
});
