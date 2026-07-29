import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { createArtifact, type DocumentContent } from '../src/core';
import { DocumentEditor } from '../src/react';

test('returns focus to an unfinished comment after cancelling a pane switch', async () => {
  const artifact = createArtifact('project-brief');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }

  render(
    <DocumentEditor
      content={artifact.content}
      onChange={() => undefined}
      theme="light"
    />,
  );

  const editor = await screen.findByRole('textbox', { name: '文档正文' });
  fireEvent.click(screen.getByRole('tab', { name: '审阅' }));
  selectDomText(editor, '这项工作的目标');

  const addComment = await screen.findByRole('button', {
    name: '添加批注',
  });
  await waitFor(() => expect(addComment).toBeEnabled());
  fireEvent.click(addComment);

  const draft = await screen.findByRole('textbox', { name: '批注内容' });
  fireEvent.change(draft, { target: { value: 'Keep this draft.' } });
  const viewChanges = screen.getByRole('button', { name: '查看修订' });
  viewChanges.focus();
  fireEvent.click(viewChanges);

  const dialog = screen.getByRole('dialog', {
    name: '放弃未完成的批注？',
  });
  fireEvent.click(within(dialog).getByRole('button', { name: '取消' }));

  await waitFor(() => expect(draft).toHaveFocus());
  expect(draft).toHaveValue('Keep this draft.');
});

test('returns focus to an edited citation after cancelling a ribbon switch', async () => {
  const artifact = createArtifact('project-brief');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const content: DocumentContent = {
    ...artifact.content,
    bibliography: { style: 'apa', sources: [] },
  };

  render(
    <DocumentEditor
      content={content}
      onChange={() => undefined}
      theme="light"
    />,
  );

  await screen.findByRole('textbox', { name: '文档正文' });
  fireEvent.click(screen.getByRole('tab', { name: '引用' }));
  fireEvent.click(await screen.findByRole('button', { name: '文献库' }));

  const title = await screen.findByRole('textbox', { name: '文献标题' });
  title.focus();
  fireEvent.change(title, { target: { value: 'Architecture Handbook' } });
  const reviewTab = screen.getByRole('tab', { name: '审阅' });
  reviewTab.focus();
  fireEvent.click(reviewTab);

  const dialog = screen.getByRole('dialog', {
    name: '放弃未保存的文献更改？',
  });
  fireEvent.click(within(dialog).getByRole('button', { name: '取消' }));

  await waitFor(() => expect(title).toHaveFocus());
  expect(title).toHaveValue('Architecture Handbook');
});

function selectDomText(root: HTMLElement, text: string): void {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let node = walker.nextNode();
  while (node) {
    const value = node.textContent ?? '';
    const offset = value.indexOf(text);
    if (offset >= 0) {
      const range = document.createRange();
      range.setStart(node, offset);
      range.setEnd(node, offset + text.length);
      const selection = window.getSelection();
      selection?.removeAllRanges();
      selection?.addRange(range);
      root.focus();
      document.dispatchEvent(new Event('selectionchange', { bubbles: true }));
      return;
    }
    node = walker.nextNode();
  }
  throw new Error(`Unable to select "${text}" in the editor.`);
}
