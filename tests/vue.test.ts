import { Extension } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import { waitFor } from '@testing-library/dom';
import { createApp, h, nextTick } from 'vue';
import {
  createArtifact,
  type DocumentContent,
  type MarkdownContent,
} from '../src/core';
import { DocumentEditor, MarkdownEditor } from '../src/vue';

test('mounts the Vue adapter and renders the React editor', async () => {
  const target = document.createElement('div');
  document.body.append(target);
  const artifact = createArtifact('blank-document');
  const app = createApp({
    render: () =>
      h(DocumentEditor, {
        content: artifact.content as DocumentContent,
        preview: true,
      }),
  });

  app.mount(target);
  await nextTick();
  await new Promise((resolve) => setTimeout(resolve, 0));

  expect(target.querySelector('[data-a3s-office]')).not.toBeNull();

  app.unmount();
  target.remove();
});

test('mounts the Vue Markdown adapter', async () => {
  const target = document.createElement('div');
  document.body.append(target);
  const hostShortcuts = Extension.create({
    name: 'testVueHostShortcuts',
  });
  const content: MarkdownContent = {
    type: 'markdown',
    markdown: '## Vue adapter',
  };
  const app = createApp({
    render: () =>
      h(MarkdownEditor, {
        content,
        extensions: [hostShortcuts],
        preview: true,
      }),
  });

  app.mount(target);
  await nextTick();
  await new Promise((resolve) => setTimeout(resolve, 0));

  await waitFor(() => {
    expect(target.querySelector('[aria-label="Markdown 预览"]')).not.toBeNull();
  });

  app.unmount();
  target.remove();
});

test('passes the document selection-menu factory through the Vue adapter', async () => {
  const target = document.createElement('div');
  document.body.append(target);
  const artifact = createArtifact('project-brief');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const app = createApp({
    render: () =>
      h(DocumentEditor, {
        content: artifact.content,
        getSelectionMenuItems: () => [
          {
            id: 'host-action',
            label: '宿主操作',
            onSelect: () => undefined,
          },
        ],
      }),
  });

  app.mount(target);
  await nextTick();
  const editor = await waitFor(() => {
    const element = target.querySelector<HTMLElement>(
      '[aria-label="文档正文"]',
    );
    expect(element).not.toBeNull();
    if (!element) throw new Error('Expected the document editor to mount.');
    return element;
  });
  selectDomText(editor, '这项工作的目标');
  await new Promise((resolve) => setTimeout(resolve, 0));
  editor.dispatchEvent(
    new MouseEvent('contextmenu', {
      bubbles: true,
      cancelable: true,
      clientX: 120,
      clientY: 180,
    }),
  );

  await waitFor(() => {
    expect(
      target.querySelector('[role="menu"][aria-label="选中文本操作"]'),
    ).not.toBeNull();
  });
  expect(target.textContent).toContain('宿主操作');

  app.unmount();
  target.remove();
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
