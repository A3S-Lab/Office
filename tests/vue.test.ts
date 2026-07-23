import { expect, test } from '@rstest/core';
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
  const content: MarkdownContent = {
    type: 'markdown',
    markdown: '## Vue adapter',
  };
  const app = createApp({
    render: () =>
      h(MarkdownEditor, {
        content,
        preview: true,
      }),
  });

  app.mount(target);
  await nextTick();
  await new Promise((resolve) => setTimeout(resolve, 0));

  expect(target.querySelector('[aria-label="Markdown 预览"]')).not.toBeNull();

  app.unmount();
  target.remove();
});
