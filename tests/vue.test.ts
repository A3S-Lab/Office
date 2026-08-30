import { expect, test } from '@rstest/core';
import { fireEvent, waitFor } from '@testing-library/dom';
import { Extension } from '@tiptap/core';
import { createApp, h, nextTick, ref } from 'vue';
import { Awareness } from 'y-protocols/awareness';
import * as Y from 'yjs';
import {
  createArtifact,
  createOfficeCollaborationPresence,
  createOfficeCollaborationSession,
  type DocumentContent,
  type DocumentReviewConflictEvent,
  initializeOfficeDocumentCollaboration,
  initializeOfficeMarkdownCollaboration,
  initializeOfficePresentationCollaboration,
  initializeOfficeSpreadsheetCollaboration,
  type MarkdownContent,
  readOfficeMarkdownCollaboration,
} from '../src/core';
import {
  DocumentEditor,
  MarkdownEditor,
  PdfViewer,
  PresentationEditor,
  SpreadsheetEditor,
} from '../src/vue';
import { presentationCollaborationFixture } from './fixtures/presentation-collaboration';
import { spreadsheetCollaborationFixture } from './fixtures/spreadsheet-collaboration';

test('exposes PDF evidence and page events through the Vue adapter', () => {
  const definition = PdfViewer as unknown as {
    emits: Record<string, (...arguments_: never[]) => boolean>;
    props: Record<string, unknown>;
  };

  expect(definition.props).toHaveProperty('evidenceOverlay');
  expect(definition.props).toHaveProperty('selectedEvidenceRegionId');
  expect(definition.props).toHaveProperty('worker');
  expect(definition.emits).toHaveProperty('evidenceRegionSelect');
  expect(definition.emits).toHaveProperty('pageChange');
});

test('exposes the Spreadsheet custom-list store through the Vue adapter', () => {
  const definition = SpreadsheetEditor as unknown as {
    props: Record<string, unknown>;
  };

  expect(definition.props).toHaveProperty('sortCustomListStore');
});

test('mounts the Vue adapter and renders the React editor', async () => {
  const target = document.createElement('div');
  document.body.append(target);
  const artifact = createArtifact('blank-document');
  const app = createApp({
    render: () =>
      h(DocumentEditor, {
        artifactId: artifact.id,
        content: artifact.content as DocumentContent,
        preview: true,
      }),
  });

  app.mount(target);
  await nextTick();
  await new Promise((resolve) => setTimeout(resolve, 0));

  expect(target.querySelector('[data-a3s-office]')).not.toBeNull();
  await waitFor(() => {
    expect(
      target.querySelector('[data-work-pdf-surface="live"]'),
    ).toHaveAttribute('data-work-pdf-artifact', artifact.id);
  });

  app.unmount();
  target.remove();
});

test('declares the PDF collaboration snapshot event on the Vue adapter', () => {
  const emits = PdfViewer.emits;
  expect(emits).toBeDefined();
  expect(
    Array.isArray(emits)
      ? emits.includes('collaborationChange')
      : Object.hasOwn(emits ?? {}, 'collaborationChange'),
  ).toBe(true);
});

test('passes a synchronized Document session through the Vue adapter', async () => {
  const target = document.createElement('div');
  document.body.append(target);
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a Document artifact.');
  }
  const sharedDocument = new Y.Doc();
  const awareness = new Awareness(sharedDocument);
  const session = createOfficeCollaborationSession({
    actor: { id: 'vue-editor', name: 'Vue editor' },
    artifactId: 'vue-shared-document',
    awareness,
    document: sharedDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(session, artifact.content);
  const presence = createOfficeCollaborationPresence(session);
  const app = createApp({
    render: () =>
      h(DocumentEditor, {
        collaboration: session,
        content: artifact.content as DocumentContent,
        presence,
      }),
  });

  app.mount(target);
  await waitFor(() => {
    expect(target.querySelector('[aria-label="文档正文"]')).not.toBeNull();
  });
  expect(target.querySelector('[data-collaboration-count="1"]')).not.toBeNull();

  app.unmount();
  presence.destroy();
  awareness.destroy();
  sharedDocument.destroy();
  target.remove();
});

test('preserves Document comment-mode review behavior through the Vue adapter', async () => {
  const target = document.createElement('div');
  document.body.append(target);
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a Document artifact.');
  }
  const sharedDocument = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    artifactId: 'vue-comment-document',
    document: sharedDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(writable, artifact.content);
  const reviewer = createOfficeCollaborationSession({
    actor: { id: 'vue-reviewer', name: 'Vue reviewer' },
    artifactId: 'vue-comment-document',
    document: sharedDocument,
    kind: 'document',
    mode: 'comment',
  });
  const app = createApp({
    render: () =>
      h(DocumentEditor, {
        collaboration: reviewer,
        content: artifact.content as DocumentContent,
      }),
  });

  app.mount(target);
  await waitFor(() => {
    expect(target.querySelector('[aria-label="文档正文"]')).not.toBeNull();
  });
  const editor = target.querySelector('[aria-label="文档正文"]');
  expect(editor).toHaveAttribute('contenteditable', 'true');
  expect(editor).toHaveAttribute('aria-readonly', 'true');
  const tabs = [...target.querySelectorAll('[role="tab"]')].map((tab) =>
    tab.textContent?.trim(),
  );
  expect(tabs).not.toContain('开始');
  expect(tabs).toContain('审阅');

  app.unmount();
  sharedDocument.destroy();
  target.remove();
});

test('preserves Document suggestion-mode behavior through the Vue adapter', async () => {
  const target = document.createElement('div');
  document.body.append(target);
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a Document artifact.');
  }
  const sharedDocument = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    artifactId: 'vue-suggestion-document',
    document: sharedDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(writable, artifact.content);
  const suggester = createOfficeCollaborationSession({
    actor: { id: 'vue-suggester', name: 'Vue suggester' },
    artifactId: 'vue-suggestion-document',
    document: sharedDocument,
    kind: 'document',
    mode: 'suggest',
  });
  const app = createApp({
    render: () =>
      h(DocumentEditor, {
        collaboration: suggester,
        content: artifact.content as DocumentContent,
      }),
  });

  app.mount(target);
  await waitFor(() => {
    expect(target.querySelector('[aria-label="文档正文"]')).not.toBeNull();
  });
  const editor = target.querySelector('[aria-label="文档正文"]');
  expect(editor).toHaveAttribute('contenteditable', 'true');
  expect(editor).toHaveAttribute('aria-readonly', 'false');
  expect(target.querySelector('[aria-label="建议模式"]')).toBeDisabled();
  expect(target.querySelector('[aria-label="添加批注"]')).toBeNull();

  app.unmount();
  sharedDocument.destroy();
  target.remove();
});

test('emits controlled document review conflicts through the Vue adapter', async () => {
  const target = document.createElement('div');
  document.body.append(target);
  const controlled = ref<DocumentContent>({
    type: 'document',
    html: '<p><span data-comment-id="comment-1" data-document-comment="true">Alpha</span></p>',
    pageSize: 'a4',
    comments: [
      {
        id: 'comment-1',
        author: 'Reviewer',
        date: '',
        text: 'Review Alpha.',
        resolved: false,
      },
    ],
  });
  const events: DocumentReviewConflictEvent[] = [];
  const app = createApp({
    render: () =>
      h(DocumentEditor, {
        artifactId: 'document-1',
        content: controlled.value,
        onReviewConflict: (event: DocumentReviewConflictEvent) =>
          events.push(event),
      }),
  });

  app.mount(target);
  await waitFor(() => {
    expect(target.querySelector('[role="textbox"]')).not.toBeNull();
  });
  controlled.value = {
    ...controlled.value,
    html: '<p><span data-comment-id="comment-1" data-document-comment="true">Omega</span></p>',
  };
  await nextTick();

  await waitFor(() => expect(events).toHaveLength(1));
  expect(events[0]?.conflicts[0]).toMatchObject({
    id: 'comment-1',
    kind: 'comment',
    reason: 'text-changed',
  });

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

test('passes a synchronized Markdown session through the Vue adapter', async () => {
  const target = document.createElement('div');
  document.body.append(target);
  const session = createOfficeCollaborationSession({
    artifactId: 'vue-shared-markdown',
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(session, {
    type: 'markdown',
    markdown: '# Shared through Vue',
  });
  const app = createApp({
    render: () =>
      h(MarkdownEditor, {
        collaboration: session,
        content: { type: 'markdown', markdown: '# Stale host value' },
      }),
  });

  app.mount(target);
  const source = await waitFor(() => {
    const element = target.querySelector<HTMLTextAreaElement>('textarea');
    expect(element).not.toBeNull();
    if (!element) throw new Error('Expected the Markdown source editor.');
    return element;
  });
  expect(source.value).toBe('# Shared through Vue');
  fireEvent.change(source, { target: { value: '# Vue collaboration edit' } });

  await waitFor(() =>
    expect(readOfficeMarkdownCollaboration(session).markdown).toBe(
      '# Vue collaboration edit',
    ),
  );
  app.unmount();
  target.remove();
});

test('passes a synchronized Presentation session through the Vue adapter', async () => {
  const target = document.createElement('div');
  document.body.append(target);
  const content = presentationCollaborationFixture();
  const session = createOfficeCollaborationSession({
    artifactId: 'vue-shared-presentation',
    kind: 'presentation',
  });
  initializeOfficePresentationCollaboration(session, content);
  const app = createApp({
    render: () =>
      h(PresentationEditor, {
        collaboration: session,
        content,
      }),
  });

  app.mount(target);
  await waitFor(() => {
    expect(target.textContent).toContain('Shared presentation');
  });

  app.unmount();
  target.remove();
});

test('passes a synchronized Spreadsheet session through the Vue adapter', async () => {
  const target = document.createElement('div');
  document.body.append(target);
  const content = spreadsheetCollaborationFixture();
  const session = createOfficeCollaborationSession({
    artifactId: 'vue-shared-spreadsheet',
    kind: 'spreadsheet',
  });
  initializeOfficeSpreadsheetCollaboration(session, content);
  const app = createApp({
    render: () =>
      h(SpreadsheetEditor, {
        collaboration: session,
        content,
        preview: true,
      }),
  });

  app.mount(target);
  await waitFor(() => {
    expect(target.textContent).toContain('Inputs');
  });

  app.unmount();
  target.remove();
});

test('passes the Markdown selection-menu factory through the Vue adapter', async () => {
  const target = document.createElement('div');
  document.body.append(target);
  const content: MarkdownContent = {
    type: 'markdown',
    markdown: 'Vue selection menu',
  };
  const app = createApp({
    render: () =>
      h(MarkdownEditor, {
        content,
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
  const source = await waitFor(() => {
    const element = target.querySelector<HTMLTextAreaElement>(
      '[aria-label="Markdown 源码"]',
    );
    expect(element).not.toBeNull();
    if (!element) throw new Error('Expected the Markdown editor to mount.');
    return element;
  });
  source.focus();
  source.setSelectionRange(0, 3);
  source.dispatchEvent(new Event('select', { bubbles: true }));
  source.dispatchEvent(
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
