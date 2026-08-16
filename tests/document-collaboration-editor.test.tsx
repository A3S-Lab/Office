import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { StrictMode } from 'react';
import * as Y from 'yjs';
import {
  createArtifact,
  createOfficeCollaborationSession,
  type DocumentContent,
  initializeOfficeDocumentCollaboration,
  officeDocumentCollaborationFragment,
  readOfficeDocumentCollaboration,
} from '../src/core';
import { DocumentEditor } from '../src/react';

test('projects remote Document updates into two mounted editors', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'document-editor-sync',
    document: firstDocument,
    kind: 'document',
  });
  const initial = documentFixture();
  initializeOfficeDocumentCollaboration(first, initial);
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace' },
    artifactId: 'document-editor-sync',
    document: secondDocument,
    kind: 'document',
  });
  const firstChanges: DocumentContent[] = [];
  const secondChanges: DocumentContent[] = [];

  render(
    <>
      <DocumentEditor
        collaboration={first}
        content={initial}
        onChange={(content) => firstChanges.push(content)}
        theme="light"
      />
      <DocumentEditor
        collaboration={second}
        content={initial}
        onChange={(content) => secondChanges.push(content)}
        theme="light"
      />
    </>,
  );

  const editors = await screen.findAllByRole('textbox', {
    name: '文档正文',
  });
  await waitFor(() => expect(editors).toHaveLength(2));
  const paragraph = officeDocumentCollaborationFragment(first)
    .toArray()[0]
    ?.toArray()[0];
  if (!(paragraph instanceof Y.XmlElement)) {
    throw new Error('Expected the shared paragraph node.');
  }
  const text = paragraph.toArray()[0];
  if (!(text instanceof Y.XmlText)) {
    throw new Error('Expected the shared text node.');
  }
  first.document.transact(() => text.insert(text.length, ' X'), 'test-local');
  exchangeUpdates(firstDocument, secondDocument);

  await waitFor(() => {
    expect(readOfficeDocumentCollaboration(first).html).toBe(
      readOfficeDocumentCollaboration(second).html,
    );
    expect(firstChanges.length).toBeGreaterThan(0);
    expect(secondChanges.length).toBeGreaterThan(0);
  });
});

test('mounts a Document collaboration editor under React StrictMode', async () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'document-strict-mode',
    kind: 'document',
  });
  const initial = documentFixture();
  initializeOfficeDocumentCollaboration(session, initial);

  render(
    <StrictMode>
      <DocumentEditor
        collaboration={session}
        content={initial}
        onChange={() => undefined}
        theme="light"
      />
    </StrictMode>,
  );

  expect(
    await screen.findByRole('textbox', { name: '文档正文' }),
  ).toBeInTheDocument();
});

test('keeps view-mode Document collaboration read-only', async () => {
  const document = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    artifactId: 'document-editor-view',
    document,
    kind: 'document',
  });
  const initial = documentFixture();
  initializeOfficeDocumentCollaboration(writable, initial);
  const readOnly = createOfficeCollaborationSession({
    artifactId: 'document-editor-view',
    document,
    kind: 'document',
    mode: 'view',
  });

  render(
    <DocumentEditor
      collaboration={readOnly}
      content={initial}
      onChange={() => undefined}
      theme="light"
    />,
  );

  expect(
    await screen.findByRole('document', { name: '文档正文' }),
  ).toHaveAttribute('contenteditable', 'false');
});

test('lets comment-mode reviewers create actor-attributed threads without edit chrome', async () => {
  const document = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    artifactId: 'document-editor-comment',
    document,
    kind: 'document',
  });
  const initial = documentFixture();
  initializeOfficeDocumentCollaboration(writable, initial);
  const reviewer = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'document-editor-comment',
    document,
    kind: 'document',
    mode: 'comment',
  });

  render(
    <DocumentEditor
      collaboration={reviewer}
      content={initial}
      onChange={() => undefined}
      theme="light"
    />,
  );

  const editor = await screen.findByRole('textbox', { name: '文档正文' });
  expect(editor).toHaveAttribute('contenteditable', 'true');
  expect(editor).toHaveAttribute('aria-readonly', 'true');
  expect(screen.queryByRole('tab', { name: '开始' })).not.toBeInTheDocument();
  expect(screen.getByRole('tab', { name: '审阅' })).toBeInTheDocument();
  expect(screen.getByRole('tab', { name: '视图' })).toBeInTheDocument();

  selectDomText(editor, 'Shared');
  const addComment = screen.getByRole('button', { name: '添加批注' });
  await waitFor(() => expect(addComment).toBeEnabled());
  fireEvent.click(addComment);
  const composer = await screen.findByRole('dialog', { name: '添加批注' });
  expect(composer).toHaveTextContent('Ada · Shared');
  fireEvent.change(
    within(composer).getByRole('textbox', { name: '批注内容' }),
    {
      target: { value: 'Please keep this wording.' },
    },
  );
  fireEvent.click(within(composer).getByRole('button', { name: '添加批注' }));

  await waitFor(() => {
    expect(readOfficeDocumentCollaboration(writable).comments).toEqual([
      expect.objectContaining({
        actorId: 'ada',
        author: 'Ada',
        text: 'Please keep this wording.',
      }),
    ]);
    expect(readOfficeDocumentCollaboration(writable).html).toContain(
      'data-comment-id',
    );
  });

  const reply = screen.getByRole('textbox', { name: '回复批注 1' });
  fireEvent.change(reply, { target: { value: 'Follow-up from Ada.' } });
  fireEvent.click(screen.getByRole('button', { name: '发送回复 1' }));
  fireEvent.click(screen.getByRole('button', { name: '解决批注 1' }));

  await waitFor(() => {
    expect(
      readOfficeDocumentCollaboration(writable).comments?.[0],
    ).toMatchObject({
      resolved: true,
      replies: [
        expect.objectContaining({
          actorId: 'ada',
          author: 'Ada',
          text: 'Follow-up from Ada.',
        }),
      ],
    });
  });
});

function documentFixture(): DocumentContent {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a Document fixture.');
  }
  return { ...artifact.content, html: '<p>Shared editor</p>' };
}

function cloneDocument(source: Y.Doc): Y.Doc {
  const clone = new Y.Doc();
  Y.applyUpdate(clone, Y.encodeStateAsUpdate(source));
  return clone;
}

function exchangeUpdates(first: Y.Doc, second: Y.Doc): void {
  const firstUpdate = Y.encodeStateAsUpdate(first, Y.encodeStateVector(second));
  const secondUpdate = Y.encodeStateAsUpdate(
    second,
    Y.encodeStateVector(first),
  );
  Y.applyUpdate(first, secondUpdate, 'test-network');
  Y.applyUpdate(second, firstUpdate, 'test-network');
}

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
