import { expect, test } from '@rstest/core';
import { render, screen, waitFor } from '@testing-library/react';
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
