import { expect, test } from '@rstest/core';
import { render, screen, waitFor } from '@testing-library/react';
import { StrictMode } from 'react';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  createOfficePresentationCollaborationBinding,
  initializeOfficePresentationCollaboration,
  type PresentationContent,
} from '../src/core';
import { PresentationEditor } from '../src/react';
import { presentationCollaborationFixture as presentationFixture } from './fixtures/presentation-collaboration';

test('projects remote Presentation updates into a mounted editor', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'presentation-editor-sync',
    document: firstDocument,
    kind: 'presentation',
  });
  const initial = presentationFixture();
  initializeOfficePresentationCollaboration(first, initial);
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    artifactId: 'presentation-editor-sync',
    document: secondDocument,
    kind: 'presentation',
  });
  const changes: PresentationContent[] = [];

  render(
    <StrictMode>
      <PresentationEditor
        collaboration={second}
        content={initial}
        onChange={(content) => changes.push(content)}
        theme="light"
      />
    </StrictMode>,
  );
  expect(await screen.findAllByText('Shared presentation')).not.toHaveLength(0);
  const binding = createOfficePresentationCollaborationBinding(first);
  const before = binding.content();
  binding.replace(before, {
    ...before,
    slides: before.slides.map((slide) =>
      slide.id === 'slide-1'
        ? {
            ...slide,
            elements: slide.elements.map((element) =>
              element.id === 'element-title'
                ? { ...element, text: 'Remote presentation title' }
                : element,
            ),
          }
        : slide,
    ),
  });
  exchangeUpdates(firstDocument, secondDocument);

  expect(
    await screen.findAllByText('Remote presentation title'),
  ).not.toHaveLength(0);
  await waitFor(() => expect(changes.length).toBeGreaterThan(0));
  binding.destroy();
});

test('keeps view-mode Presentation collaboration read-only', async () => {
  const document = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    artifactId: 'presentation-editor-view',
    document,
    kind: 'presentation',
  });
  const initial = presentationFixture();
  initializeOfficePresentationCollaboration(writable, initial);
  const readOnly = createOfficeCollaborationSession({
    artifactId: 'presentation-editor-view',
    document,
    kind: 'presentation',
    mode: 'view',
  });

  render(
    <PresentationEditor
      collaboration={readOnly}
      content={initial}
      onChange={() => undefined}
      theme="light"
    />,
  );

  expect(await screen.findAllByText('Shared presentation')).not.toHaveLength(0);
  expect(screen.queryByRole('button', { name: /add slide/i })).toBeNull();
});

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
