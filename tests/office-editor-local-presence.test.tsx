import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { Awareness } from 'y-protocols/awareness';
import * as Y from 'yjs';
import {
  createArtifact,
  createOfficeCollaborationPresence,
  createOfficeCollaborationSession,
  initializeOfficeDocumentCollaboration,
  initializeOfficeMarkdownCollaboration,
  initializeOfficePresentationCollaboration,
  initializeOfficeSpreadsheetCollaboration,
  type SpreadsheetContent,
} from '../src/core';
import {
  DocumentEditor,
  MarkdownEditor,
  PresentationEditor,
  SpreadsheetEditor,
} from '../src/react';
import { OfficeCollaborationPresenceProvider } from '../src/internal/features/work/editors/office-collaboration-presence-context';
import { useOfficePublishPresenceLocation } from '../src/internal/features/work/editors/office-collaboration-presence-ui';
import { presentationCollaborationFixture } from './fixtures/presentation-collaboration';

test('publishes and clears the mounted Document editor selection', async () => {
  const fixture = editorPresenceFixture('document');
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') throw new Error('Document fixture');
  initializeOfficeDocumentCollaboration(fixture.session, artifact.content);
  const view = render(
    <DocumentEditor
      collaboration={fixture.session}
      content={artifact.content}
      presence={fixture.presence}
      onChange={() => undefined}
      theme="light"
    />,
  );

  await screen.findByRole('textbox', { name: '文档正文' });
  await waitFor(() =>
    expect(fixture.presence.local().location).toMatchObject({
      kind: 'document',
      anchor: expect.any(Number),
      head: expect.any(Number),
    }),
  );
  view.unmount();
  expect(fixture.presence.local().location).toBeUndefined();
  fixture.destroy();
});

test('publishes directional Markdown source selections', async () => {
  const fixture = editorPresenceFixture('markdown');
  const content = { type: 'markdown' as const, markdown: '# Shared source' };
  initializeOfficeMarkdownCollaboration(fixture.session, content);
  const view = render(
    <MarkdownEditor
      collaboration={fixture.session}
      content={content}
      presence={fixture.presence}
      onChange={() => undefined}
      theme="light"
    />,
  );

  const source = await screen.findByRole('textbox', {
    name: 'Markdown 源码',
  });
  source.setSelectionRange(2, 8, 'backward');
  fireEvent.select(source);
  await waitFor(() =>
    expect(fixture.presence.local().location).toEqual({
      kind: 'markdown',
      surface: 'source',
      anchor: 8,
      head: 2,
    }),
  );
  view.unmount();
  fixture.destroy();
});

test('publishes the active Spreadsheet sheet and finite range', async () => {
  const fixture = editorPresenceFixture('spreadsheet');
  const content: SpreadsheetContent = {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Inputs',
        status: 1,
        row: 2,
        column: 2,
        data: [
          [
            { v: 'Input', m: 'Input' },
            { v: 7, m: '7' },
          ],
        ],
      },
    ],
  };
  initializeOfficeSpreadsheetCollaboration(fixture.session, content);
  const view = render(
    <SpreadsheetEditor
      collaboration={fixture.session}
      content={content}
      presence={fixture.presence}
      onChange={() => undefined}
      theme="light"
    />,
  );

  await screen.findByText('Inputs');
  await waitFor(() =>
    expect(fixture.presence.local().location).toMatchObject({
      kind: 'spreadsheet',
      sheetId: 'sheet-1',
      ranges: [
        {
          startRow: expect.any(Number),
          startColumn: expect.any(Number),
          endRow: expect.any(Number),
          endColumn: expect.any(Number),
        },
      ],
    }),
  );
  view.unmount();
  fixture.destroy();
});

test('publishes Presentation slide and object selections', async () => {
  const fixture = editorPresenceFixture('presentation');
  const content = presentationCollaborationFixture();
  initializeOfficePresentationCollaboration(fixture.session, content);
  const view = render(
    <PresentationEditor
      collaboration={fixture.session}
      content={content}
      presence={fixture.presence}
      onChange={() => undefined}
      theme="light"
    />,
  );

  await screen.findAllByText('Shared presentation');
  const element = view.container.querySelector<HTMLElement>(
    '[data-slide-element-id="element-title"]',
  );
  expect(element).not.toBeNull();
  fireEvent.focus(element as HTMLElement);
  await waitFor(() =>
    expect(fixture.presence.local().location).toEqual({
      kind: 'presentation',
      slideId: 'slide-1',
      elementIds: ['element-title'],
    }),
  );
  view.unmount();
  fixture.destroy();
});

test('publishes and clears PDF page and annotation locations', async () => {
  const document = new Y.Doc();
  const awareness = new Awareness(document);
  const session = createOfficeCollaborationSession({
    actor: { id: 'local-pdf', name: 'Local PDF editor' },
    artifactId: 'local-presence-pdf',
    awareness,
    document,
    kind: 'pdf',
  });
  const presence = createOfficeCollaborationPresence(session);
  const view = render(
    <OfficeCollaborationPresenceProvider presence={presence}>
      <PdfPresencePublisher pageIndex={3} annotationId="note-2" />
    </OfficeCollaborationPresenceProvider>,
  );

  await waitFor(() =>
    expect(presence.local().location).toEqual({
      kind: 'pdf',
      pageIndex: 3,
      annotationId: 'note-2',
    }),
  );
  view.rerender(
    <OfficeCollaborationPresenceProvider presence={presence}>
      <PdfPresencePublisher pageIndex={null} />
    </OfficeCollaborationPresenceProvider>,
  );
  await waitFor(() => expect(presence.local().location).toBeUndefined());
  view.unmount();
  presence.destroy();
  session.destroy();
  awareness.destroy();
  document.destroy();
});

function editorPresenceFixture(
  kind: 'document' | 'markdown' | 'presentation' | 'spreadsheet',
) {
  const document = new Y.Doc();
  const awareness = new Awareness(document);
  const session = createOfficeCollaborationSession({
    actor: { id: `local-${kind}`, name: 'Local editor' },
    artifactId: `local-presence-${kind}`,
    awareness,
    document,
    kind,
  });
  const presence = createOfficeCollaborationPresence(session);
  return {
    presence,
    session,
    destroy() {
      presence.destroy();
      session.destroy();
      awareness.destroy();
      document.destroy();
    },
  };
}

function PdfPresencePublisher({
  annotationId,
  pageIndex,
}: {
  annotationId?: string;
  pageIndex: number | null;
}) {
  useOfficePublishPresenceLocation(
    pageIndex === null
      ? null
      : {
          kind: 'pdf',
          pageIndex,
          ...(annotationId ? { annotationId } : {}),
        },
  );
  return null;
}
