import { expect, test } from '@rstest/core';
import { act, render, screen, waitFor } from '@testing-library/react';
import { useRef } from 'react';
import {
  Awareness,
  applyAwarenessUpdate,
  encodeAwarenessUpdate,
} from 'y-protocols/awareness';
import * as Y from 'yjs';
import {
  createOfficeCollaborationPresence,
  createOfficeCollaborationSession,
  type OfficeArtifactKind,
  type OfficeCollaborationPresenceLocation,
} from '../src/core';
import { OfficeCollaborationPresenceProvider } from '../src/internal/features/work/editors/office-collaboration-presence-context';
import { MarkdownSourcePresenceLayer } from '../src/internal/features/work/editors/office-collaboration-presence-ui';
import { PdfCollaborationPresenceLayer } from '../src/internal/features/work/editors/pdf-collaboration-presence';
import { PresentationCollaborationPresenceLayer } from '../src/internal/features/work/editors/presentation-collaboration-presence';
import { spreadsheetPresenceProjection } from '../src/internal/features/work/editors/spreadsheet-collaboration-presence';
import type {
  WorkSlideElement,
  WorkSpreadsheetContent,
} from '../src/internal/features/work/work-types';

test('projects spreadsheet participants through the native workbook Presence model', () => {
  const fixture = presenceFixture('spreadsheet', {
    kind: 'spreadsheet',
    sheetId: 'sheet-1',
    ranges: [
      {
        startRow: 2,
        startColumn: 3,
        endRow: 4,
        endColumn: 5,
      },
    ],
    activeCell: { row: 3, column: 4 },
  });
  const content: WorkSpreadsheetContent = {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        row: 20,
        column: 10,
        celldata: [],
      },
    ],
  };

  try {
    const projected = spreadsheetPresenceProjection(
      content,
      fixture.localPresence
        .snapshot()
        .participants.filter((participant) => !participant.local),
    );
    expect(projected).toEqual([
      {
        sheetId: 'sheet-1',
        username: 'Remote agent',
        userId: expect.stringMatching(/^a3s-office:/),
        color: '#6d28d9',
        selection: { r: 3, c: 4 },
      },
    ]);

    fixture.remotePresence.update({
      location: {
        kind: 'spreadsheet',
        sheetId: 'sheet-1',
        ranges: [
          {
            startRow: 99,
            startColumn: 0,
            endRow: 99,
            endColumn: 0,
          },
        ],
      },
    });
    fixture.relay();
    expect(
      spreadsheetPresenceProjection(
        content,
        fixture.localPresence
          .snapshot()
          .participants.filter((participant) => !participant.local),
      ),
    ).toEqual([]);
  } finally {
    fixture.destroy();
  }
});

test('projects presentation object geometry without disturbing local focus', async () => {
  const fixture = presenceFixture('presentation', {
    kind: 'presentation',
    slideId: 'slide-1',
    elementIds: ['shape-1'],
  });
  const element: WorkSlideElement = {
    id: 'shape-1',
    type: 'shape',
    x: 12,
    y: 18,
    width: 36,
    height: 24,
    text: 'Roadmap',
    fontSize: 24,
    color: '#111827',
    fill: '#ffffff',
    bold: false,
    align: 'left',
  };
  const view = render(
    <OfficeCollaborationPresenceProvider presence={fixture.localPresence}>
      <button type="button">Local focus</button>
      <PresentationCollaborationPresenceLayer
        elements={[element]}
        slideId="slide-1"
      />
    </OfficeCollaborationPresenceProvider>,
  );

  try {
    const focus = screen.getByRole('button', { name: 'Local focus' });
    focus.focus();
    const frame = document.querySelector<HTMLElement>(
      '[data-remote-slide-element-id="shape-1"]',
    );
    expect(frame).not.toBeNull();
    expect(frame).toHaveStyle({
      left: '12%',
      top: '18%',
      width: '36%',
      height: '24%',
    });
    expect(screen.getByText('Remote agent')).toBeVisible();

    act(() => {
      fixture.remotePresence.update({ activity: 'idle' });
      fixture.relay();
    });
    await waitFor(() => expect(frame).toHaveAttribute('data-activity', 'idle'));
    expect(focus).toHaveFocus();
  } finally {
    view.unmount();
    fixture.destroy();
  }
});

test('shows PDF participants only on their current page and preserves annotation identity', () => {
  const fixture = presenceFixture('pdf', {
    kind: 'pdf',
    pageIndex: 2,
    annotationId: 'note-7',
  });
  const view = render(
    <OfficeCollaborationPresenceProvider presence={fixture.localPresence}>
      <PdfCollaborationPresenceLayer pageIndex={2} />
    </OfficeCollaborationPresenceProvider>,
  );

  try {
    const participant = document.querySelector<HTMLElement>(
      '[data-participant-id="remote-agent"]',
    );
    expect(participant).toHaveAttribute('data-annotation-id', 'note-7');
    expect(participant).toHaveTextContent('Remote agent');
    expect(participant).toHaveTextContent('正在查看批注');

    view.rerender(
      <OfficeCollaborationPresenceProvider presence={fixture.localPresence}>
        <PdfCollaborationPresenceLayer pageIndex={1} />
      </OfficeCollaborationPresenceProvider>,
    );
    expect(
      document.querySelector('[data-participant-id="remote-agent"]'),
    ).toBeNull();
  } finally {
    view.unmount();
    fixture.destroy();
  }
});

test('ignores stale Markdown source offsets until the remote location is valid', async () => {
  const fixture = presenceFixture('markdown', {
    kind: 'markdown',
    surface: 'source',
    anchor: 50,
    head: 60,
  });
  const markdown = 'Short source';
  const view = render(
    <OfficeCollaborationPresenceProvider presence={fixture.localPresence}>
      <MarkdownSourcePresenceProbe markdown={markdown} />
    </OfficeCollaborationPresenceProvider>,
  );

  try {
    expect(
      document.querySelector('[data-participant-id="remote-agent"]'),
    ).toBeNull();

    act(() => {
      fixture.remotePresence.update({
        location: {
          kind: 'markdown',
          surface: 'source',
          anchor: 0,
          head: 5,
        },
      });
      fixture.relay();
    });
    await waitFor(() =>
      expect(
        document.querySelector('[data-participant-id="remote-agent"]'),
      ).not.toBeNull(),
    );
  } finally {
    view.unmount();
    fixture.destroy();
  }
});

function MarkdownSourcePresenceProbe({ markdown }: { markdown: string }) {
  const sourceRef = useRef<HTMLTextAreaElement>(null);
  return (
    <section>
      <textarea ref={sourceRef} defaultValue={markdown} />
      <MarkdownSourcePresenceLayer markdown={markdown} sourceRef={sourceRef} />
    </section>
  );
}

function presenceFixture(
  kind: OfficeArtifactKind,
  location: OfficeCollaborationPresenceLocation,
) {
  const localDocument = new Y.Doc();
  const remoteDocument = new Y.Doc();
  const localAwareness = new Awareness(localDocument);
  const remoteAwareness = new Awareness(remoteDocument);
  const artifactId = `projection-${kind}`;
  const localSession = createOfficeCollaborationSession({
    actor: { id: 'local-user', name: 'Local user' },
    artifactId,
    awareness: localAwareness,
    document: localDocument,
    kind,
  });
  const remoteSession = createOfficeCollaborationSession({
    actor: {
      id: 'remote-agent',
      name: 'Remote agent',
      color: '#6d28d9',
      kind: 'agent',
    },
    artifactId,
    awareness: remoteAwareness,
    document: remoteDocument,
    kind,
  });
  const localPresence = createOfficeCollaborationPresence(localSession);
  const remotePresence = createOfficeCollaborationPresence(remoteSession, {
    location,
  });
  const relay = () =>
    applyAwarenessUpdate(
      localAwareness,
      encodeAwarenessUpdate(remoteAwareness, [remoteAwareness.clientID]),
      'test-transport',
    );
  relay();
  return {
    localPresence,
    remotePresence,
    relay,
    destroy() {
      localPresence.destroy();
      remotePresence.destroy();
      localSession.destroy();
      remoteSession.destroy();
      localAwareness.destroy();
      remoteAwareness.destroy();
      localDocument.destroy();
      remoteDocument.destroy();
    },
  };
}
