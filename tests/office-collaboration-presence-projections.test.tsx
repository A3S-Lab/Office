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
import { PresentationCollaborationPresenceLayer, PresentationFrameCaretProvider } from '../src/internal/features/work/editors/presentation-collaboration-presence';
import {
  spreadsheetCellLayoutBox,
  spreadsheetFrameCaretPaint,
  spreadsheetPresenceProjection,
  spreadsheetWrittenCellProjection,
  useSpreadsheetCollaborationPresenceProjection,
} from '../src/internal/features/work/editors/spreadsheet-collaboration-presence';
import { presentationFrameCaretPaint } from '../src/internal/features/work/editors/presentation-collaboration-presence';
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

test('highlights the cell a frame wrote and does not blink when that cell is unchanged', () => {
  const content: WorkSpreadsheetContent = {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-data',
        name: 'Data',
        row: 20,
        column: 10,
        celldata: [],
      },
    ],
  };
  const cell = { sheetId: 'sheet-data', row: 1, column: 0 };
  const written = spreadsheetWrittenCellProjection(content, cell, {
    kind: 'spreadsheet',
    sheetId: 'somewhere-else',
    row: 9,
    column: 9,
  });
  expect(written).toEqual([
    {
      sheetId: 'sheet-data',
      username: 'Agent',
      userId: 'a3s-office:frame-cell',
      color: '#6d28d9',
      selection: { r: 1, c: 0 },
    },
  ]);
  expect(JSON.stringify(spreadsheetWrittenCellProjection(content, cell))).toBe(
    JSON.stringify(written),
  );

  const adds: number[] = [];
  const removes: number[] = [];
  const workbook = {
    addPresences: () => adds.push(1),
    removePresences: () => removes.push(1),
  };
  function Probe({ tick }: { tick: number }) {
    useSpreadsheetCollaborationPresenceProjection({
      content,
      frameCell: cell,
      refreshKey: 0,
      workbook: workbook as never,
    });
    return <span>{tick}</span>;
  }
  const view = render(<Probe tick={0} />);
  view.rerender(<Probe tick={1} />);
  expect(adds).toEqual([1]);
  expect(removes).toEqual([]);
  view.unmount();
});

test('paints a spreadsheet splice caret inside the plain-text cell and skips field writes', () => {
  const measure = (slice: string) => slice.length * 8;
  const box = { left: 44, top: 24, width: 96, height: 24 };
  const content: WorkSpreadsheetContent = {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-data',
        name: 'Data',
        row: 20,
        column: 10,
        celldata: [
          { r: 1, c: 0, v: { v: '中文', m: '中文' } },
          { r: 2, c: 0, v: { v: 12, m: '12' } },
          { r: 3, c: 0, v: { f: '=A1', v: '中文', m: '中文' } },
        ],
      },
    ],
  };
  expect(
    spreadsheetFrameCaretPaint({
      box,
      cell: { sheetId: 'sheet-data', row: 1, column: 0, indexUtf16: 1 },
      content,
      layoutSettled: true,
      measure,
    }),
  ).toMatchObject({ head: 1, left: 52, width: 2 });
  expect(
    spreadsheetFrameCaretPaint({
      box,
      cell: { sheetId: 'sheet-data', row: 1, column: 0 },
      content,
      layoutSettled: true,
      measure,
    }),
  ).toBeNull();
  expect(
    spreadsheetFrameCaretPaint({
      box,
      cell: { sheetId: 'sheet-data', row: 2, column: 0, indexUtf16: 1 },
      content,
      layoutSettled: true,
      measure,
    }),
  ).toBeNull();
  expect(
    spreadsheetFrameCaretPaint({
      box,
      cell: { sheetId: 'sheet-data', row: 3, column: 0, indexUtf16: 1 },
      content,
      layoutSettled: true,
      measure,
    }),
  ).toBeNull();
});

test('paints a presentation splice caret inside the shape text', () => {
  const element: WorkSlideElement = {
    id: 'shape-1',
    type: 'shape',
    x: 12,
    y: 18,
    width: 36,
    height: 24,
    text: '中文',
    fontSize: 16,
    color: '#111827',
    fill: '#ffffff',
    bold: false,
    align: 'left',
  };
  const box = { left: 0, top: 0, width: 120, height: 40 };
  const measure = (slice: string) => slice.length * 16;
  expect(
    presentationFrameCaretPaint({
      box,
      caret: {
        kind: 'presentation',
        containerKind: 'slide',
        containerId: 'slide-1',
        elementId: 'shape-1',
        indexUtf16: 1,
      },
      element,
      layoutSettled: true,
      measure,
    }),
  ).toMatchObject({ head: 1, left: 16, width: 2 });
  expect(
    presentationFrameCaretPaint({
      box,
      caret: {
        kind: 'presentation',
        containerKind: 'slide',
        containerId: 'slide-1',
        elementId: 'shape-1',
      },
      element,
      layoutSettled: true,
      measure,
    }),
  ).toBeNull();
  expect(
    presentationFrameCaretPaint({
      box,
      caret: {
        kind: 'presentation',
        containerKind: 'slide',
        containerId: 'slide-1',
        elementId: 'other',
        indexUtf16: 1,
      },
      element,
      layoutSettled: true,
      measure,
    }),
  ).toBeNull();

  const view = render(
    <PresentationFrameCaretProvider
      caret={{
        kind: 'presentation',
        containerKind: 'slide',
        containerId: 'slide-1',
        elementId: 'shape-1',
        indexUtf16: 1,
      }}
    >
      <PresentationCollaborationPresenceLayer
        elements={[element]}
        measure={measure}
        slideId="slide-1"
        textBox={box}
      />
    </PresentationFrameCaretProvider>,
  );
  const caret = document.querySelector<HTMLElement>(
    '[data-frame-caret-index="1"]',
  );
  expect(caret).not.toBeNull();
  expect(caret).toHaveStyle({ left: '16px', width: '2px' });
  view.unmount();
});

test('offsets the spreadsheet cell box by headers and scroll', () => {
  const box = spreadsheetCellLayoutBox({
    column: 1,
    columnHeaderHeight: 24,
    row: 1,
    rowHeaderWidth: 44,
    scrollLeft: 10,
    scrollTop: 4,
    sheet: {
      id: 'sheet-data',
      name: 'Data',
      defaultColWidth: 96,
      defaultRowHeight: 24,
      config: { columnlen: { 0: 80 } },
    },
  });
  expect(box).toEqual({ left: 114, top: 44, width: 96, height: 24 });
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
