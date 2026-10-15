import { expect, test } from '@rstest/core';
import { render, waitFor } from '@testing-library/react';
import { useState } from 'react';
import {
  applyAwarenessUpdate,
  Awareness,
  encodeAwarenessUpdate,
} from 'y-protocols/awareness';
import * as Y from 'yjs';
import {
  createArtifact,
  createArtifactBlob,
  createOfficeCollaborationPresence,
  createOfficeCollaborationSession,
  importOfficeFile,
  type SpreadsheetContent,
} from '../src/core';
import { SpreadsheetEditor } from '../src/react';

test('opens a round-tripped blank workbook without React 185', async () => {
  const artifact = createArtifact('blank-spreadsheet');
  const blob = await createArtifactBlob(artifact);
  const imported = await importOfficeFile(
    new File([blob], 'Excel-blank.xlsx', { type: blob.type }),
  );
  if (imported.content.type !== 'spreadsheet') {
    throw new Error('Expected a spreadsheet.');
  }
  const sheet = imported.content.sheets[0];
  const changes: SpreadsheetContent[] = [];

  function Harness() {
    const [content, setContent] = useState(
      imported.content as SpreadsheetContent,
    );
    return (
      <SpreadsheetEditor
        content={content}
        onChange={(next) => {
          changes.push(next);
          setContent(next);
        }}
        theme="light"
      />
    );
  }

  const view = render(<Harness />);
  await waitFor(() => {
    const failed = view.container.querySelector('[data-office-editor-error]');
    const grid = view.container.querySelector('.fortune-sheet-overlay');
    expect(failed ?? grid).not.toBeNull();
  });

  const failure = view.container.querySelector('[data-office-editor-error]');
  expect(
    failure?.textContent ?? '',
    JSON.stringify(sheetSummary(sheet)),
  ).not.toMatch(/185|Maximum update depth/);
  expect(failure).toBeNull();
  expect(changes.length).toBeLessThan(8);
});

test('opens a blank workbook while a remote collaborator cell is projected', async () => {
  const artifact = createArtifact('blank-spreadsheet');
  if (artifact.content.type !== 'spreadsheet') {
    throw new Error('Expected a spreadsheet.');
  }
  const content = artifact.content;
  const sheetId = content.sheets[0]?.id;
  if (!sheetId) throw new Error('Expected a sheet.');
  const local = presenceSide('local');
  const remote = presenceSide('remote');
  remote.presence.update({
    activity: 'active',
    location: {
      kind: 'spreadsheet',
      sheetId,
      ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }],
      activeCell: { row: 0, column: 0 },
    },
  });
  applyAwarenessUpdate(
    local.awareness,
    encodeAwarenessUpdate(remote.awareness, [remote.awareness.clientID]),
    'remote',
  );

  const view = render(
    <SpreadsheetEditor
      content={content}
      onChange={() => undefined}
      presence={local.presence}
      theme="light"
    />,
  );
  await waitFor(() => {
    const failed = view.container.querySelector('[data-office-editor-error]');
    const remoteCaret = view.container.querySelector(
      '.fortune-presence-username',
    );
    expect(failed ?? remoteCaret).not.toBeNull();
  });

  expect(
    view.container.querySelector('[data-office-editor-error]')?.textContent ??
      null,
  ).toBeNull();
  view.unmount();
  local.destroy();
  remote.destroy();
});

function presenceSide(name: string) {
  const document = new Y.Doc();
  const awareness = new Awareness(document);
  const session = createOfficeCollaborationSession({
    actor: { id: `${name}-actor`, name },
    artifactId: 'blank-workbook-presence',
    awareness,
    document,
    kind: 'spreadsheet',
  });
  const presence = createOfficeCollaborationPresence(session);
  return {
    awareness,
    presence,
    destroy() {
      presence.destroy();
      session.destroy();
      awareness.destroy();
      document.destroy();
    },
  };
}

function sheetSummary(sheet: SpreadsheetContent['sheets'][number] | undefined) {
  const cell = sheet?.data?.[0]?.[0];
  return {
    keys: sheet ? Object.keys(sheet) : [],
    row: sheet?.row,
    column: sheet?.column,
    dataRows: sheet?.data?.length,
    dataColumns: sheet?.data?.[0]?.length,
    cell,
    config: sheet?.config ?? null,
    status: sheet?.status,
  };
}
