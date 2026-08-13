import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { StrictMode } from 'react';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  createOfficeSpreadsheetCollaborationBinding,
  initializeOfficeSpreadsheetCollaboration,
  type SpreadsheetContent,
  readOfficeSpreadsheetCollaboration,
} from '../src/core';
import { SpreadsheetEditor } from '../src/react';

test('projects remote Spreadsheet updates into a mounted editor', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'spreadsheet-editor-sync',
    document: firstDocument,
    kind: 'spreadsheet',
  });
  const initial = editorFixture();
  initializeOfficeSpreadsheetCollaboration(first, initial);
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    artifactId: 'spreadsheet-editor-sync',
    document: secondDocument,
    kind: 'spreadsheet',
  });
  const changes: SpreadsheetContent[] = [];

  render(
    <StrictMode>
      <SpreadsheetEditor
        collaboration={second}
        content={initial}
        onChange={(content) => changes.push(content)}
        theme="light"
      />
    </StrictMode>,
  );
  expect(await screen.findByText('Inputs')).toBeInTheDocument();
  const undo = screen.getByRole('button', { name: '\u64a4\u9500' });
  expect(undo).toBeDisabled();
  const binding = createOfficeSpreadsheetCollaborationBinding(first);
  const before = binding.content();
  binding.replace(before, {
    ...before,
    sheets: before.sheets.map((sheet) =>
      sheet.id === 'sheet-input' ? { ...sheet, name: 'Remote inputs' } : sheet,
    ),
  });
  exchangeUpdates(firstDocument, secondDocument);

  expect(await screen.findByText('Remote inputs')).toBeInTheDocument();
  await waitFor(() => expect(changes.length).toBeGreaterThan(0));
  expect(undo).toBeDisabled();
  binding.destroy();
});

test('keeps view-mode Spreadsheet collaboration read-only', async () => {
  const document = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    artifactId: 'spreadsheet-editor-view',
    document,
    kind: 'spreadsheet',
  });
  const initial = editorFixture();
  initializeOfficeSpreadsheetCollaboration(writable, initial);
  const readOnly = createOfficeCollaborationSession({
    artifactId: 'spreadsheet-editor-view',
    document,
    kind: 'spreadsheet',
    mode: 'view',
  });

  render(
    <SpreadsheetEditor
      collaboration={readOnly}
      content={initial}
      onChange={() => undefined}
      theme="light"
    />,
  );

  expect(await screen.findByText('Inputs')).toBeInTheDocument();
  expect(screen.queryByRole('button', { name: /add sheet/i })).toBeNull();
});

test('preserves a local read-only sheet activation across remote updates', async () => {
  const firstDocument = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    artifactId: 'spreadsheet-editor-view-state',
    document: firstDocument,
    kind: 'spreadsheet',
  });
  const initial = editorFixture();
  initializeOfficeSpreadsheetCollaboration(writable, initial);
  const secondDocument = cloneDocument(firstDocument);
  const readOnly = createOfficeCollaborationSession({
    artifactId: 'spreadsheet-editor-view-state',
    document: secondDocument,
    kind: 'spreadsheet',
    mode: 'view',
  });
  render(
    <SpreadsheetEditor
      collaboration={readOnly}
      content={initial}
      onChange={() => undefined}
      theme="light"
    />,
  );

  const resultsTab = await screen.findByRole('tab', { name: 'Results' });
  fireEvent.click(resultsTab);
  await waitFor(() =>
    expect(resultsTab).toHaveAttribute('aria-selected', 'true'),
  );

  const binding = createOfficeSpreadsheetCollaborationBinding(writable);
  const before = binding.content();
  binding.replace(before, {
    ...before,
    sheets: before.sheets.map((sheet) =>
      sheet.id === 'sheet-input' ? { ...sheet, name: 'Remote inputs' } : sheet,
    ),
  });
  exchangeUpdates(firstDocument, secondDocument);

  expect(await screen.findByText('Remote inputs')).toBeInTheDocument();
  expect(screen.getByRole('tab', { name: 'Results' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  binding.destroy();
});

test('keeps sheet activation local and routes edits through Yjs history', async () => {
  const document = new Y.Doc();
  const session = createOfficeCollaborationSession({
    artifactId: 'spreadsheet-editor-local-view',
    document,
    kind: 'spreadsheet',
  });
  const initial = editorFixture();
  initializeOfficeSpreadsheetCollaboration(session, initial);
  const { container } = render(
    <SpreadsheetEditor
      collaboration={session}
      content={initial}
      onChange={() => undefined}
      theme="light"
    />,
  );

  const resultsTab = await screen.findByRole('tab', { name: 'Results' });
  fireEvent.click(resultsTab);
  await waitFor(() =>
    expect(resultsTab).toHaveAttribute('aria-selected', 'true'),
  );
  expect(
    readOfficeSpreadsheetCollaboration(session).sheets.every(
      (sheet) => sheet.status === undefined,
    ),
  ).toBe(true);

  fireEvent.doubleClick(resultsTab);
  const renameInput = container.querySelector<HTMLInputElement>(
    '.work-spreadsheet-sheet-tab input',
  );
  expect(renameInput).not.toBeNull();
  fireEvent.change(renameInput as HTMLInputElement, {
    target: { value: 'Local results' },
  });
  fireEvent.keyDown(renameInput as HTMLInputElement, { key: 'Enter' });
  await waitFor(() =>
    expect(
      readOfficeSpreadsheetCollaboration(session).sheets.find(
        (sheet) => sheet.id === 'sheet-results',
      )?.name,
    ).toBe('Local results'),
  );

  const undo = screen.getByRole('button', { name: '\u64a4\u9500' });
  await waitFor(() => expect(undo).toBeEnabled());
  fireEvent.click(undo);
  await waitFor(() =>
    expect(
      readOfficeSpreadsheetCollaboration(session).sheets.find(
        (sheet) => sheet.id === 'sheet-results',
      )?.name,
    ).toBe('Results'),
  );
});

function editorFixture(): SpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-input',
        name: 'Inputs',
        status: 1,
        row: 2,
        column: 2,
        data: [
          [
            { v: 'Revenue', m: 'Revenue' },
            { v: 10, m: '10' },
          ],
          [null, null],
        ],
      },
      {
        id: 'sheet-results',
        name: 'Results',
        status: 0,
        row: 2,
        column: 2,
        data: [
          [
            { v: 'Total', m: 'Total' },
            { f: '=Inputs!B1', v: 10 },
          ],
        ],
      },
    ],
  };
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
