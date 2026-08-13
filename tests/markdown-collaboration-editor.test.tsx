import { expect, test } from '@rstest/core';
import {
  act,
  fireEvent,
  render,
  screen,
  waitFor,
} from '@testing-library/react';
import { StrictMode } from 'react';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  createOfficeMarkdownCollaborationBinding,
  initializeOfficeMarkdownCollaboration,
  readOfficeMarkdownCollaboration,
  type MarkdownContent,
} from '../src/core';
import { MarkdownEditor } from '../src/react';

test('uses the synchronized Markdown document as the editor authority', async () => {
  const document = new Y.Doc();
  const session = createOfficeCollaborationSession({
    artifactId: 'shared-editor-1',
    document,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(session, markdown('# Shared source'));
  const changes: MarkdownContent[] = [];
  const view = render(
    <MarkdownEditor
      collaboration={session}
      content={markdown('# Stale host value')}
      onChange={(content) => changes.push(content)}
      theme="light"
    />,
  );

  const source = await findSource(view.container);
  expect(source).toHaveValue('# Shared source');
  fireEvent.change(source, { target: { value: '# Local source edit' } });

  expect(readOfficeMarkdownCollaboration(session).markdown).toBe(
    '# Local source edit',
  );
  expect(changes.at(-1)?.markdown).toBe('# Local source edit');

  view.rerender(
    <MarkdownEditor
      collaboration={session}
      content={markdown('# Another stale host value')}
      onChange={(content) => changes.push(content)}
      theme="light"
    />,
  );
  expect(source).toHaveValue('# Local source edit');
  expect(readOfficeMarkdownCollaboration(session).markdown).toBe(
    '# Local source edit',
  );
});

test('applies remote Markdown updates and keeps undo local to the editor', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'shared-editor-2',
    document: firstDocument,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(first, markdown('# Shared'));
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace' },
    artifactId: 'shared-editor-2',
    document: secondDocument,
    kind: 'markdown',
  });
  const remote = createOfficeMarkdownCollaborationBinding(second);
  const changes: MarkdownContent[] = [];

  const view = render(
    <MarkdownEditor
      collaboration={first}
      content={markdown('# Ignored')}
      onChange={(content) => changes.push(content)}
      theme="light"
    />,
  );

  const source = await findSource(view.container);
  const undo = screen.getByRole('button', { name: '\u64a4\u9500' });
  expect(undo).toBeDisabled();

  act(() => {
    remote.replace('# Shared remotely');
    exchangeUpdates(firstDocument, secondDocument);
  });
  await waitFor(() => expect(source).toHaveValue('# Shared remotely'));
  expect(undo).toBeDisabled();
  expect(changes.at(-1)?.markdown).toBe('# Shared remotely');

  fireEvent.change(source, {
    target: { value: '# Shared remotely and locally' },
  });
  await waitFor(() => expect(undo).toBeEnabled());
  act(() => exchangeUpdates(firstDocument, secondDocument));

  act(() => {
    remote.replace(`${remote.content().markdown} plus Grace`);
    exchangeUpdates(firstDocument, secondDocument);
  });
  await waitFor(() =>
    expect(source).toHaveValue('# Shared remotely and locally plus Grace'),
  );

  fireEvent.click(undo);
  await waitFor(() =>
    expect(source).toHaveValue('# Shared remotely plus Grace'),
  );
  expect(readOfficeMarkdownCollaboration(first).markdown).toBe(
    '# Shared remotely plus Grace',
  );
});

test('survives React StrictMode remounts without owning the shared document', async () => {
  const document = new Y.Doc();
  const session = createOfficeCollaborationSession({
    artifactId: 'shared-editor-strict',
    document,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(session, markdown('Strict source'));
  const view = render(
    <StrictMode>
      <MarkdownEditor
        collaboration={session}
        content={markdown('Ignored')}
        onChange={() => undefined}
        theme="light"
      />
    </StrictMode>,
  );

  const source = await findSource(view.container);
  fireEvent.change(source, { target: { value: 'Strict edit' } });
  expect(readOfficeMarkdownCollaboration(session).markdown).toBe('Strict edit');
  view.unmount();

  document.getMap('host-owned').set('alive', true);
  expect(document.getMap('host-owned').get('alive')).toBe(true);
});

test('renders non-edit collaboration modes as read-only surfaces', async () => {
  const writable = createOfficeCollaborationSession({
    artifactId: 'shared-editor-view',
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(writable, markdown('Read-only source'));
  const session = createOfficeCollaborationSession({
    artifactId: 'shared-editor-view',
    document: writable.document,
    kind: 'markdown',
    mode: 'view',
  });
  const view = render(
    <MarkdownEditor
      collaboration={session}
      content={markdown('Ignored')}
      onChange={() => undefined}
      theme="light"
    />,
  );

  await waitFor(() => {
    expect(view.container.querySelector('textarea')).toBeNull();
    expect(
      view.container.querySelector('[contenteditable="false"]'),
    ).not.toBeNull();
  });
  fireEvent.keyDown(
    view.container.querySelector<HTMLElement>('[contenteditable="false"]') ??
      missingElement('read-only Markdown surface'),
    { ctrlKey: true, key: 'z' },
  );
  expect(readOfficeMarkdownCollaboration(session).markdown).toBe(
    'Read-only source',
  );
});

test('projects remote changes immediately while the visual editor is editable', async () => {
  const document = new Y.Doc();
  const session = createOfficeCollaborationSession({
    artifactId: 'shared-editor-visual',
    document,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(session, markdown('# Before'));
  const remote = createOfficeMarkdownCollaborationBinding(session);
  const view = render(
    <MarkdownEditor
      collaboration={session}
      content={markdown('Ignored')}
      onChange={() => undefined}
      theme="light"
    />,
  );

  fireEvent.click(await screen.findByRole('tab', { name: '\u89c6\u56fe' }));
  fireEvent.click(
    view.container.querySelector<HTMLButtonElement>(
      '[aria-label="\u53ef\u89c6\u5316\u7f16\u8f91"]',
    ) ?? missingElement('visual editing button'),
  );
  const visual = await waitFor(() => {
    const element = view.container.querySelector<HTMLElement>(
      '[contenteditable="true"]',
    );
    expect(element).not.toBeNull();
    return element ?? missingElement('editable Markdown surface');
  });

  act(() => remote.replace('# Remote now'));

  expect(visual).toHaveTextContent('Remote now');
  expect(visual).not.toHaveTextContent('Before');
});

function markdown(value: string): MarkdownContent {
  return { type: 'markdown', markdown: value };
}

function missingElement(label: string): never {
  throw new Error(`Expected ${label}.`);
}

async function findSource(
  container: HTMLElement,
): Promise<HTMLTextAreaElement> {
  return waitFor(() => {
    const source = container.querySelector<HTMLTextAreaElement>('textarea');
    expect(source).not.toBeNull();
    if (!source) throw new Error('Expected the Markdown source editor.');
    return source;
  });
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
