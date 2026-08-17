import { Editor } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { DocumentChangesPanel } from '../src/internal/features/work/editors/document-changes-panel';
import {
  collectDocumentChanges,
  type WorkDocumentChange,
  type WorkDocumentChangeKind,
} from '../src/internal/features/work/work-document-changes';
import { DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT } from '../src/internal/features/work/editors/document-navigation-window';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

test('keeps the empty review pane aligned with the tracking state', () => {
  const editor = createEditor();
  document.body.append(editor.view.dom);
  const trackingChanges: boolean[] = [];
  try {
    const view = render(
      <DocumentChangesPanel
        editor={editor}
        changes={[]}
        trackChanges={false}
        onTrackChangesChange={(enabled) => trackingChanges.push(enabled)}
        onClose={() => undefined}
      />,
    );

    expect(screen.getByRole('status')).toHaveTextContent(
      '当前没有记录新的改动。',
    );
    const enableTracking = screen.getByRole('button', { name: '开启修订' });
    enableTracking.focus();
    fireEvent.click(enableTracking);
    expect(trackingChanges).toEqual([true]);

    view.rerender(
      <DocumentChangesPanel
        editor={editor}
        changes={[]}
        trackChanges
        onTrackChangesChange={(enabled) => trackingChanges.push(enabled)}
        onClose={() => undefined}
      />,
    );

    expect(screen.getByRole('status')).toHaveTextContent('正在记录新的改动。');
    const stopTracking = screen.getByRole('button', { name: '停止记录' });
    expect(stopTracking).toHaveFocus();
    fireEvent.click(stopTracking);
    expect(trackingChanges).toEqual([true, false]);
    expect(
      screen.queryByRole('button', { name: '全部接受' }),
    ).not.toBeInTheDocument();
    expect(
      screen.queryByRole('button', { name: '全部拒绝' }),
    ).not.toBeInTheDocument();
  } finally {
    editor.view.dom.remove();
    editor.destroy();
  }
});

test('labels character-formatting revisions as formatting instead of deletion', () => {
  const editor = createEditor();
  document.body.append(editor.view.dom);
  const formattingChange: WorkDocumentChange = {
    id: 'formatting-1',
    kind: 'formatting',
    author: 'Ada Reviewer',
    date: '2026-08-17T14:30:00.000Z',
    from: 1,
    to: 6,
    text: 'Alpha',
  };
  try {
    const view = render(
      <DocumentChangesPanel
        editor={editor}
        changes={[formattingChange]}
        decisions={[
          {
            id: 'formatting:formatting-decided',
            changeId: 'formatting-decided',
            changeKind: 'formatting',
            suggestedBy: 'Ada Reviewer',
            suggestedAt: '2026-08-17T14:00:00.000Z',
            text: 'Beta',
            decision: 'accept',
            decidedBy: 'Grace Editor',
            decidedAt: '2026-08-17T14:31:00.000Z',
          },
        ]}
        trackChanges
        onTrackChangesChange={() => undefined}
        onClose={() => undefined}
      />,
    );

    expect(screen.getByText('格式')).toBeVisible();
    expect(screen.queryByText('删除')).toBeNull();
    expect(
      view.container.querySelector('.work-document-change-item.formatting'),
    ).toBeInTheDocument();
    expect(
      view.container.querySelector('[data-document-change-kind="formatting"]'),
    ).toBeInTheDocument();
  } finally {
    editor.view.dom.remove();
    editor.destroy();
  }
});

test('labels and navigates paragraph-formatting revisions as a paragraph range', () => {
  const editor = createEditor();
  document.body.append(editor.view.dom);
  const range = textRange(editor, 'Alpha');
  const paragraphChange: WorkDocumentChange = {
    id: 'paragraph-formatting-1',
    kind: 'paragraph-formatting',
    author: 'Ada Reviewer',
    date: '2026-08-18T10:05:00.000Z',
    from: range.from,
    to: range.to,
    text: 'Alpha',
  };
  try {
    const view = render(
      <DocumentChangesPanel
        editor={editor}
        changes={[paragraphChange]}
        trackChanges
        onTrackChangesChange={() => undefined}
        onClose={() => undefined}
      />,
    );

    expect(screen.getByText('段落格式')).toBeVisible();
    expect(
      view.container.querySelector(
        '.work-document-change-item.paragraph-formatting',
      ),
    ).toBeInTheDocument();
    fireEvent.click(screen.getByRole('button', { name: '定位修订 1' }));
    expect(editor.state.selection.from).toBe(range.from);
    expect(editor.state.selection.to).toBe(range.to);
  } finally {
    editor.view.dom.remove();
    editor.destroy();
  }
});

test('observes the revision viewport when changes arrive after the empty state', () => {
  const editor = createEditor();
  document.body.append(editor.view.dom);
  const originalResizeObserver = globalThis.ResizeObserver;
  const observedElements: Element[] = [];
  let disconnectCount = 0;

  class TestResizeObserver {
    observe(element: Element) {
      observedElements.push(element);
    }

    unobserve() {}

    disconnect() {
      disconnectCount += 1;
    }
  }

  globalThis.ResizeObserver =
    TestResizeObserver as unknown as typeof ResizeObserver;

  try {
    const view = render(
      <DocumentChangesPanel
        editor={editor}
        changes={[]}
        trackChanges={false}
        onTrackChangesChange={() => undefined}
        onClose={() => undefined}
      />,
    );
    expect(observedElements).toHaveLength(0);

    view.rerender(
      <DocumentChangesPanel
        editor={editor}
        changes={longDocumentChanges(120)}
        trackChanges
        onTrackChangesChange={() => undefined}
        onClose={() => undefined}
      />,
    );

    const list = view.container.querySelector('.work-document-change-list');
    expect(observedElements).toEqual([list]);
    view.unmount();
    expect(disconnectCount).toBe(1);
  } finally {
    globalThis.ResizeObserver = originalResizeObserver;
    editor.view.dom.remove();
    editor.destroy();
  }
});

test('confirms before rejecting every tracked change', async () => {
  const editor = createEditor();
  document.body.append(editor.view.dom);
  const range = textRange(editor, 'Alpha');
  editor.commands.replaceDocumentTextWithTrackedChange(
    range.from,
    range.to,
    'Omega',
  );
  const changes = collectDocumentChanges(editor.state.doc);

  try {
    render(
      <DocumentChangesPanel
        editor={editor}
        changes={changes}
        trackChanges
        onTrackChangesChange={() => undefined}
        onClose={() => undefined}
      />,
    );

    fireEvent.click(screen.getByRole('button', { name: '全部拒绝' }));
    const dialog = screen.getByRole('dialog', { name: '拒绝全部修订？' });
    expect(dialog).toHaveAccessibleDescription(
      `将撤销当前 ${changes.length} 项修订。`,
    );
    expect(within(dialog).getByRole('button', { name: '取消' })).toHaveFocus();
    expect(editor.getText()).toContain('Omega');

    fireEvent.click(within(dialog).getByRole('button', { name: '全部拒绝' }));
    await waitFor(() =>
      expect(collectDocumentChanges(editor.state.doc)).toEqual([]),
    );
    expect(editor.getText()).toContain('Alpha');
    expect(editor.getText()).not.toContain('Omega');
    await waitFor(() => expect(editor.view.dom).toHaveFocus());
  } finally {
    editor.view.dom.remove();
    editor.destroy();
  }
});

test('keeps keyboard focus in the review flow after individual decisions', async () => {
  const editor = createEditor();
  document.body.append(editor.view.dom);
  const range = textRange(editor, 'Alpha');
  editor.commands.replaceDocumentTextWithTrackedChange(
    range.from,
    range.to,
    'Omega',
  );
  const initialChanges = collectDocumentChanges(editor.state.doc);
  expect(initialChanges).toHaveLength(2);

  try {
    const view = render(
      <DocumentChangesPanel
        editor={editor}
        changes={initialChanges}
        trackChanges
        onTrackChangesChange={() => undefined}
        onClose={() => undefined}
      />,
    );

    fireEvent.click(screen.getByRole('button', { name: '拒绝修订 1' }));
    const remainingChanges = collectDocumentChanges(editor.state.doc);
    expect(remainingChanges).toHaveLength(1);
    view.rerender(
      <DocumentChangesPanel
        editor={editor}
        changes={remainingChanges}
        trackChanges
        onTrackChangesChange={() => undefined}
        onClose={() => undefined}
      />,
    );
    await waitFor(() =>
      expect(screen.getByRole('button', { name: '拒绝修订 1' })).toHaveFocus(),
    );

    fireEvent.click(screen.getByRole('button', { name: '接受修订 1' }));
    view.rerender(
      <DocumentChangesPanel
        editor={editor}
        changes={collectDocumentChanges(editor.state.doc)}
        trackChanges
        onTrackChangesChange={() => undefined}
        onClose={() => undefined}
      />,
    );
    await waitFor(() => expect(editor.view.dom).toHaveFocus());
  } finally {
    editor.view.dom.remove();
    editor.destroy();
  }
});

test('keeps long revision lists bounded and keyboard reachable', async () => {
  const editor = createEditor();
  document.body.append(editor.view.dom);
  const changes = longDocumentChanges(120);

  try {
    const view = render(
      <DocumentChangesPanel
        editor={editor}
        changes={changes}
        trackChanges
        onTrackChangesChange={() => undefined}
        onClose={() => undefined}
      />,
    );
    const list = view.container.querySelector('.work-document-change-list');
    if (!(list instanceof HTMLElement)) {
      throw new Error('Document revision list is missing.');
    }

    expect(list).toHaveAttribute('data-document-change-count', '120');
    expect(list).toHaveAttribute('data-document-change-windowed', 'true');
    expect(
      view.container.querySelectorAll('[data-document-change-item]').length,
    ).toBeLessThanOrEqual(DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT + 2);

    const firstChange = screen.getByRole('button', {
      name: '定位修订 1',
    });
    firstChange.focus();
    fireEvent.keyDown(firstChange, { key: 'End' });

    const lastChange = await screen.findByRole('button', {
      name: '定位修订 120',
    });
    await waitFor(() => expect(lastChange).toHaveFocus());
    expect(
      view.container.querySelector('[data-document-change-spacer="before"]'),
    ).toBeInTheDocument();
    expect(
      view.container.querySelectorAll('[data-document-change-item]').length,
    ).toBeLessThanOrEqual(DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT + 2);

    fireEvent.keyDown(lastChange, { key: 'Home' });
    const restoredFirstChange = await screen.findByRole('button', {
      name: '定位修订 1',
    });
    await waitFor(() => expect(restoredFirstChange).toHaveFocus());
    expect(
      view.container.querySelector('[data-document-change-spacer="after"]'),
    ).toBeInTheDocument();

    fireEvent.keyDown(restoredFirstChange, { key: 'ArrowDown' });
    const secondChange = await screen.findByRole('button', {
      name: '定位修订 2',
    });
    await waitFor(() => expect(secondChange).toHaveFocus());

    fireEvent.keyDown(secondChange, { key: 'PageDown' });
    const tenthChange = await screen.findByRole('button', {
      name: '定位修订 10',
    });
    await waitFor(() => expect(tenthChange).toHaveFocus());

    fireEvent.keyDown(tenthChange, { key: 'PageUp' });
    const restoredSecondChange = await screen.findByRole('button', {
      name: '定位修订 2',
    });
    await waitFor(() => expect(restoredSecondChange).toHaveFocus());
    fireEvent.keyDown(restoredSecondChange, { key: 'ArrowUp' });
    const finalFirstChange = await screen.findByRole('button', {
      name: '定位修订 1',
    });
    await waitFor(() => expect(finalFirstChange).toHaveFocus());
  } finally {
    editor.view.dom.remove();
    editor.destroy();
  }
});

function createEditor(): Editor {
  let sequence = 0;
  return new Editor({
    extensions: createWorkDocumentExtensions({
      isTracking: () => true,
      createChange: (kind: WorkDocumentChangeKind) => ({
        id: `${kind}-${++sequence}`,
        author: 'Reviewer',
        date: '2026-07-25T00:00:00.000Z',
      }),
    }),
    content:
      '<section data-document-section="true"><p>Alpha beta</p></section>',
  });
}

function textRange(editor: Editor, text: string): { from: number; to: number } {
  let range: { from: number; to: number } | null = null;
  editor.state.doc.descendants((node, position) => {
    if (range || !node.isText || !node.text) return;
    const offset = node.text.indexOf(text);
    if (offset < 0) return;
    range = {
      from: position + offset,
      to: position + offset + text.length,
    };
  });
  if (!range) throw new Error(`Unable to find "${text}" in the document.`);
  return range;
}

function longDocumentChanges(count: number): WorkDocumentChange[] {
  return Array.from({ length: count }, (_, index) => ({
    id: `revision-${index + 1}`,
    kind: index % 2 === 0 ? 'insertion' : 'deletion',
    author: 'A3S Test',
    date: '2026-08-02T00:00:00.000Z',
    from: index * 10 + 1,
    to: index * 10 + 9,
    text: `Deterministic revision marker ${String(index + 1).padStart(3, '0')}`,
  }));
}
