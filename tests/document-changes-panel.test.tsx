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
  type WorkDocumentChangeKind,
} from '../src/internal/features/work/work-document-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

test('keeps the empty review pane quiet', () => {
  const editor = createEditor();
  document.body.append(editor.view.dom);
  try {
    render(
      <DocumentChangesPanel
        editor={editor}
        changes={[]}
        onClose={() => undefined}
      />,
    );

    expect(screen.getByRole('status')).toHaveTextContent(
      '开启修订后，改动会显示在这里。',
    );
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
        onClose={() => undefined}
      />,
    );
    await waitFor(() => expect(editor.view.dom).toHaveFocus());
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
