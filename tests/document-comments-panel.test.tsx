import { Editor } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { useRef, useState } from 'react';
import { DocumentCommentsPanel } from '../src/internal/features/work/editors/document-comments-panel';
import type { DocumentCommentDraft } from '../src/internal/features/work/editors/document-comment-composer';
import type { WorkDocumentCommentView } from '../src/internal/features/work/work-document-comments';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

const comment: WorkDocumentCommentView = {
  id: 'comment-1',
  author: 'Reviewer',
  date: '2026-07-25T00:00:00.000Z',
  text: 'Clarify this sentence.',
  resolved: false,
  from: 1,
  to: 6,
  anchorText: 'Alpha',
  detached: false,
  replies: [
    {
      id: 'reply-1',
      author: 'Author',
      date: '2026-07-25T01:00:00.000Z',
      text: 'Will do.',
    },
  ],
};

test('confirms before deleting a comment and its pending reply', async () => {
  const editor = new Editor({
    extensions: [StarterKit],
    content: '<p>Alpha</p>',
  });
  const deleted: string[] = [];

  try {
    const view = render(
      <CommentsPanelHarness
        editor={editor}
        onDelete={(id) => deleted.push(id)}
      />,
    );

    const reply = screen.getByRole('textbox', { name: '回复批注 1' });
    fireEvent.change(reply, {
      target: { value: 'Unsaved reply' },
    });
    const deleteButton = screen.getByRole('button', { name: '删除批注 1' });
    deleteButton.focus();
    fireEvent.click(deleteButton);

    const dialog = screen.getByRole('dialog', { name: '删除批注？' });
    expect(dialog).toHaveAccessibleDescription(
      '批注、已有回复和未发送的回复都将删除。',
    );
    expect(deleted).toEqual([]);

    fireEvent.click(within(dialog).getByRole('button', { name: '取消' }));
    await waitFor(() => expect(reply).toHaveFocus());
    expect(reply).toHaveValue('Unsaved reply');
    expect(deleted).toEqual([]);

    fireEvent.click(deleteButton);
    fireEvent.click(
      within(screen.getByRole('dialog', { name: '删除批注？' })).getByRole(
        'button',
        { name: '删除' },
      ),
    );
    await waitFor(() => expect(deleted).toEqual(['comment-1']));
    view.unmount();
  } finally {
    editor.destroy();
  }
});

test('closes an empty comment draft directly and confirms written content', async () => {
  const editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Alpha</p>',
  });
  const cancelled: string[] = [];
  const draft: DocumentCommentDraft = {
    id: 'comment-draft-1',
    from: 1,
    to: 6,
    anchorText: 'Alpha',
  };
  editor.commands.showDocumentCommentDraft({ from: 1, to: 6 });

  try {
    const emptyView = render(
      <CommentsPanelHarness
        editor={editor}
        comments={[]}
        draft={draft}
        onCancelDraft={() => cancelled.push('empty')}
        onDelete={() => undefined}
      />,
    );
    fireEvent.click(screen.getByRole('button', { name: '取消' }));
    expect(cancelled).toEqual(['empty']);
    expect(
      screen.queryByRole('dialog', { name: '放弃未完成的批注？' }),
    ).not.toBeInTheDocument();
    emptyView.unmount();

    const dirtyView = render(
      <CommentsPanelHarness
        editor={editor}
        comments={[]}
        draft={draft}
        onCancelDraft={() => cancelled.push('dirty')}
        onDelete={() => undefined}
      />,
    );
    const input = screen.getByRole('textbox', { name: '批注内容' });
    fireEvent.change(input, { target: { value: 'Clarify this.' } });
    fireEvent.click(screen.getByRole('button', { name: '取消' }));

    let dialog = screen.getByRole('dialog', {
      name: '放弃未完成的批注？',
    });
    expect(cancelled).toEqual(['empty']);
    fireEvent.click(within(dialog).getByRole('button', { name: '取消' }));
    await waitFor(() => expect(input).toHaveFocus());
    expect(input).toHaveValue('Clarify this.');

    fireEvent.click(screen.getByRole('button', { name: '取消' }));
    dialog = screen.getByRole('dialog', { name: '放弃未完成的批注？' });
    fireEvent.click(within(dialog).getByRole('button', { name: '放弃内容' }));
    await waitFor(() => expect(cancelled).toEqual(['empty', 'dirty']));
    dirtyView.unmount();
  } finally {
    editor.destroy();
  }
});

test('keeps long comment review bounded and keyboard reachable', async () => {
  const editor = new Editor({
    extensions: [StarterKit],
    content: '<p>Alpha</p>',
  });

  try {
    const view = render(
      <CommentsPanelHarness
        editor={editor}
        comments={longComments(120)}
        onDelete={() => undefined}
      />,
    );
    const track = view.container.querySelector('.work-document-comment-track');
    if (!(track instanceof HTMLElement)) {
      throw new Error('Document comment track is missing.');
    }

    expect(track.tagName).toBe('OL');
    expect(track).toHaveAttribute('aria-label', '文档批注');
    expect(track).toHaveAttribute('data-document-comment-count', '120');
    expect(track).toHaveAttribute('data-document-comment-windowed', 'true');
    expect(
      view.container.querySelectorAll('[data-document-comment-item]').length,
    ).toBeLessThanOrEqual(34);

    const firstComment = screen.getByRole('button', { name: '定位批注 1' });
    firstComment.focus();
    fireEvent.keyDown(firstComment, { key: 'End' });

    const lastComment = await screen.findByRole('button', {
      name: '定位批注 120',
    });
    await waitFor(() => expect(lastComment).toHaveFocus());
    expect(editor.state.selection.from).toBe(editor.state.doc.content.size - 1);
    expect(editor.state.selection.to).toBe(editor.state.doc.content.size - 1);
    expect(track).toHaveAttribute('data-document-comment-window-start', '88');
    expect(
      view.container.querySelectorAll('[data-document-comment-item]').length,
    ).toBeLessThanOrEqual(34);

    fireEvent.keyDown(lastComment, { key: 'Home' });
    const restoredFirstComment = await screen.findByRole('button', {
      name: '定位批注 1',
    });
    await waitFor(() => expect(restoredFirstComment).toHaveFocus());
    expect(editor.state.selection.from).toBe(1);
    expect(editor.state.selection.to).toBe(2);
    expect(track).toHaveAttribute('data-document-comment-window-start', '0');

    fireEvent.keyDown(restoredFirstComment, { key: 'PageDown' });
    const ninthComment = await screen.findByRole('button', {
      name: '定位批注 9',
    });
    await waitFor(() => expect(ninthComment).toHaveFocus());
  } finally {
    editor.destroy();
  }
});

test('moves focus to the adjacent mounted comment after deletion', async () => {
  const editor = new Editor({
    extensions: [StarterKit],
    content: '<p>Alpha</p>',
  });

  try {
    render(<StatefulLongCommentsHarness editor={editor} />);
    const firstComment = screen.getByRole('button', { name: '定位批注 1' });
    firstComment.focus();
    fireEvent.keyDown(firstComment, { key: 'End' });
    const lastComment = await screen.findByRole('button', {
      name: '定位批注 120',
    });
    await waitFor(() => expect(lastComment).toHaveFocus());

    fireEvent.click(screen.getByRole('button', { name: '删除批注 120' }));
    fireEvent.click(
      within(screen.getByRole('dialog', { name: '删除批注？' })).getByRole(
        'button',
        { name: '删除' },
      ),
    );

    const adjacentComment = await screen.findByRole('button', {
      name: '定位批注 119',
    });
    await waitFor(() => expect(adjacentComment).toHaveFocus());
    expect(
      screen.queryByRole('button', { name: '定位批注 120' }),
    ).not.toBeInTheDocument();
  } finally {
    editor.destroy();
  }
});

test('keeps a detached thread visible without moving the editor selection', () => {
  const editor = new Editor({
    extensions: [StarterKit],
    content: '<p>Alpha</p>',
  });
  editor.commands.setTextSelection(3);
  const selection = {
    from: editor.state.selection.from,
    to: editor.state.selection.to,
  };
  const detached: WorkDocumentCommentView = {
    ...comment,
    from: null,
    to: null,
    anchorText: '',
    detached: true,
  };

  try {
    const view = render(
      <CommentsPanelHarness
        editor={editor}
        comments={[detached]}
        onDelete={() => undefined}
      />,
    );
    const card = view.container.querySelector('[data-comment-id="comment-1"]');
    expect(card).toHaveAttribute('data-document-comment-detached', 'true');
    expect(card).toHaveTextContent('原文锚点已删除');
    fireEvent.click(
      screen.getByRole('button', { name: '查看已脱离正文的批注 1' }),
    );
    expect(editor.state.selection).toMatchObject(selection);
  } finally {
    editor.destroy();
  }
});

function CommentsPanelHarness({
  editor,
  comments = [comment],
  draft = null,
  onCancelDraft = () => undefined,
  onDelete,
}: {
  editor: Editor;
  comments?: WorkDocumentCommentView[];
  draft?: DocumentCommentDraft | null;
  onCancelDraft?: () => void;
  onDelete: (id: string) => void;
}) {
  const surfaceRef = useRef<HTMLDivElement>(null);
  return (
    <div ref={surfaceRef}>
      <DocumentCommentsPanel
        editor={editor}
        comments={comments}
        draft={draft}
        surfaceRef={surfaceRef}
        onReply={() => undefined}
        onToggleResolved={() => undefined}
        onDelete={onDelete}
        onCancelDraft={onCancelDraft}
        onSubmitDraft={() => null}
        onClose={() => undefined}
      />
    </div>
  );
}

function longComments(count: number): WorkDocumentCommentView[] {
  return Array.from({ length: count }, (_, index) => {
    const number = index + 1;
    return {
      id: `comment-${number}`,
      author: 'Reviewer',
      date: '2026-08-02T00:00:00.000Z',
      text: `Review comment ${number}.`,
      resolved: false,
      from: number * 2 - 1,
      to: number * 2,
      anchorText: `Marker ${number}`,
      detached: false,
    };
  });
}

function StatefulLongCommentsHarness({ editor }: { editor: Editor }) {
  const [comments, setComments] = useState(() => longComments(120));
  return (
    <CommentsPanelHarness
      editor={editor}
      comments={comments}
      onDelete={(id) =>
        setComments((current) => current.filter((comment) => comment.id !== id))
      }
    />
  );
}
