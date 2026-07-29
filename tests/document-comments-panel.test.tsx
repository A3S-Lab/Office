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
import { useRef } from 'react';
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

    fireEvent.change(screen.getByRole('textbox', { name: '回复批注 1' }), {
      target: { value: 'Unsaved reply' },
    });
    fireEvent.click(screen.getByRole('button', { name: '删除批注 1' }));

    const dialog = screen.getByRole('dialog', { name: '删除批注？' });
    expect(dialog).toHaveAccessibleDescription(
      '批注、已有回复和未发送的回复都将删除。',
    );
    expect(deleted).toEqual([]);

    fireEvent.click(within(dialog).getByRole('button', { name: '取消' }));
    expect(deleted).toEqual([]);

    fireEvent.click(screen.getByRole('button', { name: '删除批注 1' }));
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
