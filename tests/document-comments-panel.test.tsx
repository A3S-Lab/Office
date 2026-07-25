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
import type { WorkDocumentCommentView } from '../src/internal/features/work/work-document-comments';

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

function CommentsPanelHarness({
  editor,
  onDelete,
}: {
  editor: Editor;
  onDelete: (id: string) => void;
}) {
  const surfaceRef = useRef<HTMLDivElement>(null);
  return (
    <div ref={surfaceRef}>
      <DocumentCommentsPanel
        editor={editor}
        comments={[comment]}
        draft={null}
        surfaceRef={surfaceRef}
        onReply={() => undefined}
        onToggleResolved={() => undefined}
        onDelete={onDelete}
        onCancelDraft={() => undefined}
        onSubmitDraft={() => null}
        onClose={() => undefined}
      />
    </div>
  );
}
