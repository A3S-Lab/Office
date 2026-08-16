import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import { act, renderHook } from '@testing-library/react';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import type { WorkDocumentContent } from '../src/internal/features/work/work-types';
import { useDocumentComments } from '../src/internal/features/work/editors/use-document-comments';

describe('document comment module', () => {
  test('owns draft, submit, reply, resolve, and delete orchestration', () => {
    const initial = {
      type: 'document',
      html: '<p>Alpha beta</p>',
      comments: [],
    } satisfies WorkDocumentContent;
    const contentRef = { current: initial as WorkDocumentContent };
    const changes: WorkDocumentContent[] = [];
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        getContent: () => contentRef.current,
        onContentChange: (next) => {
          contentRef.current = next;
          changes.push(next);
        },
      }),
      content: '<p>Alpha beta</p>',
    });
    editor.commands.setTextSelection({ from: 1, to: 6 });
    const { result, rerender, unmount } = renderHook(() =>
      useDocumentComments({
        contentRef,
        editor,
        onBeforeDraft: () => undefined,
      }),
    );

    act(() => result.current.startDraft());
    expect(result.current.open).toBe(true);
    expect(result.current.draft?.anchorText).toBe('Alpha');

    act(() => {
      expect(result.current.submitDraft('Clarify this')).toBeNull();
    });
    rerender();
    const comment = contentRef.current.comments?.[0];
    expect(comment).toMatchObject({
      author: '我',
      text: 'Clarify this',
      resolved: false,
    });
    expect(result.current.draft).toBeNull();
    expect(result.current.comments).toHaveLength(1);

    act(() => result.current.reply(comment?.id ?? '', 'Updated'));
    expect(changes.at(-1)?.comments?.[0].replies?.[0]).toMatchObject({
      author: '我',
      text: 'Updated',
    });

    act(() => result.current.toggleResolved(comment?.id ?? ''));
    expect(changes.at(-1)?.comments?.[0].resolved).toBe(true);

    act(() => result.current.deleteComment(comment?.id ?? ''));
    expect(contentRef.current.comments).toEqual([]);
    expect(
      editor.state.doc.rangeHasMark(1, 6, editor.schema.marks.documentComment),
    ).toBe(false);

    unmount();
    editor.destroy();
  });

  test('attributes collaborative comments and replies to the session actor', () => {
    const initial = {
      type: 'document',
      html: '<p>Alpha beta</p>',
      comments: [],
    } satisfies WorkDocumentContent;
    const contentRef = { current: initial as WorkDocumentContent };
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        getContent: () => contentRef.current,
        onContentChange: (next) => {
          contentRef.current = next;
        },
      }),
      content: '<p>Alpha beta</p>',
    });
    editor.commands.setTextSelection({ from: 1, to: 6 });
    const { result, rerender, unmount } = renderHook(() =>
      useDocumentComments({
        actor: { id: 'ada', name: 'Ada' },
        contentRef,
        deleteOwnOnly: true,
        editor,
      }),
    );

    act(() => result.current.startDraft());
    act(() => expect(result.current.submitDraft('Review this')).toBeNull());
    rerender();
    const comment = contentRef.current.comments?.[0];
    expect(comment).toMatchObject({ actorId: 'ada', author: 'Ada' });

    act(() => result.current.reply(comment?.id ?? '', 'Ada reply'));
    expect(contentRef.current.comments?.[0]?.replies?.[0]).toMatchObject({
      actorId: 'ada',
      author: 'Ada',
    });
    expect(result.current.canDeleteComment(comment?.id ?? '')).toBe(true);

    contentRef.current = {
      ...contentRef.current,
      comments: [
        ...(contentRef.current.comments ?? []),
        {
          id: 'comment-grace',
          actorId: 'grace',
          author: 'Grace',
          date: '2026-08-17T00:00:00.000Z',
          text: 'Foreign thread',
          resolved: false,
        },
      ],
    };
    rerender();
    expect(result.current.canDeleteComment('comment-grace')).toBe(false);
    act(() => result.current.deleteComment('comment-grace'));
    expect(
      contentRef.current.comments?.some(({ id }) => id === 'comment-grace'),
    ).toBe(true);

    unmount();
    editor.destroy();
  });
});
