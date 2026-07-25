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
});
