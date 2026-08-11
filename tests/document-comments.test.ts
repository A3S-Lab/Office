import { Editor } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  appendDocumentCommentReply,
  canInsertDocumentComment,
  documentCommentDraftRange,
  removeDocumentCommentRecord,
  retainAnchoredDocumentComments,
  toggleDocumentCommentResolved,
} from '../src/internal/features/work/work-document-comments';
import type { WorkDocumentComment } from '../src/internal/features/work/work-types';

const comment: WorkDocumentComment = {
  id: 'comment-1',
  author: 'Alice',
  date: '2026-07-24T04:00:00.000Z',
  text: 'Clarify this requirement.',
  resolved: false,
};

test('updates document comment threads without mutating the source', () => {
  const source = [comment];
  const withReply = appendDocumentCommentReply(source, comment.id, {
    id: 'reply-1',
    author: 'Bob',
    date: '2026-07-24T04:05:00.000Z',
    text: 'Updated.',
  });
  const resolved = toggleDocumentCommentResolved(withReply, comment.id);
  const removed = removeDocumentCommentRecord(resolved, comment.id);

  expect(source).toEqual([comment]);
  expect(withReply[0]?.replies).toEqual([
    {
      id: 'reply-1',
      author: 'Bob',
      date: '2026-07-24T04:05:00.000Z',
      text: 'Updated.',
    },
  ]);
  expect(resolved[0]?.resolved).toBe(true);
  expect(removed).toEqual([]);
});

test('leaves unrelated document comments unchanged', () => {
  const source = [comment];

  expect(
    appendDocumentCommentReply(source, 'missing', {
      id: 'reply-1',
      author: 'Bob',
      date: '',
      text: 'No target.',
    }),
  ).toEqual(source);
  expect(toggleDocumentCommentResolved(source, 'missing')).toEqual(source);
  expect(removeDocumentCommentRecord(source, 'missing')).toEqual(source);
});

test('retains an explicitly conflicted comment whose anchor is temporarily missing', () => {
  const source = [comment, { ...comment, id: 'comment-2' }];

  const retained = retainAnchoredDocumentComments(
    source,
    [
      {
        id: 'comment-2',
        from: 1,
        to: 6,
        anchorText: 'Alpha',
      },
    ],
    new Set(['comment-1']),
  );

  expect(retained.map((item) => item.id)).toEqual(['comment-1', 'comment-2']);
});

test('keeps a comment draft anchored to its original text range', () => {
  const element = document.createElement('div');
  document.body.append(element);
  const editor = new Editor({
    element,
    extensions: createWorkDocumentExtensions(),
    content: '<p>Alpha beta gamma</p>',
  });

  try {
    expect(canInsertDocumentComment(editor)).toBe(false);
    editor.commands.setTextSelection({ from: 1, to: 6 });
    expect(canInsertDocumentComment(editor)).toBe(true);
    expect(editor.commands.showDocumentCommentDraft({ from: 1, to: 6 })).toBe(
      true,
    );
    expect(documentCommentDraftRange(editor)).toEqual({ from: 1, to: 6 });
    expect(
      editor.view.dom.querySelector('[data-document-comment-draft]'),
    ).toHaveTextContent('Alpha');

    editor.commands.setTextSelection({ from: 7, to: 11 });
    expect(
      editor.commands.insertDocumentComment({
        id: 'comment-1',
        range: documentCommentDraftRange(editor) ?? undefined,
      }),
    ).toBe(true);
    expect(editor.getHTML()).toContain(
      '<span data-comment-id="comment-1" data-document-comment="true">Alpha</span>',
    );
    expect(editor.getHTML()).not.toContain(
      'data-document-comment="true">beta</span>',
    );
    expect(documentCommentDraftRange(editor)).toBeNull();

    editor.commands.setTextSelection({ from: 1, to: 6 });
    expect(canInsertDocumentComment(editor)).toBe(false);
    expect(editor.commands.removeDocumentComment('comment-1')).toBe(true);
    expect(editor.getHTML()).not.toContain('data-document-comment="true"');

    editor.commands.setTextSelection({ from: 7, to: 11 });
    expect(editor.commands.showDocumentCommentDraft({ from: 7, to: 11 })).toBe(
      true,
    );
    expect(editor.commands.clearDocumentCommentDraft()).toBe(true);
    expect(documentCommentDraftRange(editor)).toBeNull();
  } finally {
    editor.destroy();
    element.remove();
  }
});
