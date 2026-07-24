import { expect, test } from '@rstest/core';
import {
  appendDocumentCommentReply,
  removeDocumentCommentRecord,
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
