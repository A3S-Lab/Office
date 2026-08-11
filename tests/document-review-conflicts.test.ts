import { describe, expect, test } from '@rstest/core';
import {
  reconcileDocumentReviewConflicts,
  type WorkDocumentReviewSnapshot,
} from '../src/internal/features/work/work-document-review-conflicts';
import type { WorkDocumentComment } from '../src/internal/features/work/work-types';

const commentRecord = (id: string): WorkDocumentComment => ({
  id,
  author: 'Reviewer',
  date: '2026-08-11T00:00:00.000Z',
  text: `Thread for ${id}`,
  resolved: false,
});

const snapshot = (
  overrides: Partial<WorkDocumentReviewSnapshot> = {},
): WorkDocumentReviewSnapshot => ({
  comments: [],
  changes: [],
  ...overrides,
});

describe('controlled document review conflicts', () => {
  test('ignores position-only movement and newly introduced review items', () => {
    const result = reconcileDocumentReviewConflicts({
      active: [],
      before: snapshot({
        comments: [{ id: 'comment-1', from: 1, to: 6, anchorText: 'Alpha' }],
        changes: [
          {
            id: 'change-1',
            kind: 'insertion',
            author: 'Reviewer',
            date: '',
            from: 7,
            to: 11,
            text: 'beta',
          },
        ],
      }),
      after: snapshot({
        comments: [
          { id: 'comment-1', from: 20, to: 25, anchorText: 'Alpha' },
          { id: 'comment-2', from: 30, to: 35, anchorText: 'Gamma' },
        ],
        changes: [
          {
            id: 'change-1',
            kind: 'insertion',
            author: 'Reviewer',
            date: '',
            from: 26,
            to: 30,
            text: 'beta',
          },
          {
            id: 'change-2',
            kind: 'deletion',
            author: 'Reviewer',
            date: '',
            from: 36,
            to: 41,
            text: 'delta',
          },
        ],
      }),
      comments: [commentRecord('comment-1'), commentRecord('comment-2')],
    });

    expect(result).toEqual({ active: [], detected: [] });
  });

  test('reports changed and removed ranges while honoring an intentional comment deletion', () => {
    const result = reconcileDocumentReviewConflicts({
      active: [],
      before: snapshot({
        comments: [
          { id: 'comment-1', from: 1, to: 6, anchorText: 'Alpha' },
          { id: 'comment-2', from: 7, to: 11, anchorText: 'beta' },
        ],
        changes: [
          {
            id: 'change-1',
            kind: 'insertion',
            author: 'Reviewer',
            date: '',
            from: 12,
            to: 17,
            text: 'Gamma',
          },
          {
            id: 'change-2',
            kind: 'deletion',
            author: 'Reviewer',
            date: '',
            from: 18,
            to: 23,
            text: 'delta',
          },
        ],
      }),
      after: snapshot({
        comments: [{ id: 'comment-1', from: 1, to: 6, anchorText: 'Omega' }],
        changes: [
          {
            id: 'change-1',
            kind: 'deletion',
            author: 'Reviewer',
            date: '',
            from: 12,
            to: 17,
            text: 'Gamma',
          },
        ],
      }),
      comments: [commentRecord('comment-1')],
    });

    expect(result.active).toEqual([
      {
        id: 'comment-1',
        kind: 'comment',
        reason: 'text-changed',
        previousText: 'Alpha',
        nextText: 'Omega',
      },
      {
        id: 'change-1',
        kind: 'insertion',
        reason: 'kind-changed',
        previousText: 'Gamma',
        nextKind: 'deletion',
        nextText: 'Gamma',
      },
      {
        id: 'change-2',
        kind: 'deletion',
        reason: 'removed',
        previousText: 'delta',
      },
    ]);
    expect(result.detected).toEqual(result.active);
  });

  test('keeps unresolved conflicts without reporting them twice and clears restored ranges', () => {
    const first = reconcileDocumentReviewConflicts({
      active: [],
      before: snapshot({
        comments: [{ id: 'comment-1', from: 1, to: 6, anchorText: 'Alpha' }],
      }),
      after: snapshot(),
      comments: [commentRecord('comment-1')],
    });
    const stillMissing = reconcileDocumentReviewConflicts({
      active: first.active,
      before: snapshot(),
      after: snapshot(),
      comments: [commentRecord('comment-1')],
    });
    const restored = reconcileDocumentReviewConflicts({
      active: stillMissing.active,
      before: snapshot(),
      after: snapshot({
        comments: [{ id: 'comment-1', from: 15, to: 20, anchorText: 'Alpha' }],
      }),
      comments: [commentRecord('comment-1')],
    });

    expect(first.detected).toHaveLength(1);
    expect(stillMissing.active).toEqual(first.active);
    expect(stillMissing.detected).toEqual([]);
    expect(restored).toEqual({ active: [], detected: [] });
  });

  test('reports a removed comment thread when its body anchor remains', () => {
    const before = snapshot({
      comments: [{ id: 'comment-1', from: 1, to: 6, anchorText: 'Alpha' }],
    });
    const orphaned = reconcileDocumentReviewConflicts({
      active: [],
      before,
      after: before,
      comments: [],
    });
    const intentionallyDeleted = reconcileDocumentReviewConflicts({
      active: orphaned.active,
      before,
      after: snapshot(),
      comments: [],
    });

    expect(orphaned.active).toEqual([
      {
        id: 'comment-1',
        kind: 'comment',
        reason: 'record-removed',
        previousText: 'Alpha',
        nextText: 'Alpha',
      },
    ]);
    expect(intentionallyDeleted).toEqual({ active: [], detected: [] });
  });

  test('keeps the original recovery baseline while a conflict remains active', () => {
    const alpha = snapshot({
      comments: [{ id: 'comment-1', from: 1, to: 6, anchorText: 'Alpha' }],
    });
    const omega = snapshot({
      comments: [{ id: 'comment-1', from: 1, to: 6, anchorText: 'Omega' }],
    });
    const delta = snapshot({
      comments: [{ id: 'comment-1', from: 1, to: 6, anchorText: 'Delta' }],
    });
    const comments = [commentRecord('comment-1')];
    const first = reconcileDocumentReviewConflicts({
      active: [],
      before: alpha,
      after: omega,
      comments,
    });
    const changedAgain = reconcileDocumentReviewConflicts({
      active: first.active,
      before: omega,
      after: delta,
      comments,
    });
    const latestValueRestored = reconcileDocumentReviewConflicts({
      active: changedAgain.active,
      before: delta,
      after: omega,
      comments,
    });
    const originalValueRestored = reconcileDocumentReviewConflicts({
      active: latestValueRestored.active,
      before: omega,
      after: alpha,
      comments,
    });

    expect(changedAgain.active).toEqual(first.active);
    expect(changedAgain.detected).toEqual([]);
    expect(latestValueRestored.active).toEqual(first.active);
    expect(originalValueRestored).toEqual({ active: [], detected: [] });
  });

  test('deduplicates a revision conflict across later kind and text changes', () => {
    const change = (
      kind: 'insertion' | 'deletion',
      text: string,
    ): WorkDocumentReviewSnapshot =>
      snapshot({
        changes: [
          {
            id: 'change-1',
            kind,
            author: 'Reviewer',
            date: '',
            from: 1,
            to: 1 + text.length,
            text,
          },
        ],
      });
    const insertion = change('insertion', 'Alpha');
    const deletion = change('deletion', 'Alpha');
    const changedDeletion = change('deletion', 'Omega');
    const first = reconcileDocumentReviewConflicts({
      active: [],
      before: insertion,
      after: deletion,
      comments: [],
    });
    const changedAgain = reconcileDocumentReviewConflicts({
      active: first.active,
      before: deletion,
      after: changedDeletion,
      comments: [],
    });
    const restored = reconcileDocumentReviewConflicts({
      active: changedAgain.active,
      before: changedDeletion,
      after: insertion,
      comments: [],
    });

    expect(changedAgain.active).toEqual(first.active);
    expect(changedAgain.detected).toEqual([]);
    expect(restored).toEqual({ active: [], detected: [] });
  });
});
