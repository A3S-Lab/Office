import { expect, test } from '@rstest/core';
import {
  commentConnectorMutationNeedsLayout,
  syncDocumentCommentMarkClasses,
} from '../src/internal/features/work/editors/document-comments-panel';
import type { WorkDocumentCommentView } from '../src/internal/features/work/work-document-comments';

test('syncDocumentCommentMarkClasses only writes when class state changes', () => {
  const root = document.createElement('div');
  const active = document.createElement('span');
  active.setAttribute('data-document-comment', 'true');
  active.setAttribute('data-comment-id', 'a');
  active.classList.add('is-active-comment');
  const resolved = document.createElement('span');
  resolved.setAttribute('data-document-comment', 'true');
  resolved.setAttribute('data-comment-id', 'b');
  resolved.classList.add('is-resolved-comment');
  root.append(active, resolved);

  const commentsById = new Map<string, WorkDocumentCommentView>([
    [
      'a',
      {
        id: 'a',
        author: 'Reviewer',
        date: '2026-01-01T00:00:00.000Z',
        text: 'A',
        resolved: false,
        from: 1,
        to: 2,
        anchorText: 'A',
        detached: false,
      },
    ],
    [
      'b',
      {
        id: 'b',
        author: 'Reviewer',
        date: '2026-01-01T00:00:00.000Z',
        text: 'B',
        resolved: true,
        from: 3,
        to: 4,
        anchorText: 'B',
        detached: false,
      },
    ],
  ]);

  const activeToggle = spyToggle(active);
  const resolvedToggle = spyToggle(resolved);

  syncDocumentCommentMarkClasses(root, 'a', commentsById);
  expect(activeToggle.calls).toEqual([]);
  expect(resolvedToggle.calls).toEqual([]);

  syncDocumentCommentMarkClasses(root, 'b', commentsById);
  expect(activeToggle.calls).toEqual([['is-active-comment', false]]);
  expect(resolvedToggle.calls).toEqual([['is-active-comment', true]]);
  expect(active.classList.contains('is-active-comment')).toBe(false);
  expect(resolved.classList.contains('is-active-comment')).toBe(true);
  expect(resolved.classList.contains('is-resolved-comment')).toBe(true);
});

test('commentConnectorMutationNeedsLayout ignores active/resolved class churn', () => {
  const mark = document.createElement('span');
  mark.setAttribute('data-document-comment', 'true');
  mark.setAttribute('data-comment-id', 'a');
  const records: MutationRecord[] = [
    {
      type: 'attributes',
      target: mark,
      attributeName: 'class',
      attributeNamespace: null,
      oldValue: null,
      addedNodes: [] as unknown as NodeList,
      removedNodes: [] as unknown as NodeList,
      previousSibling: null,
      nextSibling: null,
    } as MutationRecord,
  ];
  expect(commentConnectorMutationNeedsLayout(records)).toBe(false);
});

function spyToggle(element: HTMLElement): { calls: Array<[string, boolean]> } {
  const calls: Array<[string, boolean]> = [];
  const original = element.classList.toggle.bind(element.classList);
  element.classList.toggle = ((token: string, force?: boolean) => {
    if (typeof force === 'boolean') calls.push([token, force]);
    return original(token, force);
  }) as DOMTokenList['toggle'];
  return { calls };
}
