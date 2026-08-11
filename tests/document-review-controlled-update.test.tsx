import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import type { DocumentContent, DocumentReviewConflictEvent } from '../src/core';
import { DocumentEditor } from '../src/react';

const comment = {
  id: 'comment-1',
  author: 'Reviewer',
  date: '2026-08-11T00:00:00.000Z',
  text: 'Review Alpha.',
  resolved: false,
};

function documentContent(html: string): DocumentContent {
  return {
    type: 'document',
    html,
    pageSize: 'a4',
    comments: [comment],
  };
}

const commentMarkup = (text: string) =>
  `<span data-comment-id="comment-1" data-document-comment="true">${text}</span>`;
const changeMarkup = (text: string) =>
  `<ins data-change-author="Reviewer" data-change-date="2026-08-11T00:00:00.000Z" data-change-id="change-1" data-change-kind="insertion" data-document-change="true">${text}</ins>`;

test('reports review conflicts after applying an authoritative controlled update', async () => {
  const events: DocumentReviewConflictEvent[] = [];
  const changes: DocumentContent[] = [];
  const initial = documentContent(
    `<p>${commentMarkup('Alpha')} ${changeMarkup('beta')}</p>`,
  );
  const shifted = documentContent(
    `<p>Intro ${commentMarkup('Alpha')} ${changeMarkup('beta')}</p>`,
  );
  const overwritten = documentContent(
    `<p>Intro ${commentMarkup('Omega')} beta</p>`,
  );
  const view = render(
    <DocumentEditor
      artifactId="document-1"
      content={initial}
      onChange={(next) => changes.push(next)}
      onReviewConflict={(event) => events.push(event)}
      theme="light"
    />,
  );
  await screen.findByRole('textbox');
  changes.length = 0;

  view.rerender(
    <DocumentEditor
      artifactId="document-1"
      content={shifted}
      onChange={(next) => changes.push(next)}
      onReviewConflict={(event) => events.push(event)}
      theme="light"
    />,
  );
  await waitFor(() =>
    expect(screen.getByRole('textbox')).toHaveTextContent('Intro Alpha beta'),
  );
  expect(events).toEqual([]);

  view.rerender(
    <DocumentEditor
      artifactId="document-1"
      content={overwritten}
      onChange={(next) => changes.push(next)}
      onReviewConflict={(event) => events.push(event)}
      theme="light"
    />,
  );

  const alert = await screen.findByRole('alert');
  expect(alert).toHaveTextContent('2');
  expect(screen.getByRole('textbox')).toHaveTextContent('Intro Omega beta');
  expect(events).toEqual([
    {
      artifactId: 'document-1',
      conflicts: [
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
          reason: 'removed',
          previousText: 'beta',
        },
      ],
    },
  ]);
  expect(changes).toEqual([]);

  fireEvent.click(screen.getByRole('button', { name: '关闭审阅冲突提示' }));
  expect(screen.queryByRole('alert')).toBeNull();
});

test('treats a controlled update for another artifact as a document switch', async () => {
  const events: DocumentReviewConflictEvent[] = [];
  const initial = documentContent(`<p>${commentMarkup('Alpha')}</p>`);
  const view = render(
    <DocumentEditor
      artifactId="document-1"
      content={initial}
      onChange={() => undefined}
      onReviewConflict={(event) => events.push(event)}
      theme="light"
    />,
  );
  await screen.findByRole('textbox');

  view.rerender(
    <DocumentEditor
      artifactId="document-2"
      content={documentContent('<p>A different document</p>')}
      onChange={() => undefined}
      onReviewConflict={(event) => events.push(event)}
      theme="light"
    />,
  );

  await waitFor(() =>
    expect(screen.getByRole('textbox')).toHaveTextContent(
      'A different document',
    ),
  );
  expect(events).toEqual([]);
  expect(screen.queryByRole('alert')).toBeNull();
});
