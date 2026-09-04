import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { useState } from 'react';
import type { DocumentContent } from '../src/core';
import { DocumentEditor } from '../src/react';

test('imports a revised file through Writer, publishes revisions, and opens review', async () => {
  const publications: DocumentContent[] = [];
  render(
    <DocumentEditor
      artifactId="document-comparison-integration"
      content={{
        type: 'document',
        pageSize: 'a4',
        html: '<h1>Release review</h1><p>Alpha beta.</p><p>Stable paragraph.</p>',
      }}
      onChange={(content) => publications.push(content)}
      theme="light"
    />,
  );

  const editor = await screen.findByRole('textbox', { name: '文档正文' });
  await waitFor(() =>
    expect(editor).toHaveAttribute('data-document-editor-mount-ms'),
  );
  publications.length = 0;

  const reviewTab = screen.getByRole('tab', { name: '审阅' });
  fireEvent.click(reviewTab);
  await waitFor(() =>
    expect(reviewTab).toHaveAttribute('aria-selected', 'true'),
  );
  const compareButton = await screen.findByRole('button', {
    name: '比较文档',
  });
  compareButton.focus();
  fireEvent.click(compareButton);

  const dialog = await screen.findByRole('dialog', {
    name: '比较与合并文档',
  });
  fireEvent.change(within(dialog).getByLabelText('选择修订版本文件'), {
    target: {
      files: [
        new File(
          [
            '<h1>Release review</h1><p>Alpha gamma.</p><p>Inserted paragraph.</p><p>Stable paragraph.</p>',
          ],
          'release-review.html',
          { type: 'text/html' },
        ),
      ],
    },
  });
  fireEvent.change(
    within(dialog).getByRole('textbox', {
      name: '比较结果修订者名称',
    }),
    { target: { value: 'QA Reviewer' } },
  );
  fireEvent.click(within(dialog).getByRole('button', { name: '生成比较结果' }));

  await waitFor(() =>
    expect(screen.queryByRole('dialog', { name: '比较与合并文档' })).toBeNull(),
  );
  const revisionList = screen.getByRole('list', { name: '待处理修订' });
  const initialCount = Number(
    revisionList.getAttribute('data-document-change-count'),
  );
  expect(initialCount).toBeGreaterThan(0);
  expect(screen.getByText('修订审阅')).toBeInTheDocument();
  await waitFor(() => expect(compareButton).toHaveFocus());

  await waitFor(() => expect(publications.length).toBeGreaterThan(0));
  const compared = publications.at(-1);
  if (!compared) throw new Error('Expected the controlled comparison value.');
  expect(compared.html).toContain('data-document-change="true"');
  expect(compared.html).toContain('data-change-author="QA Reviewer"');

  const publicationCount = publications.length;
  fireEvent.click(
    screen.getByRole('button', {
      name: '接受修订 1',
    }),
  );
  await waitFor(() =>
    expect(
      Number(revisionList.getAttribute('data-document-change-count')),
    ).toBe(initialCount - 1),
  );
  expect(publications.length).toBeGreaterThan(publicationCount);
  if (initialCount > 1) {
    await waitFor(() =>
      expect(screen.getByRole('button', { name: '接受修订 1' })).toHaveFocus(),
    );
  } else {
    await waitFor(() => expect(editor).toHaveFocus());
  }
});

test('surfaces a compared text move as one review card and resolves both sides', async () => {
  const initial: DocumentContent = {
    type: 'document',
    pageSize: 'a4',
    html: '<p>Alpha beta gamma.</p>',
  };
  function ControlledMoveEditor() {
    const [content, setContent] = useState(initial);
    return (
      <DocumentEditor
        artifactId="document-comparison-move-integration"
        content={content}
        onChange={setContent}
        theme="light"
      />
    );
  }
  render(<ControlledMoveEditor />);

  const editor = await screen.findByRole('textbox', { name: '文档正文' });
  await waitFor(() =>
    expect(editor).toHaveAttribute('data-document-editor-mount-ms'),
  );
  const reviewTab = screen.getByRole('tab', { name: '审阅' });
  fireEvent.click(reviewTab);
  const compareButton = await screen.findByRole('button', {
    name: '比较文档',
  });
  compareButton.focus();
  fireEvent.click(compareButton);

  const dialog = await screen.findByRole('dialog', {
    name: '比较与合并文档',
  });
  fireEvent.change(within(dialog).getByLabelText('选择修订版本文件'), {
    target: {
      files: [
        new File(['<p>Alpha gamma beta.</p>'], 'move-review.html', {
          type: 'text/html',
        }),
      ],
    },
  });
  fireEvent.change(
    within(dialog).getByRole('textbox', {
      name: '比较结果修订者名称',
    }),
    { target: { value: 'Move Reviewer' } },
  );
  fireEvent.click(within(dialog).getByRole('button', { name: '生成比较结果' }));

  await waitFor(() =>
    expect(screen.queryByRole('dialog', { name: '比较与合并文档' })).toBeNull(),
  );
  const revisionList = screen.getByRole('list', { name: '待处理修订' });
  expect(revisionList).toHaveAttribute('data-document-change-count', '1');
  const moveCard = revisionList.querySelector(
    '.work-document-change-item.move',
  );
  expect(moveCard).not.toBeNull();
  expect(moveCard).toHaveTextContent('移动');
  expect(moveCard).toHaveTextContent('beta');
  expect(moveCard).toHaveTextContent('Move Reviewer');
  expect(
    editor.querySelector(
      '[data-change-kind="move"][data-change-move-role="from"]',
    ),
  ).not.toBeNull();
  expect(
    editor.querySelector(
      '[data-change-kind="move"][data-change-move-role="to"]',
    ),
  ).not.toBeNull();

  fireEvent.click(
    screen.getByRole('button', { name: '接受修订 1', exact: true }),
  );
  await waitFor(() =>
    expect(screen.queryByRole('list', { name: '待处理修订' })).toBeNull(),
  );
  expect(editor).toHaveTextContent('Alpha gamma beta.');
});
