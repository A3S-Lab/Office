import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import { createArtifact } from '../src/core';
import type { WorkDocumentContent } from '../src/internal/features/work/work-types';
import { DocumentEditor } from '../src/react';

function ReviewDocument({ initial }: { initial: WorkDocumentContent }) {
  const [content, setContent] = useState(initial);
  return (
    <DocumentEditor content={content} onChange={setContent} theme="light" />
  );
}

test('accepts or processes one review comment and withdraws it from the queue', async () => {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const content: WorkDocumentContent = {
    ...artifact.content,
    html: '<section data-document-section="true"><p><span data-document-comment="true" data-comment-id="finding-1">Alpha</span> and <span data-document-comment="true" data-comment-id="finding-2">Beta</span></p></section>',
    comments: [
      {
        id: 'finding-1',
        author: '审阅',
        date: '2026-09-29T00:00:00.000Z',
        text: '第一条发现',
        resolved: false,
      },
      {
        id: 'finding-2',
        author: '审阅',
        date: '2026-09-29T00:00:00.000Z',
        text: '第二条发现',
        resolved: false,
      },
    ],
  };
  const { unmount } = render(<ReviewDocument initial={content} />);

  fireEvent.click(await screen.findByRole('tab', { name: '审阅' }));
  const withdraw = await screen.findByRole('button', { name: '撤回' });
  expect(withdraw).toBeDisabled();
  expect(withdraw).toHaveAttribute(
    'title',
    '还没有已接受或已处理的批注，不能撤回',
  );
  expect(screen.getByRole('button', { name: '接受' })).toBeEnabled();
  expect(screen.getByRole('button', { name: '处理' })).toBeEnabled();

  fireEvent.click(screen.getByRole('button', { name: '查看批注（2）' }));
  expect(await screen.findByLabelText('批注审阅')).toHaveTextContent(
    '2 条待处理 · 共 2 条',
  );
  expect(
    screen.getByRole('button', { name: '定位批注 1' }),
  ).toBeInTheDocument();

  fireEvent.click(screen.getByRole('button', { name: '接受批注 1' }));
  await waitFor(() =>
    expect(screen.getByLabelText('批注审阅')).toHaveTextContent(
      '1 条待处理 · 共 2 条',
    ),
  );
  expect(screen.getByText('已接受')).toBeInTheDocument();

  fireEvent.click(screen.getByRole('button', { name: '撤回批注 1' }));
  await waitFor(() =>
    expect(screen.getByLabelText('批注审阅')).toHaveTextContent(
      '2 条待处理 · 共 2 条',
    ),
  );
  expect(screen.queryByText('已接受')).not.toBeInTheDocument();

  fireEvent.click(screen.getByRole('button', { name: '处理' }));
  await waitFor(() =>
    expect(screen.getByLabelText('批注审阅')).toHaveTextContent(
      '1 条待处理 · 共 2 条',
    ),
  );
  expect(screen.getByText('已处理')).toBeInTheDocument();

  const ribbonWithdraw = screen.getByRole('button', { name: '撤回' });
  expect(ribbonWithdraw).toBeEnabled();
  fireEvent.click(ribbonWithdraw);
  await waitFor(() =>
    expect(screen.getByLabelText('批注审阅')).toHaveTextContent(
      '2 条待处理 · 共 2 条',
    ),
  );
  expect(screen.queryByText('已处理')).not.toBeInTheDocument();
  unmount();
});
