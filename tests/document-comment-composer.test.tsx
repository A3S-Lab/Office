import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { DocumentCommentComposer } from '../src/internal/features/work/editors/document-comment-composer';

const draft = {
  id: 'comment-draft-1',
  from: 1,
  to: 6,
  anchorText: 'Alpha',
};

test('submits a focused comment draft without a modal backdrop', () => {
  const submitted: string[] = [];

  render(
    <DocumentCommentComposer
      draft={draft}
      top={48}
      onCancel={() => undefined}
      onSubmit={(text) => {
        submitted.push(text);
        return null;
      }}
    />,
  );

  const composer = screen.getByRole('dialog', { name: '添加批注' });
  const input = screen.getByRole('textbox', { name: '批注内容' });
  expect(composer).toHaveStyle({ top: '48px' });
  expect(input).toHaveFocus();
  expect(screen.getByText('Alpha')).toBeVisible();
  expect(document.querySelector('.ds-dialog-backdrop')).toBeNull();
  expect(screen.getByRole('button', { name: '添加批注' })).toHaveClass(
    'ds-button',
    'primary',
    'compact',
  );
  expect(screen.getByRole('button', { name: '取消' })).toHaveClass(
    'ds-button',
    'quiet',
    'compact',
  );

  fireEvent.change(input, { target: { value: '  Clarify this.  ' } });
  fireEvent.keyDown(input, { key: 'Enter', ctrlKey: true });
  expect(submitted).toEqual(['Clarify this.']);
});

test('keeps an invalid draft open and supports Escape to cancel', () => {
  const cancelled: boolean[] = [];

  render(
    <DocumentCommentComposer
      draft={draft}
      top={0}
      onCancel={() => cancelled.push(true)}
      onSubmit={() => '所选文字已变化，请重新选择。'}
    />,
  );

  const input = screen.getByRole('textbox', { name: '批注内容' });
  fireEvent.change(input, { target: { value: 'Comment' } });
  fireEvent.click(screen.getByRole('button', { name: '添加批注' }));
  expect(screen.getByRole('alert')).toHaveTextContent(
    '所选文字已变化，请重新选择。',
  );

  fireEvent.keyDown(input, { key: 'Escape' });
  expect(cancelled).toEqual([true]);
});
