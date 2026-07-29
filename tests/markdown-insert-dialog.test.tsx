import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { MarkdownInsertDialog } from '../src/internal/features/work/editors/markdown-insert-dialog';

test('cancels link drafts without committing on blur or Escape', () => {
  let closeCount = 0;
  let submitCount = 0;
  render(
    <MarkdownInsertDialog
      request={{
        kind: 'link',
        action: 'insert',
        label: 'Office',
        source: 'https://',
      }}
      restoreFocusTarget={() => null}
      onClose={() => {
        closeCount += 1;
      }}
      onSubmit={() => {
        submitCount += 1;
      }}
    />,
  );

  const source = screen.getByRole('textbox', { name: '链接地址' });
  fireEvent.change(source, { target: { value: 'https://a3s.dev' } });
  fireEvent.blur(source);
  expect(submitCount).toBe(0);

  fireEvent.keyDown(screen.getByRole('dialog', { name: '添加链接' }), {
    key: 'Escape',
  });
  expect(closeCount).toBe(1);
  expect(submitCount).toBe(0);
});

test('commits a valid link form at most once', () => {
  const submissions: string[] = [];
  render(
    <MarkdownInsertDialog
      request={{
        kind: 'link',
        action: 'edit',
        label: 'Office',
        source: 'https://a3s.dev',
      }}
      restoreFocusTarget={() => null}
      onClose={() => undefined}
      onSubmit={(result) => submissions.push(result.source)}
    />,
  );

  const save = screen.getByRole('button', { name: '保存' });
  fireEvent.click(save);
  fireEvent.click(save);
  expect(submissions).toEqual(['https://a3s.dev']);
});

test('keeps blank and invalid link drafts out of the document', () => {
  let submitCount = 0;
  render(
    <MarkdownInsertDialog
      request={{
        kind: 'link',
        action: 'insert',
        label: '',
        source: 'https://',
      }}
      restoreFocusTarget={() => null}
      onClose={() => undefined}
      onSubmit={() => {
        submitCount += 1;
      }}
    />,
  );

  const add = screen.getByRole('button', { name: '添加' });
  expect(add).toBeDisabled();
  const source = screen.getByRole('textbox', { name: '链接地址' });
  fireEvent.change(source, { target: { value: 'javascript:alert(1)' } });
  expect(add).toBeDisabled();
  expect(source).toHaveAttribute('aria-invalid', 'true');
  expect(submitCount).toBe(0);
});
