import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { DocumentTaskPane } from '../src/internal/features/work/editors/document-task-pane';

test('describes a task pane and closes it with Escape from a field', () => {
  const closed: boolean[] = [];
  render(
    <DocumentTaskPane
      className="example-pane"
      title="页面设置"
      description="第 2 节 · 共 4 节"
      closeLabel="关闭页面设置"
      onClose={() => closed.push(true)}
    >
      <input aria-label="纸张大小" />
    </DocumentTaskPane>,
  );

  const pane = screen.getByRole('complementary', { name: '页面设置' });
  expect(pane).toHaveAccessibleDescription('第 2 节 · 共 4 节');
  fireEvent.keyDown(screen.getByRole('textbox', { name: '纸张大小' }), {
    key: 'Escape',
  });
  expect(closed).toEqual([true]);
});
