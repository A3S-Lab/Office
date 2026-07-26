import { expect, test } from '@rstest/core';
import { render, screen } from '@testing-library/react';
import { EditorExportButton } from '../playground/src/editor-workspace';

test('keeps compact export actions named when their text is hidden', () => {
  const view = render(
    <EditorExportButton
      kind="markdown"
      exporting={false}
      onExport={() => undefined}
    />,
  );

  expect(screen.getByRole('button', { name: '导出' })).toHaveAttribute(
    'aria-label',
    '导出',
  );

  view.rerender(
    <EditorExportButton
      kind="pdf"
      exporting={false}
      onExport={() => undefined}
    />,
  );
  expect(screen.getByRole('button', { name: '下载 PDF' })).toHaveAttribute(
    'title',
    '下载 PDF',
  );
});
