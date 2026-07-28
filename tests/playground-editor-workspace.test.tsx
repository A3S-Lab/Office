import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
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

test('keeps Word export formats behind one compact header action', () => {
  let docxExports = 0;
  let pdfExports = 0;
  render(
    <EditorExportButton
      kind="document"
      exporting={false}
      onExport={() => {
        docxExports += 1;
      }}
      onExportPdf={() => {
        pdfExports += 1;
      }}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '导出' }));
  expect(screen.getByRole('menu', { name: '导出格式' })).toBeVisible();
  fireEvent.click(screen.getByRole('menuitem', { name: '导出 PDF' }));
  expect(pdfExports).toBe(1);
  expect(docxExports).toBe(0);

  fireEvent.click(screen.getByRole('button', { name: '导出' }));
  fireEvent.click(screen.getByRole('menuitem', { name: '下载 DOCX' }));
  expect(docxExports).toBe(1);
});
