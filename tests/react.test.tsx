import { expect, test } from '@rstest/core';
import { render, screen } from '@testing-library/react';
import { createArtifact, type DocumentContent } from '../src/core';
import { defaultPdfiumWasmUrl, DocumentEditor } from '../src/react';

test('renders the React document editor in preview mode', () => {
  const artifact = createArtifact('blank-document');

  render(
    <DocumentEditor
      content={artifact.content as DocumentContent}
      onChange={() => undefined}
      preview
      theme="light"
    />,
  );

  expect(screen.getByLabelText('文字预览')).toBeInTheDocument();
  expect(document.querySelector('[data-a3s-office]')).toHaveAttribute(
    'data-theme',
    'light',
  );
});

test('publishes a colocated default PDFium URL', () => {
  expect(new URL(defaultPdfiumWasmUrl).pathname).toMatch(/\/pdfium\.wasm$/);
});
