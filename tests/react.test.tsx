import { expect, test } from '@rstest/core';
import { render, screen } from '@testing-library/react';
import {
  createArtifact,
  type DocumentContent,
  type MarkdownContent,
} from '../src/core';
import {
  defaultPdfiumWasmUrl,
  DocumentEditor,
  MarkdownEditor,
} from '../src/react';

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

test('renders controlled Markdown content with the TipTap editor', () => {
  const artifact = createArtifact('blank-markdown');

  render(
    <MarkdownEditor
      content={
        {
          type: 'markdown',
          markdown: [
            '# A3S Office',
            '',
            'Controlled Markdown.',
            '',
            '| Feature | State |',
            '| --- | --- |',
            '| Tables | Ready |',
          ].join('\n'),
        } satisfies MarkdownContent
      }
      onChange={() => undefined}
      preview
      theme="light"
    />,
  );

  expect(screen.getByLabelText('Markdown 预览')).toHaveTextContent(
    'A3S Office',
  );
  expect(screen.getByRole('table')).toHaveTextContent('Tables');
  expect(artifact.content.type).toBe('markdown');
});
