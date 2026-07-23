import { expect, test } from '@rstest/core';
import { render, screen, waitFor } from '@testing-library/react';
import { StrictMode } from 'react';
import {
  createArtifact,
  type DocumentContent,
  type MarkdownContent,
} from '../src/core';
import {
  defaultOfficeKernelWasmUrl,
  defaultPdfiumWasmUrl,
  DocumentEditor,
  MarkdownEditor,
  preloadOfficeEditor,
} from '../src/react';

test('renders the React document editor in preview mode', async () => {
  const artifact = createArtifact('blank-document');

  render(
    <DocumentEditor
      content={artifact.content as DocumentContent}
      onChange={() => undefined}
      preview
      theme="light"
    />,
  );

  expect(await screen.findByLabelText('文字预览')).toBeInTheDocument();
  expect(document.querySelector('[data-a3s-office]')).toHaveAttribute(
    'data-theme',
    'light',
  );
});

test('preloads an editor without mounting it', async () => {
  await expect(preloadOfficeEditor('document')).resolves.toBeUndefined();
});

test('publishes a colocated default PDFium URL', () => {
  expect(new URL(defaultPdfiumWasmUrl).pathname).toMatch(/\/pdfium\.wasm$/);
});

test('publishes a colocated default Office kernel URL', () => {
  expect(new URL(defaultOfficeKernelWasmUrl).pathname).toMatch(
    /\/office-kernel\.wasm$/,
  );
});

test('keeps document pagination available under React strict effects', async () => {
  const artifact = createArtifact('blank-document');

  render(
    <StrictMode>
      <DocumentEditor
        content={artifact.content as DocumentContent}
        onChange={() => undefined}
        theme="light"
      />
    </StrictMode>,
  );

  await waitFor(() => {
    expect(screen.getByLabelText('文档正文')).toHaveAttribute(
      'data-pagination-state',
      'ready',
    );
  });
  expect(screen.getByLabelText('文档正文')).toHaveAttribute(
    'data-pagination-engine',
    'javascript',
  );
});

test('renders controlled Markdown content with the TipTap editor', async () => {
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

  expect(await screen.findByLabelText('Markdown 预览')).toHaveTextContent(
    'A3S Office',
  );
  expect(screen.getByRole('table')).toHaveTextContent('Tables');
  expect(artifact.content.type).toBe('markdown');
});
