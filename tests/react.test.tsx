import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { StrictMode } from 'react';
import {
  createArtifact,
  type DocumentContent,
  type MarkdownContent,
} from '../src/core';
import {
  defaultDocumentArabicLayoutFontUrl,
  defaultDocumentHebrewLayoutFontUrl,
  defaultDocumentLatinLayoutFontUrl,
  defaultDocumentLayoutFonts,
  defaultDocumentLayoutFontUrl,
  defaultOfficeKernelWasmUrl,
  defaultPdfiumWasmUrl,
  DocumentEditor,
  MarkdownEditor,
  preloadOfficeEditor,
} from '../src/react';
import { createWorkDocumentModel } from '../src/internal/features/work/work-document-model';
import {
  documentInitialSectionLayout,
  documentSectionNodeAttributes,
  normalizeDocumentHtml,
} from '../src/internal/features/work/work-document-section';

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

test('publishes the exact default document layout font stack', () => {
  expect(new URL(defaultDocumentLatinLayoutFontUrl).pathname).toMatch(
    /\/noto-sans-regular\.ttf$/,
  );
  expect(new URL(defaultDocumentLayoutFontUrl).pathname).toMatch(
    /\/noto-sans-hans-regular\.otf$/,
  );
  expect(new URL(defaultDocumentArabicLayoutFontUrl).pathname).toMatch(
    /\/noto-naskh-arabic-regular\.ttf$/,
  );
  expect(new URL(defaultDocumentHebrewLayoutFontUrl).pathname).toMatch(
    /\/noto-sans-hebrew-regular\.ttf$/,
  );
  expect(defaultDocumentLayoutFonts).toEqual([
    expect.objectContaining({
      family: 'A3S Office Noto Sans',
      id: 'noto-sans-regular',
      url: defaultDocumentLatinLayoutFontUrl,
      weight: 400,
    }),
    expect.objectContaining({
      family: 'A3S Office Noto Sans Hans',
      id: 'noto-sans-hans-regular',
      url: defaultDocumentLayoutFontUrl,
      weight: 400,
    }),
    expect.objectContaining({
      family: 'A3S Office Noto Naskh Arabic',
      id: 'noto-naskh-arabic-regular',
      url: defaultDocumentArabicLayoutFontUrl,
      weight: 400,
    }),
    expect.objectContaining({
      family: 'A3S Office Noto Sans Hebrew',
      id: 'noto-sans-hebrew-regular',
      url: defaultDocumentHebrewLayoutFontUrl,
      weight: 400,
    }),
  ]);
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
  expect(screen.getByLabelText('文档正文')).toHaveAttribute(
    'data-pagination-document-revision',
    '0',
  );
  expect(screen.getByLabelText('文档正文')).toHaveAttribute(
    'data-pagination-measured-blocks',
    '1',
  );
  expect(screen.getByLabelText('文档正文')).toHaveAttribute(
    'data-pagination-reused-blocks',
    '0',
  );
  expect(screen.getByRole('combobox', { name: '字体' })).toHaveTextContent(
    '默认字体',
  );
  expect(screen.getByRole('combobox', { name: '字号' })).toHaveTextContent(
    '10.5',
  );
  expect(screen.getByRole('combobox', { name: '行距' })).toBeInTheDocument();
  expect(screen.getByRole('button', { name: '增加缩进' })).toBeInTheDocument();
  expect(screen.getByRole('button', { name: '两端对齐' })).toBeInTheDocument();
  expect(screen.getByRole('button', { name: '从左向右' })).toBeInTheDocument();
  expect(screen.getByRole('button', { name: '从右向左' })).toBeInTheDocument();
});

test('edits page chrome directly on the paper with a contextual ribbon', async () => {
  const artifact = createArtifact('blank-document');

  render(
    <DocumentEditor
      content={artifact.content as DocumentContent}
      onChange={() => undefined}
      theme="light"
    />,
  );

  fireEvent.click(
    await screen.findByRole('button', {
      name: '编辑页眉',
    }),
  );
  expect(
    await screen.findByRole('textbox', { name: '页内页眉' }),
  ).toBeInTheDocument();
  await waitFor(() => {
    expect(screen.getByRole('tab', { name: '页眉和页脚' })).toHaveAttribute(
      'aria-selected',
      'true',
    );
  });

  fireEvent.click(screen.getByRole('button', { name: '切换到页脚' }));
  expect(
    await screen.findByRole('textbox', { name: '页内页脚' }),
  ).toBeInTheDocument();
  expect(
    screen.queryByRole('textbox', { name: '页内页眉' }),
  ).not.toBeInTheDocument();

  fireEvent.click(screen.getByRole('button', { name: '关闭页眉和页脚' }));
  await waitFor(() => {
    expect(
      screen.queryByRole('textbox', { name: '页内页脚' }),
    ).not.toBeInTheDocument();
  });
  expect(screen.getByRole('button', { name: '编辑页脚' })).toBeInTheDocument();
});

test('edits the page-chrome variant resolved for the current physical page', async () => {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document')
    throw new Error('Expected a document artifact.');
  const changes: DocumentContent[] = [];
  const content: DocumentContent = {
    ...artifact.content,
    pageChrome: {
      differentFirstPage: true,
      differentOddEvenPages: true,
      default: {
        headerHtml: '<p>Default header</p>',
        footerHtml: '',
        showPageNumber: false,
      },
      first: {
        headerHtml: '<p>First page header</p>',
        footerHtml: '',
        showPageNumber: false,
      },
      even: {
        headerHtml: '<p>Even page header</p>',
        footerHtml: '',
        showPageNumber: false,
      },
    },
  };

  render(
    <DocumentEditor
      content={content}
      onChange={(next) => changes.push(next)}
      theme="light"
    />,
  );

  expect(await screen.findByText('First page header')).toBeInTheDocument();
  fireEvent.click(screen.getByRole('button', { name: '编辑页眉' }));
  fireEvent.click(
    await screen.findByRole('button', {
      name: '显示页码',
    }),
  );

  await waitFor(() => {
    expect(changes.at(-1)?.pageChrome?.first.showPageNumber).toBe(true);
  });
  expect(changes.at(-1)?.pageChrome?.default.showPageNumber).toBe(false);
});

test('uses the structured document model and falls back to host HTML when it is stale', async () => {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document')
    throw new Error('Expected a document artifact.');
  const normalizedHtml = normalizeDocumentHtml(artifact.content);
  const content: DocumentContent = {
    ...artifact.content,
    html: normalizedHtml,
    model: createWorkDocumentModel(normalizedHtml, {
      type: 'doc',
      content: [
        {
          type: 'documentSection',
          attrs: documentSectionNodeAttributes(
            documentInitialSectionLayout(artifact.content),
            'document-section-1',
          ),
          content: [
            {
              type: 'paragraph',
              content: [{ type: 'text', text: 'Structured model' }],
            },
          ],
        },
      ],
    }),
  };
  const { rerender } = render(
    <DocumentEditor
      content={content}
      onChange={() => undefined}
      theme="light"
    />,
  );

  expect(await screen.findByLabelText('文档正文')).toHaveTextContent(
    'Structured model',
  );
  await waitFor(() => {
    expect(screen.getByLabelText('文档正文')).toHaveAttribute(
      'data-pagination-document-revision',
      '1',
    );
  });

  rerender(
    <DocumentEditor
      content={{ ...content, html: '<p>Host HTML override</p>' }}
      onChange={() => undefined}
      theme="light"
    />,
  );

  await waitFor(() => {
    expect(screen.getByLabelText('文档正文')).toHaveTextContent(
      'Host HTML override',
    );
  });
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

test('opens Markdown in source-and-preview split mode by default', async () => {
  const artifact = createArtifact('blank-markdown');
  if (artifact.content.type !== 'markdown') {
    throw new Error('Expected a Markdown artifact.');
  }

  const { container } = render(
    <MarkdownEditor
      content={artifact.content}
      onChange={() => undefined}
      theme="light"
    />,
  );

  expect(await screen.findByLabelText('Markdown 源码')).toBeInTheDocument();
  expect(screen.getByLabelText('Markdown 编辑区')).toBeInTheDocument();
  expect(
    container.querySelector('.work-markdown-workspace.split'),
  ).toBeInTheDocument();
});
