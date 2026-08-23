import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { createArtifact, type DocumentContent } from '../src/core';
import { DocumentEditor } from '../src/react';
import { mountWorkLiveDocumentCapture } from '../src/internal/features/work/work-document-page-capture';

test('shares one show-hidden-text view across body and page chrome and suppresses it in preview', async () => {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html =
    '<p>Visible <span data-office-hidden-text="true">body secret</span></p>';
  artifact.content.pageChrome = {
    differentFirstPage: false,
    differentOddEvenPages: false,
    default: {
      headerHtml:
        '<p>Header <span data-office-hidden-text="true">header secret</span></p>',
      footerHtml: '',
      showPageNumber: false,
    },
    first: { headerHtml: '', footerHtml: '', showPageNumber: false },
    even: { headerHtml: '', footerHtml: '', showPageNumber: false },
  };
  const properties = {
    content: artifact.content as DocumentContent,
    onChange: () => undefined,
    theme: 'light' as const,
  };
  const view = render(<DocumentEditor {...properties} preview={false} />);
  const body = await screen.findByRole('textbox', { name: '文档正文' });
  const root = body.closest('.work-document-editor');
  expect(root).not.toBeNull();
  expect(root).not.toHaveClass('show-hidden-text');
  expect(
    root?.querySelectorAll('[data-office-hidden-text="true"]'),
  ).toHaveLength(2);

  fireEvent.click(screen.getByRole('tab', { name: '视图' }));
  const toggle = await screen.findByRole('button', {
    name: '显示隐藏文字',
  });
  expect(toggle).toHaveAttribute('aria-pressed', 'false');
  fireEvent.click(toggle);
  expect(root).toHaveClass('show-hidden-text');
  expect(toggle).toHaveAttribute('aria-pressed', 'true');

  view.rerender(<DocumentEditor {...properties} preview />);
  expect(await screen.findByLabelText('文字预览')).toBeInTheDocument();
  expect(root).toHaveClass('preview');
  expect(root).not.toHaveClass('show-hidden-text');
});

test('detaches PDF snapshots from the editing-only show-hidden-text scope', () => {
  const editorRoot = document.createElement('section');
  editorRoot.className = 'work-document-editor show-hidden-text';
  const page = document.createElement('article');
  page.className = 'work-document-page';
  page.innerHTML =
    '<p>Visible <span data-office-hidden-text="true">PDF secret</span></p>';
  editorRoot.append(page);
  document.body.append(editorRoot);

  const capture = mountWorkLiveDocumentCapture({
    element: page,
    pageCount: 1,
    pageGap: 20,
    pageHeight: 400,
    pageWidth: 300,
  });
  expect(capture.snapshot).toHaveClass('work-document-live-pdf-snapshot');
  expect(capture.snapshot.closest('.show-hidden-text')).toBeNull();
  expect(
    capture.snapshot.querySelector('[data-office-hidden-text="true"]'),
  ).toHaveTextContent('PDF secret');

  capture.host.remove();
  editorRoot.remove();
});
