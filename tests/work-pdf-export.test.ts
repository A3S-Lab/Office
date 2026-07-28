import { afterEach, expect, test } from '@rstest/core';
import {
  workLiveDocumentPdfSurfaceForExport,
  workLiveDocumentPdfCaptureBatches,
  workPdfPagesForExport,
} from '../src/internal/features/work/work-pdf-export';

afterEach(() => {
  document.body.replaceChildren();
});

test('prefers a ready live document pagination surface', () => {
  document.body.innerHTML = `
    <section data-work-pdf-artifact="document]with-special-id" data-work-pdf-surface="live">
      <article
        data-work-pdf-live-document
        data-pdf-orientation="portrait"
        data-pdf-page-count="3"
        data-pdf-page-gap="28"
        data-pdf-page-height="1122.519685"
        data-pdf-page-size="a4"
        data-pdf-page-width="793.700787"
      >
        <header>
          <div class="ProseMirror">Header editor</div>
        </header>
        <section class="work-document-editable">
          <div
            class="ProseMirror"
            data-pagination-pages="3"
            data-pagination-state="ready"
          ></div>
        </section>
      </article>
    </section>
  `;

  const surface = workLiveDocumentPdfSurfaceForExport(
    'document]with-special-id',
  );

  expect(surface).not.toBeNull();
  expect(surface).toMatchObject({
    orientation: 'portrait',
    pageCount: 3,
    pageGap: 28,
    pageHeight: 1122.519685,
    pageSize: 'a4',
    pageWidth: 793.700787,
  });
});

test('rejects stale or malformed live pagination metadata', () => {
  document.body.innerHTML = `
    <section data-work-pdf-artifact="document-1" data-work-pdf-surface="live">
      <article
        data-work-pdf-live-document
        data-pdf-page-count="2"
        data-pdf-page-gap="28"
        data-pdf-page-height="1123"
        data-pdf-page-size="a4"
        data-pdf-page-width="794"
      >
        <section class="work-document-editable">
          <div
            class="ProseMirror"
            data-pagination-pages="1"
            data-pagination-state="ready"
          ></div>
        </section>
      </article>
    </section>
  `;

  expect(workLiveDocumentPdfSurfaceForExport('document-1')).toBeNull();
  document
    .querySelector<HTMLElement>('.ProseMirror')
    ?.setAttribute('data-pagination-pages', '2');
  document
    .querySelector<HTMLElement>('[data-work-pdf-live-document]')
    ?.setAttribute('data-pdf-page-height', 'not-a-number');
  expect(workLiveDocumentPdfSurfaceForExport('document-1')).toBeNull();
});

test('keeps explicit PDF page selection in physical document order', () => {
  document.body.innerHTML = `
    <section data-work-pdf-artifact="document-2" data-work-pdf-surface="export">
      <article data-work-pdf-page>Page 1</article>
      <article data-work-pdf-page>Page 2</article>
      <article data-work-pdf-page>Page 3</article>
    </section>
  `;

  expect(
    workPdfPagesForExport('document-2', {
      pageIndexes: [2, 0, 2, -1, 9],
    }).map((page) => page.textContent),
  ).toEqual(['Page 1', 'Page 3']);
});

test('captures consecutive live pages in bounded batches', () => {
  expect(
    workLiveDocumentPdfCaptureBatches([0, 1, 2, 3, 4, 5, 7], {
      pageCount: 8,
      pageGap: 28,
      pageHeight: 1123,
    }),
  ).toEqual([
    {
      captureHeight: 5727,
      firstPageIndex: 0,
      pageIndexes: [0, 1, 2, 3, 4],
    },
    {
      captureHeight: 1123,
      firstPageIndex: 5,
      pageIndexes: [5],
    },
    {
      captureHeight: 1123,
      firstPageIndex: 7,
      pageIndexes: [7],
    },
  ]);
});
