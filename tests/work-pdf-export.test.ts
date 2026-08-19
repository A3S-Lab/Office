import { afterEach, expect, test } from '@rstest/core';
import {
  workLiveDocumentPdfSurfaceForExport,
  workLiveDocumentPdfCaptureBatches,
  workPdfPagesForExport,
} from '../src/internal/features/work/work-pdf-export';
import { registerDocumentPageSurfaceGeometry } from '../src/internal/features/work/work-document-page-surface-registry';

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

test('keeps exact custom page points on a live pagination surface', () => {
  document.body.innerHTML = `
    <section data-work-pdf-artifact="document-custom" data-work-pdf-surface="live">
      <article
        data-work-pdf-live-document
        data-pdf-orientation="landscape"
        data-pdf-page-count="1"
        data-pdf-page-gap="28"
        data-pdf-page-height="640.4"
        data-pdf-page-height-points="480.3"
        data-pdf-page-size="custom"
        data-pdf-page-width="1067.2"
        data-pdf-page-width-points="800.4"
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

  expect(workLiveDocumentPdfSurfaceForExport('document-custom')).toMatchObject({
    orientation: 'landscape',
    pageHeightPoints: 480.3,
    pageSize: 'a4',
    pageWidthPoints: 800.4,
  });
});

test('reads exact geometry for every mixed-size live document page', () => {
  document.body.innerHTML = `
    <section data-work-pdf-artifact="document-mixed" data-work-pdf-surface="live">
      <article
        data-work-pdf-live-document
        data-pdf-orientation="portrait"
        data-pdf-page-count="2"
        data-pdf-page-gap="20"
        data-pdf-page-height="400"
        data-pdf-page-size="a4"
        data-pdf-page-width="500"
      >
        <div class="work-document-page-stack">
          <div
            data-work-document-page-sheet
            data-page-index="1"
            data-page-top="0"
            data-page-left="100"
            data-page-width="300"
            data-page-height="400"
            data-pdf-orientation="portrait"
            data-pdf-page-size="a4"
            data-pdf-page-width-points="225"
            data-pdf-page-height-points="300"
          ></div>
          <div
            data-work-document-page-sheet
            data-page-index="2"
            data-page-top="420"
            data-page-left="0"
            data-page-width="500"
            data-page-height="200"
            data-pdf-orientation="landscape"
            data-pdf-page-size="letter"
            data-pdf-page-width-points="375"
            data-pdf-page-height-points="150"
          ></div>
        </div>
        <section class="work-document-editable">
          <div
            class="ProseMirror"
            data-pagination-pages="2"
            data-pagination-state="ready"
          ></div>
        </section>
      </article>
    </section>
  `;

  const surface = workLiveDocumentPdfSurfaceForExport('document-mixed');

  expect(surface).toMatchObject({
    pageCount: 2,
    pageGap: 20,
    pageHeight: 400,
    pageWidth: 500,
    pages: [
      {
        height: 400,
        left: 100,
        orientation: 'portrait',
        pageHeightPoints: 300,
        pageWidthPoints: 225,
        top: 0,
        width: 300,
      },
      {
        height: 200,
        left: 0,
        orientation: 'landscape',
        pageHeightPoints: 150,
        pageWidthPoints: 375,
        top: 420,
        width: 500,
      },
    ],
  });
});

test('reads complete registered geometry from a virtualized page stack', () => {
  document.body.innerHTML = `
    <section data-work-pdf-artifact="document-windowed" data-work-pdf-surface="live">
      <article
        data-work-pdf-live-document
        data-pdf-orientation="portrait"
        data-pdf-page-count="2"
        data-pdf-page-gap="20"
        data-pdf-page-height="400"
        data-pdf-page-size="a4"
        data-pdf-page-width="500"
      >
        <div class="work-document-page-stack">
          <div
            data-work-document-page-sheet
            data-page-index="1"
            data-page-top="0"
            data-page-left="100"
            data-page-width="300"
            data-page-height="400"
            data-pdf-orientation="portrait"
            data-pdf-page-size="a4"
            data-pdf-page-width-points="225"
            data-pdf-page-height-points="300"
          ></div>
        </div>
        <section class="work-document-editable">
          <div
            class="ProseMirror"
            data-pagination-pages="2"
            data-pagination-state="ready"
          ></div>
        </section>
      </article>
    </section>
  `;
  const stack = document.querySelector<HTMLElement>(
    '.work-document-page-stack',
  );
  if (!stack) throw new Error('Expected a virtualized page stack.');
  registerDocumentPageSurfaceGeometry(stack, () => [
    {
      height: 400,
      left: 100,
      orientation: 'portrait',
      pageHeightPoints: 300,
      pageIndex: 0,
      pageSize: 'a4',
      pageWidthPoints: 225,
      top: 0,
      width: 300,
    },
    {
      height: 200,
      left: 0,
      orientation: 'landscape',
      pageHeightPoints: 150,
      pageIndex: 1,
      pageSize: 'letter',
      pageWidthPoints: 375,
      top: 420,
      width: 500,
    },
  ]);

  expect(
    workLiveDocumentPdfSurfaceForExport('document-windowed'),
  ).toMatchObject({
    pageCount: 2,
    pages: [
      { height: 400, left: 100, top: 0, width: 300 },
      {
        height: 200,
        left: 0,
        orientation: 'landscape',
        pageSize: 'letter',
        top: 420,
        width: 500,
      },
    ],
  });
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
