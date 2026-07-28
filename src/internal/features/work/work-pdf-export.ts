import type { jsPDF as JsPdf } from 'jspdf';
import type { WorkArtifact, WorkSpreadsheetPaperSize } from './work-types';

type PdfPageSize = WorkSpreadsheetPaperSize;

export interface WorkPdfExportOptions {
  pageIndexes?: number[];
}

export interface WorkLiveDocumentPdfSurface {
  element: HTMLElement;
  orientation: 'portrait' | 'landscape';
  pageCount: number;
  pageGap: number;
  pageHeight: number;
  pageSize: PdfPageSize;
  pageWidth: number;
}

const PDF_PAGE_DIMENSIONS: Record<
  PdfPageSize,
  { width: number; height: number }
> = {
  a3: { width: 841.89, height: 1190.55 },
  a4: { width: 595.28, height: 841.89 },
  a5: { width: 419.53, height: 595.28 },
  letter: { width: 612, height: 792 },
  legal: { width: 612, height: 1008 },
  tabloid: { width: 792, height: 1224 },
};
const MAX_LIVE_DOCUMENT_CAPTURE_HEIGHT = 6_000;

export async function exportWorkArtifactPdf(
  artifact: WorkArtifact,
  options: WorkPdfExportOptions = {},
): Promise<void> {
  const liveDocument = workLiveDocumentPdfSurfaceForExport(artifact.id);
  const allPages = workPdfPagesForExport(artifact.id);
  if (!liveDocument && !allPages.length) {
    throw new Error(
      'PDF print layout is not ready. Please retry after the editor finishes loading.',
    );
  }
  const livePageIndexes = liveDocument
    ? selectedPdfPageIndexes(liveDocument.pageCount, options)
    : [];
  const pages = liveDocument ? [] : workPdfPagesForExport(artifact.id, options);
  if (!livePageIndexes.length && !pages.length)
    throw new Error('Select at least one page before exporting PDF.');

  await document.fonts?.ready;
  await nextPaint();
  const [{ default: html2canvas }, { jsPDF }] = await Promise.all([
    import('html2canvas'),
    import('jspdf'),
  ]);
  let pdf: JsPdf | null = null;

  if (liveDocument) {
    const capture = mountLiveDocumentPdfCapture(liveDocument);
    try {
      const batches = workLiveDocumentPdfCaptureBatches(
        livePageIndexes,
        liveDocument,
      );
      for (const batch of batches) {
        capture.snapshot.style.top = `-${
          batch.firstPageIndex *
          (liveDocument.pageHeight + liveDocument.pageGap)
        }px`;
        capture.viewport.style.height = `${batch.captureHeight}px`;
        await nextPaint();
        const batchCanvas = await html2canvas(capture.viewport, {
          backgroundColor: capture.backgroundColor,
          height: batch.captureHeight,
          logging: false,
          scale: 2,
          useCORS: true,
          width: liveDocument.pageWidth,
          windowHeight: Math.ceil(batch.captureHeight),
          windowWidth: Math.ceil(liveDocument.pageWidth),
        });
        const scaleY = batchCanvas.height / batch.captureHeight;
        for (const pageIndex of batch.pageIndexes) {
          const pageOffset =
            (pageIndex - batch.firstPageIndex) *
            (liveDocument.pageHeight + liveDocument.pageGap);
          const sourceTop = Math.round(pageOffset * scaleY);
          const sourceBottom = Math.round(
            (pageOffset + liveDocument.pageHeight) * scaleY,
          );
          const pageCanvas = cropDocumentCanvas(
            batchCanvas,
            sourceTop,
            Math.max(1, sourceBottom - sourceTop),
            capture.backgroundColor,
          );
          pdf = appendExactCanvasPage(
            pdf,
            pageCanvas,
            liveDocument.orientation,
            liveDocument.pageSize,
            jsPDF,
          );
        }
      }
    } finally {
      capture.host.remove();
    }
  } else {
    for (const page of pages) {
      const orientation =
        page.dataset.pdfOrientation === 'portrait' ? 'portrait' : 'landscape';
      const pageSize = pdfPageSize(page.dataset.pdfPageSize);
      const backgroundColor =
        getComputedStyle(page).backgroundColor || '#ffffff';
      const canvas = await html2canvas(page, {
        backgroundColor,
        logging: false,
        scale: 2,
        useCORS: true,
        windowWidth: Math.max(page.scrollWidth, page.clientWidth),
        windowHeight: Math.max(page.scrollHeight, page.clientHeight),
      });
      pdf = appendCanvas(
        pdf,
        canvas,
        orientation,
        pageSize,
        backgroundColor,
        jsPDF,
      );
    }
  }
  if (!pdf) throw new Error('PDF export did not produce any pages.');
  pdf.setProperties({
    title: artifact.title,
    author: 'A3S Work',
    creator: 'A3S Work',
  });
  pdf.save(`${safeFileName(artifact.title)}.pdf`);
}

export function workLiveDocumentPdfSurfaceForExport(
  artifactId: string,
): WorkLiveDocumentPdfSurface | null {
  const surface = pdfSurfaceForArtifact('live', artifactId);
  const element = surface?.querySelector<HTMLElement>(
    '[data-work-pdf-live-document]',
  );
  const editor = element?.querySelector<HTMLElement>(
    '.work-document-editable .ProseMirror',
  );
  if (!element || editor?.dataset.paginationState !== 'ready') return null;

  const pageCount = finiteDatasetNumber(element.dataset.pdfPageCount, 1);
  const pageGap = finiteDatasetNumber(element.dataset.pdfPageGap, 0);
  const pageHeight = finiteDatasetNumber(element.dataset.pdfPageHeight, 1);
  const pageWidth = finiteDatasetNumber(element.dataset.pdfPageWidth, 1);
  if (
    pageCount === null ||
    pageGap === null ||
    pageHeight === null ||
    pageWidth === null ||
    !Number.isSafeInteger(pageCount) ||
    pageCount > 10_000 ||
    Number(editor.dataset.paginationPages) !== pageCount
  ) {
    return null;
  }

  return {
    element,
    orientation:
      element.dataset.pdfOrientation === 'landscape' ? 'landscape' : 'portrait',
    pageCount,
    pageGap,
    pageHeight,
    pageSize: pdfPageSize(element.dataset.pdfPageSize),
    pageWidth,
  };
}

export function workPdfPagesForExport(
  artifactId: string,
  options: WorkPdfExportOptions = {},
): HTMLElement[] {
  const surface = pdfSurfaceForArtifact('export', artifactId);
  const pages = surface
    ? Array.from(surface.querySelectorAll<HTMLElement>('[data-work-pdf-page]'))
    : [];
  if (options.pageIndexes === undefined) return pages;
  const selected = new Set(
    options.pageIndexes.filter(
      (index) =>
        Number.isSafeInteger(index) && index >= 0 && index < pages.length,
    ),
  );
  return pages.filter((_, index) => selected.has(index));
}

function selectedPdfPageIndexes(
  pageCount: number,
  options: WorkPdfExportOptions,
): number[] {
  if (options.pageIndexes === undefined) {
    return Array.from({ length: pageCount }, (_, index) => index);
  }
  const selected = new Set(
    options.pageIndexes.filter(
      (index) => Number.isSafeInteger(index) && index >= 0 && index < pageCount,
    ),
  );
  return Array.from({ length: pageCount }, (_, index) => index).filter(
    (index) => selected.has(index),
  );
}

export function workLiveDocumentPdfCaptureBatches(
  pageIndexes: readonly number[],
  page: Pick<
    WorkLiveDocumentPdfSurface,
    'pageGap' | 'pageHeight' | 'pageCount'
  >,
): Array<{
  captureHeight: number;
  firstPageIndex: number;
  pageIndexes: number[];
}> {
  const batches: Array<{
    captureHeight: number;
    firstPageIndex: number;
    pageIndexes: number[];
  }> = [];
  for (const pageIndex of pageIndexes) {
    if (
      !Number.isSafeInteger(pageIndex) ||
      pageIndex < 0 ||
      pageIndex >= page.pageCount
    ) {
      continue;
    }
    const current = batches.at(-1);
    const consecutive = current && current.pageIndexes.at(-1) === pageIndex - 1;
    const candidateCount = consecutive ? current.pageIndexes.length + 1 : 1;
    const candidateHeight =
      page.pageHeight * candidateCount + page.pageGap * (candidateCount - 1);
    if (
      !current ||
      !consecutive ||
      (current.pageIndexes.length > 0 &&
        candidateHeight > MAX_LIVE_DOCUMENT_CAPTURE_HEIGHT)
    ) {
      batches.push({
        captureHeight: page.pageHeight,
        firstPageIndex: pageIndex,
        pageIndexes: [pageIndex],
      });
      continue;
    }
    current.captureHeight = candidateHeight;
    current.pageIndexes.push(pageIndex);
  }
  return batches;
}

function mountLiveDocumentPdfCapture(surface: WorkLiveDocumentPdfSurface): {
  backgroundColor: string;
  host: HTMLDivElement;
  snapshot: HTMLElement;
  viewport: HTMLDivElement;
} {
  const host = document.createElement('div');
  host.className = 'work-pdf-export-surface';
  host.setAttribute('aria-hidden', 'true');

  const viewport = document.createElement('div');
  viewport.className = 'work-document-live-pdf-viewport';
  const backgroundColor =
    getComputedStyle(surface.element).backgroundColor || '#ffffff';
  Object.assign(viewport.style, {
    backgroundColor,
    height: `${surface.pageHeight}px`,
    overflow: 'hidden',
    position: 'relative',
    width: `${surface.pageWidth}px`,
  });

  const snapshot = surface.element.cloneNode(true) as HTMLElement;
  snapshot.removeAttribute('data-work-pdf-live-document');
  snapshot.removeAttribute('aria-label');
  snapshot.classList.remove(
    'page-chrome-editing',
    'work-document-preview-page',
  );
  snapshot.classList.add(
    'document',
    'work-document-live-pdf-snapshot',
    'work-pdf-export-page',
  );
  snapshot.dataset.documentCommentAppearance = 'plain';
  snapshot.setAttribute('aria-hidden', 'true');
  const surfaceHeight =
    surface.pageCount * surface.pageHeight +
    (surface.pageCount - 1) * surface.pageGap;
  Object.assign(snapshot.style, {
    borderColor: 'transparent',
    boxShadow: 'none',
    height: `${surfaceHeight}px`,
    left: '0',
    margin: '0',
    minHeight: `${surfaceHeight}px`,
    position: 'absolute',
    top: '0',
    width: `${surface.pageWidth}px`,
  });
  normalizeLiveDocumentPdfSnapshot(snapshot);

  viewport.append(snapshot);
  host.append(viewport);
  document.body.append(host);
  return { backgroundColor, host, snapshot, viewport };
}

function normalizeLiveDocumentPdfSnapshot(snapshot: HTMLElement): void {
  snapshot
    .querySelectorAll('.work-document-selection-toolbar')
    .forEach((element) => {
      element.remove();
    });
  for (const editor of snapshot.querySelectorAll<HTMLElement>(
    '.work-document-page-chrome-inline-editor',
  )) {
    const content = editor.querySelector<HTMLElement>(
      '[data-document-page-chrome-engine="tiptap"]',
    );
    const replacement = document.createElement('div');
    replacement.className = 'work-document-page-chrome-html';
    replacement.innerHTML = content?.innerHTML ?? '';
    editor.replaceWith(replacement);
  }
  for (const element of snapshot.querySelectorAll<HTMLElement>('*')) {
    element.removeAttribute('contenteditable');
    element.classList.remove(
      'ProseMirror-focused',
      'ProseMirror-selectednode',
      'is-editor-empty',
      'selectedCell',
    );
  }
}

function finiteDatasetNumber(
  value: string | undefined,
  minimum: number,
): number | null {
  if (value === undefined || value.trim() === '') return null;
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < minimum || parsed > 1_000_000)
    return null;
  return parsed;
}

function pdfSurfaceForArtifact(
  mode: 'export' | 'live',
  artifactId: string,
): HTMLElement | undefined {
  return Array.from(
    document.querySelectorAll<HTMLElement>(
      `[data-work-pdf-surface="${mode}"][data-work-pdf-artifact]`,
    ),
  ).find((surface) => surface.dataset.workPdfArtifact === artifactId);
}

function appendCanvas(
  pdf: JsPdf | null,
  source: HTMLCanvasElement,
  orientation: 'portrait' | 'landscape',
  pageSize: PdfPageSize,
  backgroundColor: string,
  Pdf: typeof import('jspdf').jsPDF,
): JsPdf {
  let document = pdf;
  const ensurePage = () => {
    if (!document) {
      document = new Pdf({
        orientation,
        unit: 'pt',
        format: pageSize,
        compress: true,
      });
    } else {
      document.addPage(pageSize, orientation);
    }
  };

  const dimensions = PDF_PAGE_DIMENSIONS[pageSize];
  const standardWidth =
    orientation === 'portrait' ? dimensions.width : dimensions.height;
  const standardHeight =
    orientation === 'portrait' ? dimensions.height : dimensions.width;
  const sliceHeight = Math.max(
    1,
    Math.floor((source.width * standardHeight) / standardWidth),
  );
  for (let offset = 0; offset < source.height; offset += sliceHeight) {
    ensurePage();
    const height = Math.min(sliceHeight, source.height - offset);
    const slice = documentCanvas(source.width, height);
    const context = slice.getContext('2d');
    if (!context) throw new Error('The browser could not prepare a PDF page.');
    context.fillStyle = backgroundColor;
    context.fillRect(0, 0, slice.width, slice.height);
    context.drawImage(
      source,
      0,
      offset,
      source.width,
      height,
      0,
      0,
      source.width,
      height,
    );
    const renderedHeight = (standardWidth * height) / source.width;
    document?.addImage(
      slice.toDataURL('image/jpeg', 0.92),
      'JPEG',
      0,
      0,
      standardWidth,
      renderedHeight,
      undefined,
      'FAST',
    );
  }
  return document as JsPdf;
}

function appendExactCanvasPage(
  pdf: JsPdf | null,
  source: HTMLCanvasElement,
  orientation: 'portrait' | 'landscape',
  pageSize: PdfPageSize,
  Pdf: typeof import('jspdf').jsPDF,
): JsPdf {
  const document =
    pdf ??
    new Pdf({
      orientation,
      unit: 'pt',
      format: pageSize,
      compress: true,
    });
  if (pdf) document.addPage(pageSize, orientation);

  const dimensions = PDF_PAGE_DIMENSIONS[pageSize];
  const width =
    orientation === 'portrait' ? dimensions.width : dimensions.height;
  const height =
    orientation === 'portrait' ? dimensions.height : dimensions.width;
  document.addImage(
    source.toDataURL('image/jpeg', 0.92),
    'JPEG',
    0,
    0,
    width,
    height,
    undefined,
    'FAST',
  );
  return document;
}

function pdfPageSize(value: string | undefined): PdfPageSize {
  return value && Object.hasOwn(PDF_PAGE_DIMENSIONS, value)
    ? (value as PdfPageSize)
    : 'a4';
}

function documentCanvas(width: number, height: number): HTMLCanvasElement {
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  return canvas;
}

function cropDocumentCanvas(
  source: HTMLCanvasElement,
  top: number,
  height: number,
  backgroundColor: string,
): HTMLCanvasElement {
  const canvas = documentCanvas(source.width, height);
  const context = canvas.getContext('2d');
  if (!context) throw new Error('The browser could not prepare a PDF page.');
  context.fillStyle = backgroundColor;
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.drawImage(
    source,
    0,
    top,
    source.width,
    height,
    0,
    0,
    source.width,
    height,
  );
  return canvas;
}

function nextPaint(): Promise<void> {
  return new Promise((resolve) =>
    requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
  );
}

function safeFileName(value: string): string {
  return (
    value.replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').trim() || 'A3S Work file'
  );
}
