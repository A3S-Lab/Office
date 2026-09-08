import type { jsPDF as JsPdf } from 'jspdf';
import {
  workPdfCjkFontRegistered,
  workPdfTextNeedsCjkFont,
} from './work-pdf-cjk-font';
import {
  workPdfFontStyleFromCss,
  workPdfTextColorFromCss,
  type WorkPdfStyledTextRun,
} from './work-pdf-vector-text';

/** One extractable PDF text run in page-local CSS pixels (origin: page top-left). */
export interface WorkPdfTextRun {
  fontSize: number;
  height: number;
  text: string;
  width: number;
  x: number;
  y: number;
}

export interface WorkPdfPageBounds {
  height: number;
  left: number;
  top: number;
  width: number;
}

export interface WorkPdfPagePoints {
  pageHeightPoints: number;
  pageWidthPoints: number;
}

const MAX_TEXT_RUNS = 20_000;
const MAX_RUN_TEXT_LENGTH = 4_096;
/** jsPDF's built-in fonts only round-trip Basic Latin / Latin-1 reliably. */
const LATIN_TEXT_PATTERN = /^[\u0020-\u007e\u00a0-\u00ff]+$/;
/**
 * When a host CJK TrueType face is registered, admit common CJK + Latin runs.
 */
const CJK_OR_LATIN_TEXT_PATTERN =
  /^[\u0020-\u007e\u00a0-\u00ff\u3000-\u303f\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uff00-\uffef]+$/;

/**
 * Maps measured client rectangles into page-local PDF text runs. Rectangles
 * outside the page, empty text, or glyphs the active PDF font set cannot encode
 * are omitted (fail soft) rather than inventing geometry.
 */
export function workPdfTextRunsFromClientRects(
  text: string,
  rects: ReadonlyArray<{
    height: number;
    left: number;
    top: number;
    width: number;
  }>,
  page: WorkPdfPageBounds,
  fontSize: number,
): WorkPdfTextRun[] {
  const trimmed = text.replace(/\u00a0/g, ' ');
  const searchable = trimmed.replace(/\s+/g, ' ').trim();
  const pattern = workPdfCjkFontRegistered()
    ? CJK_OR_LATIN_TEXT_PATTERN
    : LATIN_TEXT_PATTERN;
  if (
    !searchable ||
    !pattern.test(searchable) ||
    searchable.length > MAX_RUN_TEXT_LENGTH ||
    !Number.isFinite(fontSize) ||
    fontSize <= 0 ||
    fontSize > 512
  ) {
    return [];
  }
  if (
    workPdfTextNeedsCjkFont(searchable) &&
    !workPdfCjkFontRegistered()
  ) {
    return [];
  }
  const runs: WorkPdfTextRun[] = [];
  for (const rect of rects) {
    if (runs.length >= MAX_TEXT_RUNS) break;
    if (
      !Number.isFinite(rect.left) ||
      !Number.isFinite(rect.top) ||
      !Number.isFinite(rect.width) ||
      !Number.isFinite(rect.height) ||
      rect.width <= 0 ||
      rect.height <= 0
    ) {
      continue;
    }
    const x = rect.left - page.left;
    const y = rect.top - page.top;
    if (
      x + rect.width <= 0 ||
      y + rect.height <= 0 ||
      x >= page.width ||
      y >= page.height
    ) {
      continue;
    }
    runs.push({
      fontSize,
      height: rect.height,
      text: searchable,
      width: rect.width,
      x: Math.max(0, x),
      y: Math.max(0, y),
    });
  }
  return runs;
}

/**
 * Collects searchable Latin text from a live capture root that has already been
 * positioned for one physical page. Missing layout geometry yields no runs.
 * Color and font style travel with each run for vector paint.
 */
export function collectWorkPdfTextRuns(
  root: HTMLElement,
  page: WorkPdfPageBounds,
): WorkPdfStyledTextRun[] {
  if (
    !Number.isFinite(page.width) ||
    !Number.isFinite(page.height) ||
    page.width <= 0 ||
    page.height <= 0
  ) {
    return [];
  }
  const runs: WorkPdfStyledTextRun[] = [];
  const document = root.ownerDocument;
  if (!document) return runs;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) {
    if (runs.length >= MAX_TEXT_RUNS) break;
    const node = walker.currentNode as Text;
    const raw = node.data;
    if (!raw || !raw.trim()) continue;
    const parent = node.parentElement;
    if (!parent || shouldSkipPdfTextElement(parent)) continue;
    const range = document.createRange();
    range.selectNodeContents(node);
    const clientRects = Array.from(range.getClientRects()).map((rect) => ({
      height: rect.height,
      left: rect.left,
      top: rect.top,
      width: rect.width,
    }));
    range.detach?.();
    const computed = getComputedStyle(parent);
    const fontSize = parseFloat(computed.fontSize);
    const baseRuns = workPdfTextRunsFromClientRects(
      raw,
      clientRects,
      page,
      Number.isFinite(fontSize) ? fontSize : 12,
    );
    const color = workPdfTextColorFromCss(computed.color);
    const fontStyle = workPdfFontStyleFromCss(computed);
    for (const run of baseRuns) {
      if (runs.length >= MAX_TEXT_RUNS) break;
      runs.push({ ...run, color, fontStyle });
    }
  }
  return runs.slice(0, MAX_TEXT_RUNS);
}

/**
 * Invisible Helvetica overlay for search/select. Prefer
 * {@link appendWorkPdfVectorTextLayer} after clearing raster text regions for
 * production export; this helper remains for encoding regression tests.
 */
export function appendWorkPdfInvisibleTextLayer(
  pdf: JsPdf,
  runs: readonly WorkPdfTextRun[],
  pageCss: Pick<WorkPdfPageBounds, 'height' | 'width'>,
  pagePoints: WorkPdfPagePoints,
): void {
  if (!runs.length) return;
  if (
    !Number.isFinite(pageCss.width) ||
    !Number.isFinite(pageCss.height) ||
    pageCss.width <= 0 ||
    pageCss.height <= 0 ||
    !Number.isFinite(pagePoints.pageWidthPoints) ||
    !Number.isFinite(pagePoints.pageHeightPoints) ||
    pagePoints.pageWidthPoints <= 0 ||
    pagePoints.pageHeightPoints <= 0
  ) {
    return;
  }
  const scaleX = pagePoints.pageWidthPoints / pageCss.width;
  const scaleY = pagePoints.pageHeightPoints / pageCss.height;
  const GState = (
    pdf as JsPdf & {
      GState?: new (options: { opacity: number }) => object;
    }
  ).GState;
  pdf.saveGraphicsState?.();
  if (GState) {
    pdf.setGState?.(new GState({ opacity: 0 }));
  } else {
    pdf.setTextColor(255, 255, 255);
  }
  pdf.setFont('helvetica', 'normal');
  for (const run of runs) {
    const fontSizePt = Math.min(
      72,
      Math.max(4, run.fontSize * ((scaleX + scaleY) / 2) * 0.75),
    );
    pdf.setFontSize(fontSizePt);
    const x = run.x * scaleX;
    const y = run.y * scaleY + fontSizePt * 0.85;
    if (
      !Number.isFinite(x) ||
      !Number.isFinite(y) ||
      x < -1 ||
      y < -1 ||
      x > pagePoints.pageWidthPoints + 1 ||
      y > pagePoints.pageHeightPoints + 1
    ) {
      continue;
    }
    try {
      pdf.text(run.text, x, y, { baseline: 'alphabetic' });
    } catch {
      // Skip runs the built-in font cannot encode.
    }
  }
  pdf.restoreGraphicsState?.();
}

function shouldSkipPdfTextElement(element: HTMLElement): boolean {
  if (element.closest('[aria-hidden="true"]')) {
    // The capture host itself is aria-hidden; skip only nested chrome.
    const hidden = element.closest('[aria-hidden="true"]');
    if (
      hidden &&
      !hidden.classList.contains('work-document-live-pdf-snapshot') &&
      !hidden.classList.contains('work-pdf-export-page')
    ) {
      return true;
    }
  }
  const tag = element.tagName.toLowerCase();
  return tag === 'script' || tag === 'style' || tag === 'noscript';
}
