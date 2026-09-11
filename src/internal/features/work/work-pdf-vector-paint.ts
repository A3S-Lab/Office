import type { jsPDF as JsPdf } from 'jspdf';
import {
  DOCUMENT_HIGHLIGHT_ATTRIBUTE,
  documentHighlightCssColor,
  documentHighlightForCssColor,
  normalizeDocumentHighlight,
} from './work-document-highlight';
import {
  DOCUMENT_UNDERLINE_STYLE_ATTRIBUTE,
  documentUnderlineFormattingFromElement,
  type WorkDocumentUnderlineStyle,
} from './work-document-underline';
import {
  workPdfTextColorFromCss,
  type WorkPdfRunHighlight,
  type WorkPdfRunUnderline,
  type WorkPdfStyledTextRun,
  type WorkPdfUnderlineKind,
} from './work-pdf-vector-text';

export type { WorkPdfRunHighlight, WorkPdfRunUnderline, WorkPdfUnderlineKind };

type WorkPdfPaintPageCss = { height: number; width: number };
type WorkPdfPaintPagePoints = {
  pageHeightPoints: number;
  pageWidthPoints: number;
};

/**
 * Resolves a Writer/CSS highlight on a text host into an opaque PDF fill color.
 * Only the portable highlight palette is admitted; `none` and transparent skip.
 */
export function workPdfRunHighlightFromElement(
  element: HTMLElement,
): WorkPdfRunHighlight | null {
  let host: HTMLElement | null = element;
  for (
    let depth = 0;
    host && depth < 6;
    depth += 1, host = host.parentElement
  ) {
    const attr = normalizeDocumentHighlight(
      host.getAttribute(DOCUMENT_HIGHLIGHT_ATTRIBUTE),
    );
    if (attr === 'none') return null;
    if (attr) {
      const color = documentHighlightCssColor(attr);
      if (color && color !== 'transparent') return { color };
    }
  }
  const fromCss = documentHighlightForCssColor(
    getComputedStyle(element).backgroundColor,
  );
  if (!fromCss || fromCss === 'none') return null;
  const color = documentHighlightCssColor(fromCss);
  return color && color !== 'transparent' ? { color } : null;
}

/**
 * Resolves a Writer/CSS underline on a text host into a PDF stroke plan.
 * Unknown decorative styles map to a single solid line (bounded approximation).
 */
export function workPdfRunUnderlineFromElement(
  element: HTMLElement,
): WorkPdfRunUnderline | null {
  let host: HTMLElement | null = element;
  for (
    let depth = 0;
    host && depth < 6;
    depth += 1, host = host.parentElement
  ) {
    const formatting = documentUnderlineFormattingFromElement(host);
    if (formatting) {
      if (formatting.style === 'none') return null;
      return {
        color:
          formatting.color ??
          workPdfTextColorFromCss(getComputedStyle(host).color),
        kind: workPdfUnderlineKindFromStyle(formatting.style),
      };
    }
    if (host.getAttribute(DOCUMENT_UNDERLINE_STYLE_ATTRIBUTE) === 'none') {
      return null;
    }
  }
  const computed = getComputedStyle(element);
  const lines = (computed.textDecorationLine || 'none')
    .toLowerCase()
    .split(/\s+/);
  if (lines.includes('none') || !lines.includes('underline')) return null;
  const cssStyle = (computed.textDecorationStyle || 'solid')
    .trim()
    .toLowerCase();
  const decorationColor = (computed.textDecorationColor || '')
    .trim()
    .toLowerCase();
  return {
    color: workPdfTextColorFromCss(
      decorationColor &&
        decorationColor !== 'currentcolor' &&
        decorationColor !== ''
        ? computed.textDecorationColor
        : computed.color || '#000000',
    ),
    kind: cssStyle === 'double' ? 'double' : thicknessHintFromCss(computed),
  };
}

/**
 * Paints highlight background fills as native PDF path operators using the same
 * page geometry as the vector text layer (under text, above the raster page).
 */
export function appendWorkPdfVectorHighlightLayer(
  pdf: JsPdf,
  runs: readonly WorkPdfStyledTextRun[],
  pageCss: WorkPdfPaintPageCss,
  pagePoints: WorkPdfPaintPagePoints,
): void {
  const highlighted = runs.filter((run) => run.highlight);
  if (!highlighted.length) return;
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
  for (const run of highlighted) {
    const highlight = run.highlight;
    if (!highlight) continue;
    const x = run.x * scaleX;
    const y = run.y * scaleY;
    const width = run.width * scaleX;
    const height = run.height * scaleY;
    if (
      !Number.isFinite(x) ||
      !Number.isFinite(y) ||
      !Number.isFinite(width) ||
      !Number.isFinite(height) ||
      width <= 0 ||
      height <= 0 ||
      x < -1 ||
      y < -1 ||
      x + width > pagePoints.pageWidthPoints + 1 ||
      y + height > pagePoints.pageHeightPoints + 1
    ) {
      continue;
    }
    const rgb = parseCssRgb(highlight.color);
    if (rgb) pdf.setFillColor(rgb[0], rgb[1], rgb[2]);
    else pdf.setFillColor(255, 255, 0);
    pdf.rect(x, y, width, height, 'F');
  }
}

/**
 * Clears a thin strip under underlined vector runs so residual CSS decoration
 * does not double-paint with the PDF stroke layer.
 */
export function clearWorkPdfUnderlineStripsOnCanvas(
  canvas: HTMLCanvasElement,
  runs: readonly WorkPdfStyledTextRun[],
  pageCss: WorkPdfPaintPageCss,
  fillStyle: string,
): void {
  const underlined = runs.filter((run) => run.underline);
  if (!underlined.length) return;
  if (
    !Number.isFinite(pageCss.width) ||
    !Number.isFinite(pageCss.height) ||
    pageCss.width <= 0 ||
    pageCss.height <= 0 ||
    canvas.width <= 0 ||
    canvas.height <= 0
  ) {
    return;
  }
  const context = canvas.getContext('2d');
  if (!context) return;
  const scaleX = canvas.width / pageCss.width;
  const scaleY = canvas.height / pageCss.height;
  context.save();
  context.fillStyle = fillStyle || '#ffffff';
  for (const run of underlined) {
    if (
      !Number.isFinite(run.x) ||
      !Number.isFinite(run.y) ||
      !Number.isFinite(run.width) ||
      !Number.isFinite(run.height) ||
      run.width <= 0 ||
      run.height <= 0
    ) {
      continue;
    }
    const stripHeight = Math.max(run.height * 0.35, 2);
    const padX = Math.max(1, scaleX * 0.5);
    context.fillRect(
      run.x * scaleX - padX,
      (run.y + run.height * 0.75) * scaleY,
      run.width * scaleX + padX * 2,
      stripHeight * scaleY,
    );
  }
  context.restore();
}

/**
 * Paints underline strokes as native PDF path operators using the same page
 * geometry as the vector text layer (no second layout model).
 */
export function appendWorkPdfVectorUnderlineLayer(
  pdf: JsPdf,
  runs: readonly WorkPdfStyledTextRun[],
  pageCss: WorkPdfPaintPageCss,
  pagePoints: WorkPdfPaintPagePoints,
): void {
  const underlined = runs.filter((run) => run.underline);
  if (!underlined.length) return;
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
  for (const run of underlined) {
    const underline = run.underline;
    if (!underline) continue;
    const x1 = run.x * scaleX;
    const x2 = (run.x + run.width) * scaleX;
    const y = (run.y + run.height * 0.92) * scaleY;
    if (
      !Number.isFinite(x1) ||
      !Number.isFinite(x2) ||
      !Number.isFinite(y) ||
      x2 <= x1 ||
      x1 < -1 ||
      x2 > pagePoints.pageWidthPoints + 1 ||
      y < -1 ||
      y > pagePoints.pageHeightPoints + 1
    ) {
      continue;
    }
    const rgb = parseCssRgb(underline.color);
    if (rgb) pdf.setDrawColor(rgb[0], rgb[1], rgb[2]);
    else pdf.setDrawColor(0, 0, 0);
    const thickness =
      underline.kind === 'thick'
        ? Math.max(1.2, run.fontSize * ((scaleX + scaleY) / 2) * 0.08)
        : Math.max(0.6, run.fontSize * ((scaleX + scaleY) / 2) * 0.045);
    pdf.setLineWidth(thickness);
    pdf.line(x1, y, x2, y);
    if (underline.kind === 'double') {
      const gap = Math.max(1.2, thickness * 1.75);
      pdf.setLineWidth(Math.max(0.5, thickness * 0.85));
      pdf.line(x1, y + gap, x2, y + gap);
    }
  }
}

function workPdfUnderlineKindFromStyle(
  style: WorkDocumentUnderlineStyle,
): WorkPdfUnderlineKind {
  if (style === 'double' || style === 'wavyDouble') return 'double';
  if (
    style === 'thick' ||
    style === 'dottedHeavy' ||
    style === 'dashedHeavy' ||
    style === 'dashLongHeavy' ||
    style === 'dashDotHeavy' ||
    style === 'dashDotDotHeavy' ||
    style === 'wavyHeavy'
  ) {
    return 'thick';
  }
  return 'single';
}

function thicknessHintFromCss(
  style: CSSStyleDeclaration,
): WorkPdfUnderlineKind {
  const raw = (style.textDecorationThickness || '').trim().toLowerCase();
  if (!raw || raw === 'auto' || raw === 'from-font') return 'single';
  const px = Number.parseFloat(raw);
  if (Number.isFinite(px) && px >= 2) return 'thick';
  return 'single';
}

function parseCssRgb(color: string): [number, number, number] | null {
  const trimmed = color.trim().toLowerCase();
  if (!trimmed || trimmed === 'transparent') return null;
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(trimmed);
  if (hex) {
    const value = hex[1];
    if (value.length === 3) {
      return [
        Number.parseInt(value[0] + value[0], 16),
        Number.parseInt(value[1] + value[1], 16),
        Number.parseInt(value[2] + value[2], 16),
      ];
    }
    return [
      Number.parseInt(value.slice(0, 2), 16),
      Number.parseInt(value.slice(2, 4), 16),
      Number.parseInt(value.slice(4, 6), 16),
    ];
  }
  const rgb =
    /^rgba?\(\s*([0-9.]+)\s*,\s*([0-9.]+)\s*,\s*([0-9.]+)(?:\s*,\s*([0-9.]+))?\s*\)$/.exec(
      trimmed,
    );
  if (!rgb) return null;
  const alpha = rgb[4] === undefined ? 1 : Number(rgb[4]);
  if (!Number.isFinite(alpha) || alpha <= 0) return null;
  const channels = [rgb[1], rgb[2], rgb[3]].map((part) => {
    const value = Number(part);
    return Number.isFinite(value)
      ? Math.max(0, Math.min(255, Math.round(value)))
      : 0;
  });
  return [channels[0], channels[1], channels[2]];
}
