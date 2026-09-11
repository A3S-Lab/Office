import type { jsPDF as JsPdf } from 'jspdf';
import {
  ensureWorkPdfCjkFontOnDocument,
  workPdfTextNeedsCjkFont,
} from './work-pdf-cjk-font';
import {
  beginWorkPdfActualTextSpan,
  endWorkPdfActualTextSpan,
} from './work-pdf-structure';
import type {
  WorkPdfPageBounds,
  WorkPdfPagePoints,
  WorkPdfTextRun,
} from './work-pdf-text-layer';

export type WorkPdfTextFontStyle = 'normal' | 'bold' | 'italic' | 'bolditalic';

/** Bounded PDF underline stroke kinds (CSS/Office styles collapse here). */
export type WorkPdfUnderlineKind = 'single' | 'double' | 'thick';

export interface WorkPdfRunUnderline {
  color: string;
  kind: WorkPdfUnderlineKind;
}

/** Opaque fill color for a Writer/CSS highlight background strip. */
export interface WorkPdfRunHighlight {
  color: string;
}

export interface WorkPdfStyledTextRun extends WorkPdfTextRun {
  color: string;
  fontStyle: WorkPdfTextFontStyle;
  /** When set, PDF export paints a vector highlight fill under the text run. */
  highlight?: WorkPdfRunHighlight;
  /** When set, PDF export paints a vector underline after clearing raster text. */
  underline?: WorkPdfRunUnderline;
}

/**
 * Clears raster glyphs under measured Latin runs so the vector text layer can
 * paint sharp, searchable text without inventing a second layout model.
 */
export function clearWorkPdfTextRunsOnCanvas(
  canvas: HTMLCanvasElement,
  runs: readonly WorkPdfTextRun[],
  pageCss: Pick<WorkPdfPageBounds, 'height' | 'width'>,
  fillStyle: string,
): void {
  if (!runs.length) return;
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
  for (const run of runs) {
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
    const padX = Math.max(1, scaleX * 0.5);
    const padY = Math.max(1, scaleY * 0.35);
    context.fillRect(
      run.x * scaleX - padX,
      run.y * scaleY - padY,
      run.width * scaleX + padX * 2,
      run.height * scaleY + padY * 2,
    );
  }
  context.restore();
}

/**
 * Paints Latin/Latin-1 text as native PDF text operators using Helvetica, and
 * CJK runs with a host-registered TrueType face when available. Non-encodable
 * runs are skipped. Geometry matches the searchable text collector.
 */
export function appendWorkPdfVectorTextLayer(
  pdf: JsPdf,
  runs: readonly WorkPdfStyledTextRun[],
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
  const cjkFamily = runs.some((run) => workPdfTextNeedsCjkFont(run.text))
    ? ensureWorkPdfCjkFontOnDocument(pdf)
    : null;
  for (const run of runs) {
    const fontSizePt = Math.min(
      72,
      Math.max(4, run.fontSize * ((scaleX + scaleY) / 2) * 0.75),
    );
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
    const needsCjk = workPdfTextNeedsCjkFont(run.text);
    if (needsCjk && !cjkFamily) continue;
    const rgb = parseCssRgb(run.color);
    if (needsCjk && cjkFamily) {
      pdf.setFont(cjkFamily, 'normal');
    } else {
      pdf.setFont('helvetica', pdfFontStyle(run.fontStyle));
    }
    pdf.setFontSize(fontSizePt);
    if (rgb) pdf.setTextColor(rgb[0], rgb[1], rgb[2]);
    else pdf.setTextColor(0, 0, 0);
    const marked = beginWorkPdfActualTextSpan(pdf, run.text);
    try {
      pdf.text(run.text, x, y, { baseline: 'alphabetic' });
    } catch {
      // Skip runs the active font cannot encode.
    } finally {
      if (marked) endWorkPdfActualTextSpan(pdf);
    }
  }
}

export function workPdfFontStyleFromCss(
  style: Pick<CSSStyleDeclaration, 'fontStyle' | 'fontWeight'>,
): WorkPdfTextFontStyle {
  const italic = style.fontStyle === 'italic' || style.fontStyle === 'oblique';
  const weight = Number.parseInt(style.fontWeight, 10);
  const bold =
    style.fontWeight === 'bold' ||
    style.fontWeight === 'bolder' ||
    (Number.isFinite(weight) && weight >= 600);
  if (bold && italic) return 'bolditalic';
  if (bold) return 'bold';
  if (italic) return 'italic';
  return 'normal';
}

export function workPdfTextColorFromCss(color: string): string {
  const rgb = parseCssRgb(color);
  if (!rgb) return '#000000';
  return `#${rgb.map((channel) => channel.toString(16).padStart(2, '0')).join('')}`;
}

function pdfFontStyle(
  style: WorkPdfTextFontStyle,
): 'normal' | 'bold' | 'italic' | 'bolditalic' {
  return style;
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
