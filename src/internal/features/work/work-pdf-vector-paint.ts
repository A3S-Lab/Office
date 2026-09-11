import type { jsPDF as JsPdf } from 'jspdf';
import {
  DOCUMENT_HIGHLIGHT_ATTRIBUTE,
  documentHighlightCssColor,
  documentHighlightForCssColor,
  normalizeDocumentHighlight,
} from './work-document-highlight';
import {
  DOCUMENT_PARAGRAPH_BORDERS_ATTRIBUTE,
  documentBorderPresentation,
  isDocumentParagraphArtBorderStyle,
  parseDocumentParagraphBordersElement,
  type DocumentParagraphBorder,
  type DocumentParagraphBorderEdge,
} from './work-document-paragraph-borders';
import {
  DOCUMENT_UNDERLINE_STYLE_ATTRIBUTE,
  documentUnderlineFormattingFromElement,
  type WorkDocumentUnderlineStyle,
} from './work-document-underline';
import type { WorkPdfPageBounds } from './work-pdf-text-layer';
import {
  workPdfTextColorFromCss,
  type WorkPdfRunHighlight,
  type WorkPdfRunUnderline,
  type WorkPdfStyledTextRun,
  type WorkPdfUnderlineKind,
} from './work-pdf-vector-text';

export type { WorkPdfRunHighlight, WorkPdfRunUnderline, WorkPdfUnderlineKind };

/** Bounded PDF paragraph-border stroke kinds (most art styles skip). */
export type WorkPdfParagraphBorderKind =
  | 'single'
  | 'double'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'wave'
  | 'doubleWave'
  | 'threeDEmboss'
  | 'threeDEngrave'
  | 'inset'
  | 'outset'
  | 'zigZag'
  | 'zigZagStitch'
  | 'sawtooth'
  | 'sharksTeeth'
  | 'triangles'
  | 'triangle1'
  | 'triangle2'
  | 'ovals'
  | 'rings'
  | 'marquee'
  | 'marqueeToothed';

/**
 * PDF stroke edges for Writer paragraph borders. `between` maps to the bottom
 * of the measured box (DOM inset-bottom approximation); `bar` maps to the left
 * (DOM inset-left / facing-page bar approximation).
 */
export type WorkPdfParagraphBorderBoxEdge =
  | 'top'
  | 'left'
  | 'bottom'
  | 'right'
  | 'between'
  | 'bar';

export interface WorkPdfParagraphBorderEdgeStroke {
  color: string;
  kind: WorkPdfParagraphBorderKind;
  /** CSS pixel width from Writer presentation helpers. */
  width: number;
}

export interface WorkPdfParagraphBorderBox {
  edges: Partial<
    Record<WorkPdfParagraphBorderBoxEdge, WorkPdfParagraphBorderEdgeStroke>
  >;
  height: number;
  width: number;
  x: number;
  y: number;
}

type WorkPdfPaintPageCss = { height: number; width: number };
type WorkPdfPaintPagePoints = {
  pageHeightPoints: number;
  pageWidthPoints: number;
};

const MAX_PARAGRAPH_BORDER_BOXES = 2_000;
const PARAGRAPH_BORDER_BOX_EDGES = [
  'top',
  'left',
  'bottom',
  'right',
  'between',
  'bar',
] as const satisfies readonly WorkPdfParagraphBorderBoxEdge[];
const PARAGRAPH_BORDER_SELECTOR = `p, h1, h2, h3, h4, h5, h6, [${DOCUMENT_PARAGRAPH_BORDERS_ATTRIBUTE}]`;

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

/**
 * Resolves Writer paragraph borders on a block into a bounded PDF stroke plan.
 * Admits top/left/bottom/right plus between/bar line styles, including explicit
 * `wave` / `doubleWave` polylines, dual-tone 3D / inset / outset relief strokes,
 * and geometric `zigZag` / `zigZagStitch` / `sawtooth` / `sharksTeeth` /
 * `triangles` / `triangle1` / `triangle2` / `ovals` / `rings` / `marquee` /
 * `marqueeToothed` art motifs; other art border styles are skipped (fail
 * closed). Not PDF/UA.
 */
export function workPdfParagraphBordersFromElement(
  element: HTMLElement,
): Partial<
  Record<WorkPdfParagraphBorderBoxEdge, WorkPdfParagraphBorderEdgeStroke>
> | null {
  const borders = parseDocumentParagraphBordersElement(element);
  if (!borders) return null;
  const edges: Partial<
    Record<WorkPdfParagraphBorderBoxEdge, WorkPdfParagraphBorderEdgeStroke>
  > = {};
  for (const edge of PARAGRAPH_BORDER_BOX_EDGES) {
    const border = borders[edge as DocumentParagraphBorderEdge];
    if (!border) continue;
    const stroke = workPdfParagraphBorderStrokeFromDocumentBorder(border);
    if (stroke) edges[edge] = stroke;
  }
  return Object.keys(edges).length ? edges : null;
}

/**
 * Collects measured paragraph border boxes in page-local CSS pixels. Geometry
 * comes from live DOM rects (same page model as text / outline collectors).
 */
export function collectWorkPdfParagraphBorderBoxes(
  root: HTMLElement,
  page: WorkPdfPageBounds,
): WorkPdfParagraphBorderBox[] {
  if (
    !Number.isFinite(page.width) ||
    !Number.isFinite(page.height) ||
    page.width <= 0 ||
    page.height <= 0
  ) {
    return [];
  }
  const boxes: WorkPdfParagraphBorderBox[] = [];
  const seen = new Set<HTMLElement>();
  for (const element of Array.from(
    root.querySelectorAll<HTMLElement>(PARAGRAPH_BORDER_SELECTOR),
  )) {
    if (boxes.length >= MAX_PARAGRAPH_BORDER_BOXES) break;
    if (seen.has(element)) continue;
    seen.add(element);
    if (shouldSkipPdfBorderElement(element)) continue;
    const edges = workPdfParagraphBordersFromElement(element);
    if (!edges) continue;
    const rect = element.getBoundingClientRect();
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
    boxes.push({
      edges,
      height: rect.height,
      width: rect.width,
      x: Math.max(0, x),
      y: Math.max(0, y),
    });
  }
  return boxes;
}

/**
 * Clears thin CSS border strips on the raster page so vector strokes do not
 * double-paint residual html2canvas borders.
 */
export function clearWorkPdfParagraphBorderStripsOnCanvas(
  canvas: HTMLCanvasElement,
  boxes: readonly WorkPdfParagraphBorderBox[],
  pageCss: WorkPdfPaintPageCss,
  fillStyle: string,
): void {
  if (!boxes.length) return;
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
  for (const box of boxes) {
    for (const edge of PARAGRAPH_BORDER_BOX_EDGES) {
      const stroke = box.edges[edge];
      if (!stroke) continue;
      const strip = Math.max(stroke.width, 1) + 1;
      const pad = Math.max(1, scaleX * 0.5);
      const paintEdge = paragraphBorderPaintEdge(edge);
      if (paintEdge === 'top') {
        context.fillRect(
          box.x * scaleX - pad,
          box.y * scaleY - pad,
          box.width * scaleX + pad * 2,
          strip * scaleY,
        );
      } else if (paintEdge === 'bottom') {
        context.fillRect(
          box.x * scaleX - pad,
          (box.y + box.height - strip) * scaleY,
          box.width * scaleX + pad * 2,
          strip * scaleY + pad,
        );
      } else if (paintEdge === 'left') {
        context.fillRect(
          box.x * scaleX - pad,
          box.y * scaleY - pad,
          strip * scaleX,
          box.height * scaleY + pad * 2,
        );
      } else {
        context.fillRect(
          (box.x + box.width - strip) * scaleX,
          box.y * scaleY - pad,
          strip * scaleX + pad,
          box.height * scaleY + pad * 2,
        );
      }
    }
  }
  context.restore();
}

/**
 * Paints paragraph borders as native PDF path operators at measured paragraph
 * geometry (common + wave + 3D + zigZag/sawtooth/triangle/oval/marquee art;
 * not PDF/UA or decorative art).
 */
export function appendWorkPdfVectorParagraphBorderLayer(
  pdf: JsPdf,
  boxes: readonly WorkPdfParagraphBorderBox[],
  pageCss: WorkPdfPaintPageCss,
  pagePoints: WorkPdfPaintPagePoints,
): void {
  if (!boxes.length) return;
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
  const scale = (scaleX + scaleY) / 2;
  for (const box of boxes) {
    const x = box.x * scaleX;
    const y = box.y * scaleY;
    const width = box.width * scaleX;
    const height = box.height * scaleY;
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
    for (const edge of PARAGRAPH_BORDER_BOX_EDGES) {
      const stroke = box.edges[edge];
      if (!stroke) continue;
      const rgb = parseCssRgb(stroke.color);
      if (rgb) pdf.setDrawColor(rgb[0], rgb[1], rgb[2]);
      else pdf.setDrawColor(0, 0, 0);
      const thickness =
        stroke.kind === 'thick'
          ? Math.max(1.4, stroke.width * scale)
          : Math.max(0.6, stroke.width * scale);
      applyWorkPdfBorderDash(pdf, stroke.kind, thickness);
      pdf.setLineWidth(thickness);
      if (stroke.kind === 'wave' || stroke.kind === 'doubleWave') {
        strokeWaveParagraphBorderEdge(
          pdf,
          edge,
          x,
          y,
          width,
          height,
          thickness,
          stroke.kind === 'doubleWave',
        );
      } else if (stroke.kind === 'zigZag' || stroke.kind === 'zigZagStitch') {
        strokeZigZagParagraphBorderEdge(
          pdf,
          edge,
          x,
          y,
          width,
          height,
          thickness,
          stroke.kind === 'zigZagStitch',
        );
      } else if (stroke.kind === 'sawtooth' || stroke.kind === 'sharksTeeth') {
        strokeSawtoothParagraphBorderEdge(
          pdf,
          edge,
          x,
          y,
          width,
          height,
          thickness,
          stroke.kind === 'sharksTeeth',
        );
      } else if (
        stroke.kind === 'triangles' ||
        stroke.kind === 'triangle1' ||
        stroke.kind === 'triangle2'
      ) {
        strokeTrianglesParagraphBorderEdge(
          pdf,
          edge,
          x,
          y,
          width,
          height,
          thickness,
          stroke.kind,
        );
      } else if (stroke.kind === 'ovals' || stroke.kind === 'rings') {
        strokeOvalParagraphBorderEdge(
          pdf,
          edge,
          x,
          y,
          width,
          height,
          thickness,
          stroke.kind === 'rings',
        );
      } else if (
        stroke.kind === 'marquee' ||
        stroke.kind === 'marqueeToothed'
      ) {
        strokeMarqueeParagraphBorderEdge(
          pdf,
          edge,
          x,
          y,
          width,
          height,
          thickness,
          stroke.kind === 'marqueeToothed',
        );
      } else if (
        stroke.kind === 'threeDEmboss' ||
        stroke.kind === 'threeDEngrave' ||
        stroke.kind === 'inset' ||
        stroke.kind === 'outset'
      ) {
        strokeThreeDParagraphBorderEdge(
          pdf,
          edge,
          x,
          y,
          width,
          height,
          thickness,
          stroke.kind,
          rgb ?? [0, 0, 0],
        );
      } else {
        strokeParagraphBorderEdge(pdf, edge, x, y, width, height, 0);
        if (stroke.kind === 'double') {
          const gap = Math.max(1.2, thickness * 1.75);
          pdf.setLineWidth(Math.max(0.5, thickness * 0.85));
          strokeParagraphBorderEdge(pdf, edge, x, y, width, height, gap);
        }
      }
      clearWorkPdfBorderDash(pdf);
    }
  }
}

function workPdfParagraphBorderStrokeFromDocumentBorder(
  border: DocumentParagraphBorder,
): WorkPdfParagraphBorderEdgeStroke | null {
  if (border.style === 'nil' || border.style === 'none') {
    return null;
  }
  if (
    border.style === 'zigZag' ||
    border.style === 'zigZagStitch' ||
    border.style === 'sawtooth' ||
    border.style === 'sharksTeeth' ||
    border.style === 'triangles' ||
    border.style === 'triangle1' ||
    border.style === 'triangle2' ||
    border.style === 'ovals' ||
    border.style === 'rings' ||
    border.style === 'marquee' ||
    border.style === 'marqueeToothed'
  ) {
    const presentation = documentBorderPresentation(border);
    if (presentation.width <= 0 || presentation.color === 'transparent') {
      return null;
    }
    return {
      color: presentation.color,
      kind: border.style,
      width: presentation.width,
    };
  }
  if (isDocumentParagraphArtBorderStyle(border.style)) {
    return null;
  }
  const presentation = documentBorderPresentation(border);
  if (presentation.width <= 0 || presentation.color === 'transparent') {
    return null;
  }
  if (border.style === 'wave' || border.style === 'doubleWave') {
    return {
      color: presentation.color,
      kind: border.style,
      width: presentation.width,
    };
  }
  if (
    border.style === 'threeDEmboss' ||
    border.style === 'threeDEngrave' ||
    border.style === 'inset' ||
    border.style === 'outset'
  ) {
    return {
      color: presentation.color,
      kind: border.style,
      width: presentation.width,
    };
  }
  if (presentation.style === 'none') return null;
  const kind = workPdfParagraphBorderKindFromPresentation(
    presentation.style,
    presentation.width,
    border.style,
  );
  if (!kind) return null;
  return {
    color: presentation.color,
    kind,
    width: presentation.width,
  };
}

function workPdfParagraphBorderKindFromPresentation(
  cssStyle:
    | 'solid'
    | 'dashed'
    | 'dotted'
    | 'double'
    | 'inset'
    | 'outset'
    | 'none',
  width: number,
  semantic: DocumentParagraphBorder['style'],
): WorkPdfParagraphBorderKind | null {
  if (cssStyle === 'none' || cssStyle === 'inset' || cssStyle === 'outset') {
    return null;
  }
  if (cssStyle === 'double') return 'double';
  if (cssStyle === 'dashed') return 'dashed';
  if (cssStyle === 'dotted') return 'dotted';
  if (semantic === 'thick' || width >= 2.5) return 'thick';
  return 'single';
}

function paragraphBorderPaintEdge(
  edge: WorkPdfParagraphBorderBoxEdge,
): 'top' | 'left' | 'bottom' | 'right' {
  if (edge === 'between') return 'bottom';
  if (edge === 'bar') return 'left';
  return edge;
}

function strokeParagraphBorderEdge(
  pdf: JsPdf,
  edge: WorkPdfParagraphBorderBoxEdge,
  x: number,
  y: number,
  width: number,
  height: number,
  inset: number,
): void {
  const paintEdge = paragraphBorderPaintEdge(edge);
  if (paintEdge === 'top') {
    pdf.line(x, y + inset, x + width, y + inset);
  } else if (paintEdge === 'bottom') {
    pdf.line(x, y + height - inset, x + width, y + height - inset);
  } else if (paintEdge === 'left') {
    pdf.line(x + inset, y, x + inset, y + height);
  } else {
    pdf.line(x + width - inset, y, x + width - inset, y + height);
  }
}

/**
 * Dual-tone relief strokes for threeDEmboss / threeDEngrave / inset / outset.
 * Not a silent single-line approximation: paints highlight + shadow offsets.
 */
function strokeThreeDParagraphBorderEdge(
  pdf: JsPdf,
  edge: WorkPdfParagraphBorderBoxEdge,
  x: number,
  y: number,
  width: number,
  height: number,
  thickness: number,
  kind: 'threeDEmboss' | 'threeDEngrave' | 'inset' | 'outset',
  baseRgb: [number, number, number],
): void {
  const gap = Math.max(0.8, thickness * 0.95);
  const light = shiftRgbToward(baseRgb, 255, 0.55);
  const dark = shiftRgbToward(baseRgb, 0, 0.55);
  const embossed = kind === 'threeDEmboss' || kind === 'outset';
  const first = embossed ? light : dark;
  const second = embossed ? dark : light;
  pdf.setLineWidth(Math.max(0.5, thickness * 0.85));
  pdf.setDrawColor(first[0], first[1], first[2]);
  strokeParagraphBorderEdge(pdf, edge, x, y, width, height, 0);
  pdf.setDrawColor(second[0], second[1], second[2]);
  strokeParagraphBorderEdge(pdf, edge, x, y, width, height, gap);
}

function shiftRgbToward(
  rgb: [number, number, number],
  target: number,
  amount: number,
): [number, number, number] {
  const t = Math.min(1, Math.max(0, amount));
  return [
    Math.round(rgb[0] + (target - rgb[0]) * t),
    Math.round(rgb[1] + (target - rgb[1]) * t),
    Math.round(rgb[2] + (target - rgb[2]) * t),
  ];
}

/**
 * Strokes an explicit sine polyline for `wave` / `doubleWave` (not a silent
 * straight-line approximation of those OOXML styles).
 */
function strokeWaveParagraphBorderEdge(
  pdf: JsPdf,
  edge: WorkPdfParagraphBorderBoxEdge,
  x: number,
  y: number,
  width: number,
  height: number,
  thickness: number,
  doubleWave: boolean,
): void {
  const paintEdge = paragraphBorderPaintEdge(edge);
  const amplitude = Math.max(1.2, thickness * 1.15);
  const wavelength = Math.max(6, thickness * 8);
  const offsets = doubleWave ? [-amplitude * 0.85, amplitude * 0.85] : [0];
  for (const offset of offsets) {
    if (paintEdge === 'top' || paintEdge === 'bottom') {
      const yBase = (paintEdge === 'top' ? y : y + height) + offset;
      strokeWavePolyline(pdf, x, yBase, width, 0, wavelength, amplitude);
    } else {
      const xBase = (paintEdge === 'left' ? x : x + width) + offset;
      strokeWavePolyline(pdf, xBase, y, height, 1, wavelength, amplitude);
    }
  }
}

/** axis: 0 = horizontal along +x, 1 = vertical along +y. */
function strokeWavePolyline(
  pdf: JsPdf,
  originX: number,
  originY: number,
  length: number,
  axis: 0 | 1,
  wavelength: number,
  amplitude: number,
): void {
  if (
    !Number.isFinite(length) ||
    length <= 0 ||
    !Number.isFinite(wavelength) ||
    wavelength <= 0
  ) {
    return;
  }
  const steps = Math.max(8, Math.ceil((length / wavelength) * 4));
  let prevX = originX;
  let prevY = originY;
  for (let index = 1; index <= steps; index += 1) {
    const t = index / steps;
    const along = length * t;
    const wave = Math.sin((along / wavelength) * Math.PI * 2) * amplitude;
    const nextX = axis === 0 ? originX + along : originX + wave;
    const nextY = axis === 0 ? originY + wave : originY + along;
    pdf.line(prevX, prevY, nextX, nextY);
    prevX = nextX;
    prevY = nextY;
  }
}

/**
 * Explicit chevron polylines for geometric art borders `zigZag` /
 * `zigZagStitch`. Other decorative art styles remain skipped.
 */
function strokeZigZagParagraphBorderEdge(
  pdf: JsPdf,
  edge: WorkPdfParagraphBorderBoxEdge,
  x: number,
  y: number,
  width: number,
  height: number,
  thickness: number,
  stitch: boolean,
): void {
  const paintEdge = paragraphBorderPaintEdge(edge);
  const amplitude = Math.max(1.4, thickness * 1.35);
  const period = Math.max(5, thickness * 5);
  const offsets = stitch ? [-amplitude * 0.55, amplitude * 0.55] : [0];
  for (const offset of offsets) {
    if (paintEdge === 'top' || paintEdge === 'bottom') {
      const yBase = (paintEdge === 'top' ? y : y + height) + offset;
      strokeZigZagPolyline(pdf, x, yBase, width, 0, period, amplitude);
    } else {
      const xBase = (paintEdge === 'left' ? x : x + width) + offset;
      strokeZigZagPolyline(pdf, xBase, y, height, 1, period, amplitude);
    }
  }
}

/** axis: 0 = horizontal along +x, 1 = vertical along +y. */
function strokeZigZagPolyline(
  pdf: JsPdf,
  originX: number,
  originY: number,
  length: number,
  axis: 0 | 1,
  period: number,
  amplitude: number,
): void {
  if (
    !Number.isFinite(length) ||
    length <= 0 ||
    !Number.isFinite(period) ||
    period <= 0
  ) {
    return;
  }
  const teeth = Math.max(2, Math.ceil(length / period));
  let prevX = originX;
  let prevY = originY;
  for (let index = 1; index <= teeth * 2; index += 1) {
    const along = Math.min(length, (length * index) / (teeth * 2));
    const peak = index % 2 === 1 ? amplitude : -amplitude;
    const nextX = axis === 0 ? originX + along : originX + peak;
    const nextY = axis === 0 ? originY + peak : originY + along;
    pdf.line(prevX, prevY, nextX, nextY);
    prevX = nextX;
    prevY = nextY;
  }
}

/**
 * One-sided triangular teeth for geometric art borders `sawtooth` /
 * `sharksTeeth`. `sharksTeeth` densifies the period and flips amplitude.
 */
function strokeSawtoothParagraphBorderEdge(
  pdf: JsPdf,
  edge: WorkPdfParagraphBorderBoxEdge,
  x: number,
  y: number,
  width: number,
  height: number,
  thickness: number,
  sharks: boolean,
): void {
  const paintEdge = paragraphBorderPaintEdge(edge);
  const amplitude = Math.max(1.6, thickness * (sharks ? 1.55 : 1.4));
  const period = Math.max(4, thickness * (sharks ? 3.6 : 5.2));
  const signedAmplitude =
    sharks && (paintEdge === 'bottom' || paintEdge === 'right')
      ? -amplitude
      : amplitude;
  if (paintEdge === 'top' || paintEdge === 'bottom') {
    const yBase = paintEdge === 'top' ? y : y + height;
    strokeSawtoothPolyline(pdf, x, yBase, width, 0, period, signedAmplitude);
  } else {
    const xBase = paintEdge === 'left' ? x : x + width;
    strokeSawtoothPolyline(pdf, xBase, y, height, 1, period, signedAmplitude);
  }
}

/** axis: 0 = horizontal along +x, 1 = vertical along +y. */
function strokeSawtoothPolyline(
  pdf: JsPdf,
  originX: number,
  originY: number,
  length: number,
  axis: 0 | 1,
  period: number,
  amplitude: number,
): void {
  if (
    !Number.isFinite(length) ||
    length <= 0 ||
    !Number.isFinite(period) ||
    period <= 0
  ) {
    return;
  }
  const teeth = Math.max(2, Math.ceil(length / period));
  let prevX = originX;
  let prevY = originY;
  for (let index = 1; index <= teeth; index += 1) {
    const peakAlong = Math.min(length, ((index - 0.5) * length) / teeth);
    const endAlong = Math.min(length, (index * length) / teeth);
    const peakX = axis === 0 ? originX + peakAlong : originX + amplitude;
    const peakY = axis === 0 ? originY + amplitude : originY + peakAlong;
    const endX = axis === 0 ? originX + endAlong : originX;
    const endY = axis === 0 ? originY : originY + endAlong;
    pdf.line(prevX, prevY, peakX, peakY);
    pdf.line(peakX, peakY, endX, endY);
    prevX = endX;
    prevY = endY;
  }
}

/**
 * Closed isosceles triangles along the measured edge for geometric art
 * borders `triangles` / `triangle1` / `triangle2`. `triangle1` densifies;
 * `triangle2` inverts amplitude.
 */
function strokeTrianglesParagraphBorderEdge(
  pdf: JsPdf,
  edge: WorkPdfParagraphBorderBoxEdge,
  x: number,
  y: number,
  width: number,
  height: number,
  thickness: number,
  kind: 'triangles' | 'triangle1' | 'triangle2',
): void {
  const paintEdge = paragraphBorderPaintEdge(edge);
  const amplitude = Math.max(
    1.8,
    thickness * (kind === 'triangle1' ? 1.15 : 1.6),
  );
  const period = Math.max(4.5, thickness * (kind === 'triangle1' ? 3.4 : 5.4));
  let signedAmplitude = amplitude;
  if (kind === 'triangle2') signedAmplitude = -amplitude;
  if (paintEdge === 'bottom' || paintEdge === 'right') {
    signedAmplitude = -signedAmplitude;
  }
  if (paintEdge === 'top' || paintEdge === 'bottom') {
    const yBase = paintEdge === 'top' ? y : y + height;
    strokeClosedTrianglePolyline(
      pdf,
      x,
      yBase,
      width,
      0,
      period,
      signedAmplitude,
    );
  } else {
    const xBase = paintEdge === 'left' ? x : x + width;
    strokeClosedTrianglePolyline(
      pdf,
      xBase,
      y,
      height,
      1,
      period,
      signedAmplitude,
    );
  }
}

/** axis: 0 = horizontal along +x, 1 = vertical along +y. */
function strokeClosedTrianglePolyline(
  pdf: JsPdf,
  originX: number,
  originY: number,
  length: number,
  axis: 0 | 1,
  period: number,
  amplitude: number,
): void {
  if (
    !Number.isFinite(length) ||
    length <= 0 ||
    !Number.isFinite(period) ||
    period <= 0
  ) {
    return;
  }
  const teeth = Math.max(2, Math.ceil(length / period));
  for (let index = 0; index < teeth; index += 1) {
    const startAlong = (index * length) / teeth;
    const peakAlong = Math.min(length, ((index + 0.5) * length) / teeth);
    const endAlong = Math.min(length, ((index + 1) * length) / teeth);
    const startX = axis === 0 ? originX + startAlong : originX;
    const startY = axis === 0 ? originY : originY + startAlong;
    const peakX = axis === 0 ? originX + peakAlong : originX + amplitude;
    const peakY = axis === 0 ? originY + amplitude : originY + peakAlong;
    const endX = axis === 0 ? originX + endAlong : originX;
    const endY = axis === 0 ? originY : originY + endAlong;
    pdf.line(startX, startY, peakX, peakY);
    pdf.line(peakX, peakY, endX, endY);
    pdf.line(endX, endY, startX, startY);
  }
}

/**
 * Repeating ellipses along the measured edge for geometric art borders
 * `ovals` / `rings`. `rings` densifies and nests a second inner ellipse.
 */
function strokeOvalParagraphBorderEdge(
  pdf: JsPdf,
  edge: WorkPdfParagraphBorderBoxEdge,
  x: number,
  y: number,
  width: number,
  height: number,
  thickness: number,
  rings: boolean,
): void {
  const paintEdge = paragraphBorderPaintEdge(edge);
  const radius = Math.max(1.4, thickness * (rings ? 1.15 : 1.45));
  const period = Math.max(4.2, thickness * (rings ? 3.8 : 5.6));
  if (paintEdge === 'top' || paintEdge === 'bottom') {
    const yBase = paintEdge === 'top' ? y : y + height;
    strokeOvalMotifs(pdf, x, yBase, width, 0, period, radius, rings);
  } else {
    const xBase = paintEdge === 'left' ? x : x + width;
    strokeOvalMotifs(pdf, xBase, y, height, 1, period, radius, rings);
  }
}

/** axis: 0 = horizontal along +x, 1 = vertical along +y. */
function strokeOvalMotifs(
  pdf: JsPdf,
  originX: number,
  originY: number,
  length: number,
  axis: 0 | 1,
  period: number,
  radius: number,
  rings: boolean,
): void {
  if (
    !Number.isFinite(length) ||
    length <= 0 ||
    !Number.isFinite(period) ||
    period <= 0
  ) {
    return;
  }
  const count = Math.max(2, Math.ceil(length / period));
  for (let index = 0; index < count; index += 1) {
    const along = ((index + 0.5) * length) / count;
    const centerX = axis === 0 ? originX + along : originX;
    const centerY = axis === 0 ? originY : originY + along;
    const rx =
      axis === 0 ? Math.min(radius * 1.35, length / (count * 2.2)) : radius;
    const ry =
      axis === 0 ? radius : Math.min(radius * 1.35, length / (count * 2.2));
    strokeEllipsePolyline(pdf, centerX, centerY, rx, ry);
    if (rings) {
      strokeEllipsePolyline(pdf, centerX, centerY, rx * 0.55, ry * 0.55);
    }
  }
}

function strokeEllipsePolyline(
  pdf: JsPdf,
  centerX: number,
  centerY: number,
  radiusX: number,
  radiusY: number,
): void {
  if (
    !Number.isFinite(radiusX) ||
    !Number.isFinite(radiusY) ||
    radiusX <= 0 ||
    radiusY <= 0
  ) {
    return;
  }
  const ellipseCapable = pdf as JsPdf & {
    ellipse?: (
      x: number,
      y: number,
      rx: number,
      ry: number,
      style?: string | null,
    ) => void;
  };
  if (typeof ellipseCapable.ellipse === 'function') {
    ellipseCapable.ellipse(centerX, centerY, radiusX, radiusY, 'S');
    return;
  }
  const steps = 16;
  let prevX = centerX + radiusX;
  let prevY = centerY;
  for (let index = 1; index <= steps; index += 1) {
    const angle = (Math.PI * 2 * index) / steps;
    const nextX = centerX + Math.cos(angle) * radiusX;
    const nextY = centerY + Math.sin(angle) * radiusY;
    pdf.line(prevX, prevY, nextX, nextY);
    prevX = nextX;
    prevY = nextY;
  }
}

/**
 * Small rectangle motifs along the measured edge for geometric art borders
 * `marquee` / `marqueeToothed`. `marqueeToothed` alternates perpendicular
 * offset for a toothed look.
 */
function strokeMarqueeParagraphBorderEdge(
  pdf: JsPdf,
  edge: WorkPdfParagraphBorderBoxEdge,
  x: number,
  y: number,
  width: number,
  height: number,
  thickness: number,
  toothed: boolean,
): void {
  const paintEdge = paragraphBorderPaintEdge(edge);
  const cell = Math.max(2.2, thickness * (toothed ? 2.4 : 2.8));
  const depth = Math.max(1.6, thickness * (toothed ? 1.35 : 1.15));
  if (paintEdge === 'top' || paintEdge === 'bottom') {
    const yBase = paintEdge === 'top' ? y : y + height;
    const outward = paintEdge === 'top' ? -1 : 1;
    strokeMarqueeRectangles(
      pdf,
      x,
      yBase,
      width,
      0,
      cell,
      depth,
      outward,
      toothed,
    );
  } else {
    const xBase = paintEdge === 'left' ? x : x + width;
    const outward = paintEdge === 'left' ? -1 : 1;
    strokeMarqueeRectangles(
      pdf,
      xBase,
      y,
      height,
      1,
      cell,
      depth,
      outward,
      toothed,
    );
  }
}

/** axis: 0 = horizontal along +x, 1 = vertical along +y. */
function strokeMarqueeRectangles(
  pdf: JsPdf,
  originX: number,
  originY: number,
  length: number,
  axis: 0 | 1,
  cell: number,
  depth: number,
  outward: 1 | -1,
  toothed: boolean,
): void {
  if (
    !Number.isFinite(length) ||
    length <= 0 ||
    !Number.isFinite(cell) ||
    cell <= 0
  ) {
    return;
  }
  const count = Math.max(2, Math.ceil(length / cell));
  for (let index = 0; index < count; index += 1) {
    const startAlong = (index * length) / count;
    const endAlong = Math.min(length, ((index + 1) * length) / count);
    const span = Math.max(0.8, endAlong - startAlong - 0.35);
    const offset = toothed && index % 2 === 1 ? outward * depth * 0.55 : 0;
    if (axis === 0) {
      const left = originX + startAlong + 0.15;
      const top = originY + offset;
      const bottom = originY + outward * depth + offset;
      pdf.line(left, top, left + span, top);
      pdf.line(left + span, top, left + span, bottom);
      pdf.line(left + span, bottom, left, bottom);
      pdf.line(left, bottom, left, top);
    } else {
      const top = originY + startAlong + 0.15;
      const left = originX + offset;
      const right = originX + outward * depth + offset;
      pdf.line(left, top, right, top);
      pdf.line(right, top, right, top + span);
      pdf.line(right, top + span, left, top + span);
      pdf.line(left, top + span, left, top);
    }
  }
}

function applyWorkPdfBorderDash(
  pdf: JsPdf,
  kind: WorkPdfParagraphBorderKind,
  thickness: number,
): void {
  const dashable = pdf as JsPdf & {
    setLineDashPattern?: (pattern: number[], phase: number) => void;
  };
  if (!dashable.setLineDashPattern) return;
  if (kind === 'dashed') {
    dashable.setLineDashPattern(
      [Math.max(2, thickness * 3), Math.max(1.5, thickness * 2)],
      0,
    );
  } else if (kind === 'dotted') {
    dashable.setLineDashPattern(
      [Math.max(0.8, thickness), Math.max(1.2, thickness * 1.5)],
      0,
    );
  } else {
    dashable.setLineDashPattern([], 0);
  }
}

function clearWorkPdfBorderDash(pdf: JsPdf): void {
  const dashable = pdf as JsPdf & {
    setLineDashPattern?: (pattern: number[], phase: number) => void;
  };
  dashable.setLineDashPattern?.([], 0);
}

function shouldSkipPdfBorderElement(element: HTMLElement): boolean {
  if (element.closest('[aria-hidden="true"]')) {
    const hidden = element.closest('[aria-hidden="true"]');
    if (
      hidden &&
      !hidden.classList.contains('work-document-live-pdf-snapshot') &&
      !hidden.classList.contains('work-pdf-export-page')
    ) {
      return true;
    }
  }
  return false;
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
