import type { jsPDF as JsPdf } from 'jspdf';
import { registerWorkPdfFigureStructEntry } from './work-pdf-structure';
import type { WorkPdfPageBounds } from './work-pdf-text-layer';

/** One alt-text figure candidate in page-local CSS pixels. */
export interface WorkPdfFigureBox {
  alt: string;
  height: number;
  width: number;
  x: number;
  y: number;
}

const MAX_FIGURES = 2_048;
const MAX_ALT_LENGTH = 2_048;

/**
 * Collects visible `img` elements with non-empty `alt` or `title` inside page
 * bounds. Decorative images without alt/title are omitted (fail soft). Not a
 * PDF/UA certification claim.
 */
export function collectWorkPdfFigureBoxes(
  root: ParentNode,
  page: WorkPdfPageBounds,
): WorkPdfFigureBox[] {
  const figures: WorkPdfFigureBox[] = [];
  const images = Array.from(root.querySelectorAll('img'));
  for (const node of images) {
    if (figures.length >= MAX_FIGURES) break;
    if (!(node instanceof HTMLImageElement)) continue;
    const alt = normalizeFigureAlt(
      node.getAttribute('alt'),
      node.getAttribute('title'),
    );
    if (!alt) continue;
    let rect: DOMRect;
    try {
      rect = node.getBoundingClientRect();
    } catch {
      continue;
    }
    if (
      !Number.isFinite(rect.width) ||
      !Number.isFinite(rect.height) ||
      rect.width <= 0 ||
      rect.height <= 0
    ) {
      continue;
    }
    const x = rect.left - page.left;
    const y = rect.top - page.top;
    const right = x + rect.width;
    const bottom = y + rect.height;
    if (right <= 0 || bottom <= 0 || x >= page.width || y >= page.height) {
      continue;
    }
    const clippedX = Math.max(0, x);
    const clippedY = Math.max(0, y);
    const clippedRight = Math.min(page.width, right);
    const clippedBottom = Math.min(page.height, bottom);
    const width = clippedRight - clippedX;
    const height = clippedBottom - clippedY;
    if (width <= 0 || height <= 0) continue;
    figures.push({ alt, height, width, x: clippedX, y: clippedY });
  }
  return figures;
}

/**
 * Registers `/Figure` StructElems for collected alt-text images on the current
 * PDF page. Raster paint already embeds the pixels; this deepens the tagged
 * tree without claiming PDF/UA certification.
 */
export function appendWorkPdfFigureStructEntries(
  pdf: JsPdf,
  figures: readonly WorkPdfFigureBox[],
): void {
  for (const figure of figures) {
    const alt = figure.alt.trim().slice(0, MAX_ALT_LENGTH);
    if (!alt) continue;
    registerWorkPdfFigureStructEntry(pdf, { alt });
  }
}

function normalizeFigureAlt(
  alt: string | null,
  title: string | null,
): string | null {
  const value = (alt?.trim() || title?.trim() || '').slice(0, MAX_ALT_LENGTH);
  return value || null;
}
