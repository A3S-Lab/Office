import type { jsPDF as JsPdf } from 'jspdf';
import {
  queueWorkPdfUriLinkAnnotation,
  registerWorkPdfLinkStructEntry,
  type WorkPdfAnnotObjectRef,
} from './work-pdf-structure';
import type {
  WorkPdfPageBounds,
  WorkPdfPagePoints,
} from './work-pdf-text-layer';

/** One external hyperlink hotspot in page-local CSS pixels. */
export interface WorkPdfExternalLinkBox {
  /** Accessible label: visible link text, else href. */
  alt: string;
  height: number;
  href: string;
  width: number;
  x: number;
  y: number;
}

const MAX_EXTERNAL_LINKS = 2_048;
const MAX_HREF_LENGTH = 2_048;
const MAX_ALT_LENGTH = 2_048;
const HTTP_HREF_PATTERN = /^https?:\/\//i;

/**
 * Collects visible `a[href]` hotspots that use http(s) targets inside a page
 * bounds. Internal `#` anchors, empty, javascript:, and oversized hrefs are
 * omitted (fail soft). Not a PDF/UA certification claim.
 */
export function collectWorkPdfExternalLinkBoxes(
  root: ParentNode,
  page: WorkPdfPageBounds,
): WorkPdfExternalLinkBox[] {
  const links: WorkPdfExternalLinkBox[] = [];
  const anchors = Array.from(root.querySelectorAll('a[href]'));
  for (const node of anchors) {
    if (links.length >= MAX_EXTERNAL_LINKS) break;
    if (!(node instanceof HTMLAnchorElement)) continue;
    const href = normalizeExternalHref(node.getAttribute('href'));
    if (!href) continue;
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
    const text = (node.textContent ?? '').replace(/\s+/g, ' ').trim();
    const alt = (text || href).slice(0, MAX_ALT_LENGTH);
    links.push({ alt, height, href, width, x: clippedX, y: clippedY });
  }
  return links;
}

/**
 * Writes deferred URI link annotations for collected hotspots and registers
 * matching `/Link` StructElems with `/OBJR` association. Coordinates map from
 * page-local CSS pixels into PDF user space (origin top-left → jsPDF y-down
 * flip via getVerticalCoordinateString on putPage).
 */
export function appendWorkPdfExternalLinkAnnotations(
  pdf: JsPdf,
  links: readonly WorkPdfExternalLinkBox[],
  pageCss: WorkPdfPageBounds,
  pagePoints: WorkPdfPagePoints,
): void {
  if (!links.length) return;
  if (pageCss.width <= 0 || pageCss.height <= 0) return;
  if (pagePoints.pageWidthPoints <= 0 || pagePoints.pageHeightPoints <= 0) {
    return;
  }
  const scaleX = pagePoints.pageWidthPoints / pageCss.width;
  const scaleY = pagePoints.pageHeightPoints / pageCss.height;
  for (const link of links) {
    const href = normalizeExternalHref(link.href);
    if (!href) continue;
    const width = link.width * scaleX;
    const height = link.height * scaleY;
    if (!(width > 0) || !(height > 0)) continue;
    const x = link.x * scaleX;
    const y = pagePoints.pageHeightPoints - link.y * scaleY - height;
    const annotRef: WorkPdfAnnotObjectRef = { objectId: null };
    queueWorkPdfUriLinkAnnotation(pdf, {
      annotRef,
      height,
      href,
      width,
      x,
      y,
    });
    registerWorkPdfLinkStructEntry(pdf, {
      alt: link.alt?.trim() || href,
      annotRef,
      href,
    });
  }
}

function normalizeExternalHref(
  value: string | null | undefined,
): string | null {
  const href = value?.trim() ?? '';
  if (!href || href.length > MAX_HREF_LENGTH) return null;
  if (!HTTP_HREF_PATTERN.test(href)) return null;
  if (/[\s<>"']/.test(href)) return null;
  return href;
}
