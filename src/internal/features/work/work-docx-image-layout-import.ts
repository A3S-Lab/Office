import {
  normalizeDocumentImageAlignment,
  normalizeDocumentImageLayoutOptions,
  type WorkDocumentImageLayout,
  type WorkDocumentImageLayoutOptions,
} from './work-document-image-layout';
import { attribute, descendants, directChild } from './work-ooxml-package';

export interface ImportedDocxImageLayoutMarker {
  startMarker: string;
  endMarker: string;
  options: WorkDocumentImageLayoutOptions;
}

export interface ImportedDocxImageLayoutMarkers {
  images: ImportedDocxImageLayoutMarker[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const EMUS_PER_MILLIMETER = 36_000;
const IMAGE_LAYOUT_MARKER_PATTERN =
  /__A3S_WORK_IMAGE_LAYOUT_\d+_(?:START|END)__/g;

export function markDocxImageLayouts(
  document: Document,
): ImportedDocxImageLayoutMarkers {
  const images: ImportedDocxImageLayoutMarker[] = [];
  const markedRuns = new Set<Element>();
  for (const anchor of descendants(document, 'anchor')) {
    const run = closestAncestor(anchor, 'r');
    if (!run || markedRuns.has(run)) continue;
    const layout = anchorLayout(anchor);
    if (!layout) continue;
    markedRuns.add(run);
    const index = images.length + 1;
    const startMarker = `__A3S_WORK_IMAGE_LAYOUT_${index}_START__`;
    const endMarker = `__A3S_WORK_IMAGE_LAYOUT_${index}_END__`;
    insertMarkerRun(document, run, startMarker, 'before');
    insertMarkerRun(document, run, endMarker, 'after');
    images.push({
      startMarker,
      endMarker,
      options: normalizeDocumentImageLayoutOptions({
        layout,
        alignment: anchorAlignment(anchor),
        wrapDistance: anchorWrapDistance(anchor, layout),
      }),
    });
  }
  return { images };
}

export function applyImportedDocxImageLayoutMarkers(
  document: Document,
  markers: ImportedDocxImageLayoutMarkers,
): void {
  const starts = new Map(
    markers.images.map((image) => [image.startMarker, image]),
  );
  const ends = new Set(markers.images.map((image) => image.endMarker));
  const walker = document.createTreeWalker(
    document.body,
    NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT,
  );
  const nodes: Node[] = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);
  const state: {
    active: ImportedDocxImageLayoutMarker | null;
    applied: boolean;
  } = { active: null, applied: false };
  for (const node of nodes) {
    if (node instanceof Text) {
      node.data = node.data.replace(IMAGE_LAYOUT_MARKER_PATTERN, (marker) => {
        const start = starts.get(marker);
        if (start) {
          state.active = start;
          state.applied = false;
        } else if (ends.has(marker)) {
          state.active = null;
          state.applied = false;
        }
        return '';
      });
      continue;
    }
    if (!state.applied && state.active && node instanceof HTMLImageElement) {
      applyImageLayout(node, state.active.options);
      state.applied = true;
    }
  }
  document.body.normalize();
}

export function hasImportedDocxImageLayoutMarkers(
  markers: ImportedDocxImageLayoutMarkers,
): boolean {
  return markers.images.length > 0;
}

function anchorLayout(anchor: Element): WorkDocumentImageLayout | null {
  if (directChild(anchor, 'wrapTopAndBottom')) return 'topBottom';
  if (
    directChild(anchor, 'wrapSquare') ||
    directChild(anchor, 'wrapTight') ||
    directChild(anchor, 'wrapThrough')
  ) {
    return 'square';
  }
  return null;
}

function anchorAlignment(anchor: Element): string {
  const position = directChild(anchor, 'positionH');
  const alignment = position ? directChild(position, 'align') : undefined;
  const value = alignment?.textContent?.trim();
  if (value === 'inside') return 'left';
  if (value === 'outside') return 'right';
  return normalizeDocumentImageAlignment(value);
}

function anchorWrapDistance(
  anchor: Element,
  layout: WorkDocumentImageLayout,
): number {
  const wrap =
    directChild(
      anchor,
      layout === 'square' ? 'wrapSquare' : 'wrapTopAndBottom',
    ) ??
    (layout === 'square'
      ? (directChild(anchor, 'wrapTight') ?? directChild(anchor, 'wrapThrough'))
      : undefined);
  const names =
    layout === 'square'
      ? (['distL', 'distR'] as const)
      : (['distT', 'distB'] as const);
  const distance = Math.max(
    0,
    ...names.flatMap((name) => [
      numericAttribute(wrap, name) ?? 0,
      numericAttribute(anchor, name) ?? 0,
    ]),
  );
  return distance / EMUS_PER_MILLIMETER;
}

function insertMarkerRun(
  document: Document,
  run: Element,
  marker: string,
  position: 'after' | 'before',
): void {
  const markerRun = document.createElementNS(WORD_NAMESPACE, 'w:r');
  const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  markerRun.append(text);
  const parent = run.parentNode;
  if (!parent) return;
  parent.insertBefore(markerRun, position === 'before' ? run : run.nextSibling);
}

function applyImageLayout(
  image: HTMLImageElement,
  options: WorkDocumentImageLayoutOptions,
): void {
  image.dataset.officeImageLayout = options.layout;
  image.dataset.officeImageAlignment = options.alignment;
  image.dataset.officeImageWrapDistance = formatNumber(options.wrapDistance);
  image.style.setProperty(
    '--work-document-image-wrap-distance',
    `${formatNumber(options.wrapDistance)}mm`,
  );
}

function numericAttribute(
  element: Element | null | undefined,
  name: string,
): number | null {
  if (!element) return null;
  const source = attribute(element, name);
  if (source === null || !source.trim()) return null;
  const value = Number(source);
  return Number.isFinite(value) && value >= 0 ? value : null;
}

function closestAncestor(element: Element, localName: string): Element | null {
  let current: Element | null = element;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
}

function formatNumber(value: number): string {
  return Number(value.toFixed(2)).toString();
}
