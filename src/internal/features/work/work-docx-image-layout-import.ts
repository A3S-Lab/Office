import {
  applyDocumentImageLayerToElement,
  applyDocumentImageCropToElement,
  isContourImageLayout,
  normalizeDocumentImageAlignment,
  normalizeDocumentImageLayer,
  normalizeDocumentImageLayoutOptions,
  normalizeDocumentImageCrop,
  normalizeDocumentImagePosition,
  wrapsBesideImage,
  type WorkDocumentImageLayout,
  type WorkDocumentImageCrop,
  type WorkDocumentImageLayer,
  type WorkDocumentImageLayoutOptions,
  type WorkDocumentImagePosition,
} from './work-document-image-layout';
import {
  applyDocumentImageWrapContourToElement,
  normalizeDocumentImageWrapContour,
  normalizeDocumentImageWrapSide,
  type WorkDocumentImageWrapContour,
} from './work-document-image-wrap-contour';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
} from './work-ooxml-package';

export interface ImportedDocxImageLayoutMarker {
  startMarker: string;
  endMarker: string;
  options: WorkDocumentImageLayoutOptions;
  position: WorkDocumentImagePosition | null;
  crop: WorkDocumentImageCrop | null;
  contour: WorkDocumentImageWrapContour | null;
  layer: WorkDocumentImageLayer | null;
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
  const drawings = [
    ...descendants(document, 'anchor'),
    ...descendants(document, 'inline'),
  ];
  for (const anchor of drawings) {
    const run = closestAncestor(anchor, 'r');
    if (!run || markedRuns.has(run)) continue;
    const crop = anchorCrop(anchor);
    const layout =
      anchor.localName === 'inline'
        ? crop
          ? 'inline'
          : null
        : anchorLayout(anchor);
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
        wrapSide: anchorWrapSide(anchor, layout),
      }),
      position: anchorPosition(anchor),
      crop,
      contour: isContourImageLayout(layout)
        ? anchorWrapContour(anchor, layout)
        : null,
      layer: anchor.localName === 'anchor' ? anchorLayer(anchor) : null,
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
      applyImageLayout(
        node,
        state.active.options,
        state.active.position,
        state.active.crop,
        state.active.contour,
        state.active.layer,
      );
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
  if (directChild(anchor, 'wrapSquare')) return 'square';
  if (directChild(anchor, 'wrapTight')) return 'tight';
  if (directChild(anchor, 'wrapThrough')) return 'through';
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
  const wrap = anchorWrapElement(anchor, layout);
  const names =
    layout === 'topBottom'
      ? (['distT', 'distB'] as const)
      : (['distL', 'distR'] as const);
  const distance = Math.max(
    0,
    ...names.flatMap((name) => [
      numericAttribute(wrap, name) ?? 0,
      numericAttribute(anchor, name) ?? 0,
    ]),
  );
  return distance / EMUS_PER_MILLIMETER;
}

function anchorWrapSide(
  anchor: Element,
  layout: WorkDocumentImageLayout,
): string {
  return normalizeDocumentImageWrapSide(
    attribute(anchorWrapElement(anchor, layout) ?? anchor, 'wrapText'),
  );
}

function anchorWrapElement(
  anchor: Element,
  layout: WorkDocumentImageLayout,
): Element | undefined {
  if (layout === 'square') return directChild(anchor, 'wrapSquare');
  if (layout === 'tight') return directChild(anchor, 'wrapTight');
  if (layout === 'through') return directChild(anchor, 'wrapThrough');
  if (layout === 'topBottom') return directChild(anchor, 'wrapTopAndBottom');
  return undefined;
}

function anchorWrapContour(
  anchor: Element,
  layout: 'through' | 'tight',
): WorkDocumentImageWrapContour | null {
  const wrap = anchorWrapElement(anchor, layout);
  const polygon = wrap ? directChild(wrap, 'wrapPolygon') : undefined;
  if (!polygon) return null;
  const vertices = directChildren(polygon);
  if (
    vertices[0]?.localName !== 'start' ||
    vertices.slice(1).some((element) => element.localName !== 'lineTo')
  ) {
    return null;
  }
  return normalizeDocumentImageWrapContour({
    edited: attribute(polygon, 'edited'),
    points: vertices.map((point) => ({
      x: attribute(point, 'x'),
      y: attribute(point, 'y'),
    })),
  });
}

function anchorPosition(anchor: Element): WorkDocumentImagePosition | null {
  const horizontal = directChild(anchor, 'positionH');
  const vertical = directChild(anchor, 'positionV');
  const horizontalOffset = positionOffset(horizontal);
  const verticalOffset = positionOffset(vertical);
  if (horizontalOffset === null && verticalOffset === null) return null;
  return normalizeDocumentImagePosition({
    horizontalOffset,
    verticalOffset,
    horizontalReference: horizontal
      ? attribute(horizontal, 'relativeFrom')
      : null,
    verticalReference: vertical ? attribute(vertical, 'relativeFrom') : null,
  });
}

function anchorCrop(anchor: Element): WorkDocumentImageCrop | null {
  const source = descendants(anchor, 'srcRect')[0];
  if (!source) return null;
  return normalizeDocumentImageCrop({
    cropTop: cropPercentage(source, 't'),
    cropRight: cropPercentage(source, 'r'),
    cropBottom: cropPercentage(source, 'b'),
    cropLeft: cropPercentage(source, 'l'),
  });
}

function anchorLayer(anchor: Element): WorkDocumentImageLayer {
  return normalizeDocumentImageLayer({
    relativeHeight: attribute(anchor, 'relativeHeight'),
    behindDocument: anchorBooleanAttribute(anchor, 'behindDoc', false),
    allowOverlap: anchorBooleanAttribute(anchor, 'allowOverlap', false),
    layoutInCell: anchorBooleanAttribute(anchor, 'layoutInCell', true),
    lockAnchor: anchorBooleanAttribute(anchor, 'locked', false),
  });
}

function cropPercentage(source: Element, name: string): number {
  const value = attribute(source, name)?.trim() ?? '';
  if (!value) return 0;
  const number = Number(value.endsWith('%') ? value.slice(0, -1) : value);
  if (!Number.isFinite(number)) return 0;
  return value.endsWith('%') ? number : number / 1_000;
}

function positionOffset(position: Element | null | undefined): number | null {
  const value = position ? directChild(position, 'posOffset') : undefined;
  if (!value?.textContent?.trim()) return null;
  const emus = Number(value.textContent);
  return Number.isFinite(emus) ? emus / EMUS_PER_MILLIMETER : null;
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
  position: WorkDocumentImagePosition | null,
  crop: WorkDocumentImageCrop | null,
  contour: WorkDocumentImageWrapContour | null,
  layer: WorkDocumentImageLayer | null,
): void {
  image.dataset.officeImageLayout = options.layout;
  image.dataset.officeImageAlignment = options.alignment;
  image.dataset.officeImageWrapDistance = formatNumber(options.wrapDistance);
  if (wrapsBesideImage(options.layout) || options.wrapSide !== 'bothSides') {
    image.dataset.officeImageWrapSide = options.wrapSide;
  } else {
    delete image.dataset.officeImageWrapSide;
  }
  image.style.setProperty(
    '--work-document-image-wrap-distance',
    `${formatNumber(options.wrapDistance)}mm`,
  );
  if (position) {
    if (position.horizontalOffset !== null) {
      image.dataset.officeImageHorizontalOffset = formatNumber(
        position.horizontalOffset,
      );
      image.style.setProperty(
        '--work-document-image-horizontal-offset',
        `${formatNumber(position.horizontalOffset)}mm`,
      );
    }
    if (position.verticalOffset !== null) {
      image.dataset.officeImageVerticalOffset = formatNumber(
        position.verticalOffset,
      );
      image.style.setProperty(
        '--work-document-image-vertical-offset',
        `${formatNumber(position.verticalOffset)}mm`,
      );
    }
    image.dataset.officeImageHorizontalReference = position.horizontalReference;
    image.dataset.officeImageVerticalReference = position.verticalReference;
  }
  applyDocumentImageCropToElement(image, crop);
  applyDocumentImageWrapContourToElement(image, contour);
  if (layer) applyDocumentImageLayerToElement(image, layer);
}

function anchorBooleanAttribute(
  element: Element,
  name: string,
  fallback: boolean,
): boolean {
  const value = attribute(element, name)?.trim().toLowerCase();
  if (value === '1' || value === 'true' || value === 'on') return true;
  if (value === '0' || value === 'false' || value === 'off') return false;
  return fallback;
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
