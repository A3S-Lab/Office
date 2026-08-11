import JSZip from 'jszip';
import {
  isContourImageLayout,
  type WorkDocumentImageLayoutOptions,
} from './work-document-image-layout';
import {
  defaultDocumentImageWrapContour,
  type WorkDocumentImageWrapContour,
} from './work-document-image-wrap-contour';
import { descendants, directChildren, parseXml } from './work-ooxml-package';

interface DocxImageWrapPatch {
  marker: string;
  layout: 'through' | 'tight';
  options: WorkDocumentImageLayoutOptions;
  contour: WorkDocumentImageWrapContour;
}

const WORDPROCESSING_DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';

export class DocxImageWrapPatchCollector {
  readonly patches: DocxImageWrapPatch[] = [];
  private nextMarker = 1;

  marker(
    options: WorkDocumentImageLayoutOptions,
    contour: WorkDocumentImageWrapContour | null,
  ): string {
    if (!isContourImageLayout(options.layout)) return '';
    const marker = `__A3S_IMAGE_WRAP_${this.nextMarker}__`;
    this.nextMarker += 1;
    this.patches.push({
      marker,
      layout: options.layout,
      options,
      contour: contour ?? defaultDocumentImageWrapContour(),
    });
    return marker;
  }
}

export async function patchDocxImageWraps(
  buffer: ArrayBuffer,
  patches: readonly DocxImageWrapPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  const byMarker = new Map(patches.map((patch) => [patch.marker, patch]));
  const entries = Object.values(archive.files).filter(
    (entry) =>
      !entry.dir &&
      /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/.test(
        entry.name,
      ),
  );
  for (const entry of entries) {
    const document = parseXml(await entry.async('text'), entry.name);
    let changed = false;
    for (const properties of descendants(document, 'docPr')) {
      const patch = imageWrapPatch(properties, byMarker);
      if (!patch) continue;
      removeWrapMarker(properties, patch.marker);
      changed = true;
      const anchor = closestAnchor(properties);
      const source = anchor ? wrappingElement(anchor) : null;
      if (!anchor || !source) continue;
      source.parentNode?.replaceChild(
        createWrappingElement(anchor, patch),
        source,
      );
    }
    if (changed) {
      archive.file(entry.name, new XMLSerializer().serializeToString(document));
    }
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function imageWrapPatch(
  properties: Element,
  patches: ReadonlyMap<string, DocxImageWrapPatch>,
): DocxImageWrapPatch | null {
  for (const attribute of Array.from(properties.attributes)) {
    for (const [marker, patch] of patches) {
      if (attribute.value.includes(marker)) return patch;
    }
  }
  return null;
}

function removeWrapMarker(properties: Element, marker: string): void {
  for (const attribute of Array.from(properties.attributes)) {
    if (attribute.value.includes(marker)) {
      attribute.value = attribute.value.replace(marker, '');
    }
  }
}

function closestAnchor(element: Element): Element | null {
  let current: Element | null = element;
  while (current) {
    if (current.localName === 'anchor') return current;
    current = current.parentElement;
  }
  return null;
}

function wrappingElement(anchor: Element): Element | null {
  return (
    directChildren(anchor).find((element) =>
      [
        'wrapNone',
        'wrapSquare',
        'wrapTight',
        'wrapThrough',
        'wrapTopAndBottom',
      ].includes(element.localName),
    ) ?? null
  );
}

function createWrappingElement(
  anchor: Element,
  patch: DocxImageWrapPatch,
): Element {
  const document = anchor.ownerDocument;
  const namespace = anchor.namespaceURI ?? WORDPROCESSING_DRAWING_NAMESPACE;
  const prefix = anchor.prefix ? `${anchor.prefix}:` : '';
  const wrap = document.createElementNS(
    namespace,
    `${prefix}${patch.layout === 'tight' ? 'wrapTight' : 'wrapThrough'}`,
  );
  wrap.setAttribute('wrapText', patch.options.wrapSide);
  const distance = String(Math.round(patch.options.wrapDistance * 36_000));
  wrap.setAttribute('distL', distance);
  wrap.setAttribute('distR', distance);

  const polygon = document.createElementNS(namespace, `${prefix}wrapPolygon`);
  polygon.setAttribute('edited', patch.contour.edited ? '1' : '0');
  patch.contour.points.forEach((point, index) => {
    const child = document.createElementNS(
      namespace,
      `${prefix}${index === 0 ? 'start' : 'lineTo'}`,
    );
    child.setAttribute('x', String(point.x));
    child.setAttribute('y', String(point.y));
    polygon.append(child);
  });
  wrap.append(polygon);
  return wrap;
}
