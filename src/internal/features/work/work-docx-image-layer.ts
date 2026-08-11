import JSZip from 'jszip';
import {
  normalizeDocumentImageLayer,
  type WorkDocumentImageLayer,
  type WorkDocumentImageLayout,
} from './work-document-image-layout';
import { descendants, parseXml } from './work-ooxml-package';

interface DocxImageLayerPatch {
  marker: string;
  layer: WorkDocumentImageLayer;
}

export class DocxImageLayerPatchCollector {
  readonly patches: DocxImageLayerPatch[] = [];
  private nextMarker = 1;

  marker(
    layout: WorkDocumentImageLayout,
    layer: WorkDocumentImageLayer,
  ): string {
    if (layout === 'inline') return '';
    const marker = `__A3S_IMAGE_LAYER_${this.nextMarker}__`;
    this.nextMarker += 1;
    this.patches.push({ marker, layer: normalizeDocumentImageLayer(layer) });
    return marker;
  }
}

export async function patchDocxImageLayers(
  buffer: ArrayBuffer,
  patches: readonly DocxImageLayerPatch[],
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
      const patch = imageLayerPatch(properties, byMarker);
      if (!patch) continue;
      removeLayerMarker(properties, patch.marker);
      changed = true;
      const anchor = closestAnchor(properties);
      if (anchor) applyImageLayer(anchor, patch.layer);
    }
    if (changed) {
      archive.file(entry.name, new XMLSerializer().serializeToString(document));
    }
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function imageLayerPatch(
  properties: Element,
  patches: ReadonlyMap<string, DocxImageLayerPatch>,
): DocxImageLayerPatch | null {
  for (const property of Array.from(properties.attributes)) {
    for (const [marker, patch] of patches) {
      if (property.value.includes(marker)) return patch;
    }
  }
  return null;
}

function removeLayerMarker(properties: Element, marker: string): void {
  for (const property of Array.from(properties.attributes)) {
    if (!property.value.includes(marker)) continue;
    property.value = property.value.replace(marker, '');
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

function applyImageLayer(
  anchor: Element,
  source: WorkDocumentImageLayer,
): void {
  const layer = normalizeDocumentImageLayer(source);
  anchor.setAttribute('relativeHeight', String(layer.relativeHeight));
  anchor.setAttribute('behindDoc', layer.behindDocument ? '1' : '0');
  anchor.setAttribute('allowOverlap', layer.allowOverlap ? '1' : '0');
  anchor.setAttribute('layoutInCell', layer.layoutInCell ? '1' : '0');
  anchor.setAttribute('locked', layer.lockAnchor ? '1' : '0');
}
