import JSZip from 'jszip';
import type { WorkDocumentImageCrop } from './work-document-image-layout';
import { descendants, parseXml } from './work-ooxml-package';

interface DocxImageCropPatch {
  marker: string;
  crop: WorkDocumentImageCrop;
}

export class DocxImageCropPatchCollector {
  readonly patches: DocxImageCropPatch[] = [];
  private nextMarker = 1;

  marker(crop: WorkDocumentImageCrop | null): string {
    if (!crop) return '';
    const marker = `__A3S_IMAGE_CROP_${this.nextMarker}__`;
    this.nextMarker += 1;
    this.patches.push({ marker, crop });
    return marker;
  }
}

export async function patchDocxImageCrops(
  buffer: ArrayBuffer,
  patches: readonly DocxImageCropPatch[],
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
      const patch = imageCropPatch(properties, byMarker);
      if (!patch) continue;
      removeCropMarker(properties, patch.marker);
      changed = true;
      const drawing = closestDrawing(properties);
      const source = drawing ? descendants(drawing, 'srcRect')[0] : undefined;
      if (!source) continue;
      applyCrop(source, patch.crop);
    }
    if (changed) {
      archive.file(entry.name, new XMLSerializer().serializeToString(document));
    }
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function imageCropPatch(
  properties: Element,
  patches: ReadonlyMap<string, DocxImageCropPatch>,
): DocxImageCropPatch | null {
  for (const attribute of Array.from(properties.attributes)) {
    for (const [marker, patch] of patches) {
      if (attribute.value.includes(marker)) return patch;
    }
  }
  return null;
}

function removeCropMarker(properties: Element, marker: string): void {
  for (const attribute of Array.from(properties.attributes)) {
    if (!attribute.value.includes(marker)) continue;
    attribute.value = attribute.value.replace(marker, '');
  }
}

function closestDrawing(element: Element): Element | null {
  let current: Element | null = element;
  while (current) {
    if (current.localName === 'anchor' || current.localName === 'inline') {
      return current;
    }
    current = current.parentElement;
  }
  return null;
}

function applyCrop(source: Element, crop: WorkDocumentImageCrop): void {
  for (const [name, value] of [
    ['t', crop.top],
    ['r', crop.right],
    ['b', crop.bottom],
    ['l', crop.left],
  ] as const) {
    if (value > 0) source.setAttribute(name, String(Math.round(value * 1_000)));
    else source.removeAttribute(name);
  }
}
