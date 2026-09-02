import JSZip from 'jszip';
import {
  normalizeDocumentImageTransform,
  type WorkDocumentImageTransform,
} from './work-document-image-layout';
import {
  attribute,
  descendants,
  directChild,
  parseXml,
} from './work-ooxml-package';

interface DocxImageTransformPatch {
  marker: string;
  transform: WorkDocumentImageTransform;
}

const DOCX_ROTATION_UNITS_PER_DEGREE = 60_000;
const DOCX_QUADRANT_ROTATION_UNITS = 5_400_000;

export interface DocxImageTransformReadResult {
  transform: WorkDocumentImageTransform | null;
  supported: boolean;
}

export function readDocxImageTransform(
  drawing: Element,
): DocxImageTransformReadResult {
  const target = pictureTransform(drawing);
  if (!target) return { transform: null, supported: true };
  const rawRotation = attribute(target, 'rot');
  const rawFlipHorizontal = attribute(target, 'flipH');
  const rawFlipVertical = attribute(target, 'flipV');
  const rotation = parseRotation(rawRotation);
  const flipHorizontal = parseFlip(rawFlipHorizontal);
  const flipVertical = parseFlip(rawFlipVertical);
  const supported =
    rotation !== null &&
    flipHorizontal !== null &&
    flipVertical !== null &&
    (rotation === 0 || rotation % DOCX_QUADRANT_ROTATION_UNITS === 0);
  const normalized = normalizeDocumentImageTransform({
    rotation: rotation === null ? 0 : rotation / DOCX_ROTATION_UNITS_PER_DEGREE,
    flipHorizontal: flipHorizontal ?? false,
    flipVertical: flipVertical ?? false,
  });
  return {
    transform:
      normalized.rotation === 0 &&
      !normalized.flipHorizontal &&
      !normalized.flipVertical
        ? null
        : normalized,
    supported,
  };
}

export class DocxImageTransformPatchCollector {
  readonly patches: DocxImageTransformPatch[] = [];
  private nextMarker = 1;

  marker(transform: WorkDocumentImageTransform | null): string {
    if (!transform) return '';
    const normalized = normalizeDocumentImageTransform(transform);
    if (
      normalized.rotation === 0 &&
      !normalized.flipHorizontal &&
      !normalized.flipVertical
    ) {
      return '';
    }
    const marker = `__A3S_IMAGE_TRANSFORM_${this.nextMarker}__`;
    this.nextMarker += 1;
    this.patches.push({ marker, transform: normalized });
    return marker;
  }
}

export async function patchDocxImageTransforms(
  buffer: ArrayBuffer,
  patches: readonly DocxImageTransformPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  const entries = Object.values(archive.files).filter(
    (entry) =>
      !entry.dir &&
      /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/.test(
        entry.name,
      ),
  );
  const byMarker = new Map(patches.map((patch) => [patch.marker, patch]));
  for (const entry of entries) {
    const document = parseXml(await entry.async('text'), entry.name);
    let changed = false;
    for (const properties of descendants(document, 'docPr')) {
      const patch = imageTransformPatch(properties, byMarker);
      if (!patch) continue;
      removeTransformMarker(properties, patch.marker);
      // Always rewrite the part after consuming the marker. This prevents an
      // implementation detail from leaking into the exported DOCX even when
      // a malformed drawing has no editable picture transform node.
      changed = true;
      const drawing = closestDrawing(properties);
      const transform = drawing ? pictureTransform(drawing) : undefined;
      if (!transform) continue;
      applyImageTransform(transform, patch.transform);
    }
    if (changed) {
      archive.file(entry.name, new XMLSerializer().serializeToString(document));
    }
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function imageTransformPatch(
  properties: Element,
  patches: ReadonlyMap<string, DocxImageTransformPatch>,
): DocxImageTransformPatch | null {
  for (const attribute of Array.from(properties.attributes)) {
    for (const [marker, patch] of patches) {
      if (attribute.value.includes(marker)) return patch;
    }
  }
  return null;
}

function removeTransformMarker(properties: Element, marker: string): void {
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

function pictureTransform(drawing: Element): Element | undefined {
  const picture = descendants(drawing, 'pic')[0];
  const properties = picture ? directChild(picture, 'spPr') : undefined;
  return properties
    ? directChild(properties, 'xfrm')
    : descendants(drawing, 'xfrm')[0];
}

function parseRotation(value: string | null): number | null {
  if (value === null || !value.trim()) return 0;
  const number = Number(value);
  return Number.isInteger(number) && number >= 0 ? number : null;
}

function parseFlip(value: string | null): boolean | null {
  if (value === null || !value.trim()) return false;
  const normalized = value.trim().toLowerCase();
  if (normalized === '1' || normalized === 'true') return true;
  if (normalized === '0' || normalized === 'false') return false;
  return null;
}

function applyImageTransform(
  target: Element,
  transform: WorkDocumentImageTransform,
): void {
  const normalized = normalizeDocumentImageTransform(transform);
  if (normalized.rotation === 0) {
    target.removeAttribute('rot');
  } else {
    target.setAttribute(
      'rot',
      String(normalized.rotation * DOCX_ROTATION_UNITS_PER_DEGREE),
    );
  }
  if (normalized.flipHorizontal) target.setAttribute('flipH', '1');
  else target.removeAttribute('flipH');
  if (normalized.flipVertical) target.setAttribute('flipV', '1');
  else target.removeAttribute('flipV');
}
