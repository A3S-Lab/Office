import JSZip from 'jszip';
import {
  applyDocumentImageIdentityToElement,
  createDocumentImageIdentityRegistry,
  documentImageIdentityFromElement,
  uniqueDocumentImageIdentity,
  type WorkDocumentImageIdentity,
  type WorkDocumentImageIdentityRegistry,
} from './work-document-image-identity';
import { descendants, parseXml } from './work-ooxml-package';

interface DocxImageIdentityPatch {
  marker: string;
  identity: WorkDocumentImageIdentity;
}

export interface DocxImageIdentityRegistration {
  marker: string;
  docPropertiesId: number;
}

const WORDPROCESSING_DRAWING_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';

export class DocxImageIdentityPatchCollector {
  readonly patches: DocxImageIdentityPatch[] = [];
  private readonly registry: WorkDocumentImageIdentityRegistry =
    createDocumentImageIdentityRegistry();
  private nextMarker = 1;

  register(element: HTMLImageElement): DocxImageIdentityRegistration {
    const identity = uniqueDocumentImageIdentity(
      documentImageIdentityFromElement(element),
      this.registry,
    );
    applyDocumentImageIdentityToElement(element, identity);
    const marker = `__A3S_IMAGE_IDENTITY_${this.nextMarker}__`;
    this.nextMarker += 1;
    this.patches.push({ marker, identity });
    return { marker, docPropertiesId: identity.docPropertiesId };
  }
}

export async function patchDocxImageIdentities(
  buffer: ArrayBuffer,
  patches: readonly DocxImageIdentityPatch[],
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
      const patch = imageIdentityPatch(properties, byMarker);
      if (!patch) continue;
      removeIdentityMarker(properties, patch.marker);
      properties.setAttribute('id', String(patch.identity.docPropertiesId));
      const drawing = closestDrawing(properties);
      if (drawing) {
        ensureWordprocessingDrawing2010Namespace(document);
        drawing.setAttributeNS(
          WORDPROCESSING_DRAWING_2010_NAMESPACE,
          'wp14:anchorId',
          patch.identity.anchorId,
        );
        drawing.setAttributeNS(
          WORDPROCESSING_DRAWING_2010_NAMESPACE,
          'wp14:editId',
          patch.identity.editId,
        );
      }
      changed = true;
    }
    if (changed) {
      archive.file(entry.name, new XMLSerializer().serializeToString(document));
    }
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function imageIdentityPatch(
  properties: Element,
  patches: ReadonlyMap<string, DocxImageIdentityPatch>,
): DocxImageIdentityPatch | null {
  for (const property of Array.from(properties.attributes)) {
    for (const [marker, patch] of patches) {
      if (property.value.includes(marker)) return patch;
    }
  }
  return null;
}

function removeIdentityMarker(properties: Element, marker: string): void {
  for (const property of Array.from(properties.attributes)) {
    if (!property.value.includes(marker)) continue;
    property.value = property.value.replace(marker, '');
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

function ensureWordprocessingDrawing2010Namespace(document: Document): void {
  const root = document.documentElement;
  if (
    root.lookupNamespaceURI?.('wp14') === WORDPROCESSING_DRAWING_2010_NAMESPACE
  ) {
    return;
  }
  root.setAttributeNS(
    XMLNS_NAMESPACE,
    'xmlns:wp14',
    WORDPROCESSING_DRAWING_2010_NAMESPACE,
  );
}
