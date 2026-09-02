import JSZip from 'jszip';
import { descendants, parseXml } from './work-ooxml-package';
import type { WorkDocumentTextBoxProperties } from './work-document-text-box';

interface DocxTextBoxIdentityPatch {
  marker: string;
  preferredId: number | null;
}

export interface DocxTextBoxIdentityRegistration {
  marker: string;
  docPropertiesId: number | null;
}

const DOC_PROPERTIES_ID_MAX = 0xffff_ffff;
const TEXT_BOX_ID_MARKER_PATTERN = /__A3S_TEXT_BOX_ID_\d+__/g;

/**
 * Allocates a marker for each text box and lets the package patcher assign a
 * document-wide `wp:docPr/@id` after all image and shape IDs are known.
 *
 * The `docx` package creates a fresh ID generator per drawing instance, so
 * relying on its default would give every WPS shape the same ID. Delaying the
 * assignment also lets us preserve an imported ID when it is still available,
 * while deterministically repairing collisions with images or other shapes.
 */
export class DocxTextBoxIdentityPatchCollector {
  readonly patches: DocxTextBoxIdentityPatch[] = [];
  private nextMarker = 1;

  register(
    properties: WorkDocumentTextBoxProperties,
  ): DocxTextBoxIdentityRegistration {
    const marker = `__A3S_TEXT_BOX_ID_${this.nextMarker}__`;
    this.nextMarker += 1;
    this.patches.push({ marker, preferredId: properties.docPropertiesId });
    return { marker, docPropertiesId: properties.docPropertiesId };
  }
}

export async function patchDocxTextBoxIdentities(
  buffer: ArrayBuffer,
  patches: readonly DocxTextBoxIdentityPatch[],
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
  const documents = await Promise.all(
    entries.map(async (entry) => ({
      entry,
      document: parseXml(await entry.async('text'), entry.name),
    })),
  );
  const properties = documents.flatMap(({ document }) =>
    descendants(document, 'docPr'),
  );
  const marked = properties
    .map((property) => ({
      property,
      patch: textBoxIdentityPatch(property, byMarker),
    }))
    .filter(
      (
        value,
      ): value is {
        property: Element;
        patch: DocxTextBoxIdentityPatch;
      } => Boolean(value.patch),
    );
  if (!marked.length) return buffer;

  const used = new Set<number>();
  for (const property of properties) {
    if (marked.some((value) => value.property === property)) continue;
    const id = parseDocPropertiesId(property.getAttribute('id'));
    if (id !== null) used.add(id);
  }
  let nextId = 1;
  const changedDocuments = new Set<Document>();
  for (const { property, patch } of marked) {
    removeTextBoxIdentityMarker(property, patch.marker);
    const preferred =
      patch.preferredId !== null && !used.has(patch.preferredId)
        ? patch.preferredId
        : nextAvailableDocPropertiesId(used, nextId);
    used.add(preferred);
    nextId = preferred + 1;
    property.setAttribute('id', String(preferred));
    changedDocuments.add(property.ownerDocument);
  }

  for (const { entry, document } of documents) {
    if (changedDocuments.has(document)) {
      archive.file(entry.name, new XMLSerializer().serializeToString(document));
    }
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function textBoxIdentityPatch(
  property: Element,
  byMarker: ReadonlyMap<string, DocxTextBoxIdentityPatch>,
): DocxTextBoxIdentityPatch | null {
  for (const attribute of Array.from(property.attributes)) {
    const marker = attribute.value
      .match(TEXT_BOX_ID_MARKER_PATTERN)
      ?.find((value) => byMarker.has(value));
    if (marker) return byMarker.get(marker) ?? null;
  }
  return null;
}

function removeTextBoxIdentityMarker(property: Element, marker: string): void {
  for (const attribute of Array.from(property.attributes)) {
    if (!attribute.value.includes(marker)) continue;
    attribute.value = attribute.value
      .replace(marker, '')
      .replace(/\s{2,}/g, ' ')
      .trim();
  }
}

function parseDocPropertiesId(value: string | null): number | null {
  if (value === null || !/^\d+$/.test(value)) return null;
  const id = Number(value);
  return Number.isSafeInteger(id) && id <= DOC_PROPERTIES_ID_MAX ? id : null;
}

function nextAvailableDocPropertiesId(
  used: ReadonlySet<number>,
  start: number,
): number {
  let candidate = Math.max(1, start);
  while (candidate <= DOC_PROPERTIES_ID_MAX && used.has(candidate)) {
    candidate += 1;
  }
  if (candidate > DOC_PROPERTIES_ID_MAX) {
    throw new Error('DOCX drawing-property IDs are exhausted.');
  }
  return candidate;
}
