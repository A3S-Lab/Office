import JSZip from 'jszip';
import {
  normalizeDocumentOpenTypeFeatures,
  serializeDocumentOpenTypeFeatures,
  type WorkDocumentOpenTypeFeatures,
} from './work-document-opentype';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { DOCX_WORD_2010_NAMESPACE } from './work-docx-opentype-import';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { descendants, directChildren, parseXml } from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

export interface DocxOpenTypePatch {
  marker: string;
  style?: string;
  features: WorkDocumentOpenTypeFeatures;
}

const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const OPEN_TYPE_PART_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/i;
const OPEN_TYPE_PROPERTY_ORDER = [
  'ligatures',
  'numForm',
  'numSpacing',
  'stylisticSets',
  'cntxtAlts',
] as const;

export class DocxOpenTypePatchCollector {
  readonly patches: DocxOpenTypePatch[] = [];
  private nextMarker = 1;
  private readonly occupied: string;
  private readonly patchesByMarker = new Map<string, DocxOpenTypePatch>();
  private readonly markersByFeaturesAndStyle = new Map<string, string>();

  constructor(source: string) {
    this.occupied = source;
  }

  marker(value: unknown, inheritedStyle?: string): string {
    const features = normalizeDocumentOpenTypeFeatures(value);
    if (!features) {
      throw new Error('Document contains invalid OpenType typography.');
    }
    const style = this.originalStyle(inheritedStyle);
    const serialized = serializeDocumentOpenTypeFeatures(features);
    if (!serialized) {
      throw new Error('Document contains invalid OpenType typography.');
    }
    const key = JSON.stringify([serialized, style ?? null]);
    const existing = this.markersByFeaturesAndStyle.get(key);
    if (existing) return existing;
    let marker = '';
    do {
      marker = `A3SOfficeOpenType${String(this.nextMarker).padStart(8, '0')}`;
      this.nextMarker += 1;
    } while (
      this.occupied.includes(marker) ||
      this.patches.some((patch) => patch.marker === marker)
    );
    const patch = { marker, features, ...(style ? { style } : {}) };
    this.patches.push(patch);
    this.patchesByMarker.set(marker, patch);
    this.markersByFeaturesAndStyle.set(key, marker);
    return marker;
  }

  lookup(marker: string | undefined): DocxOpenTypePatch | undefined {
    return marker ? this.patchesByMarker.get(marker) : undefined;
  }

  private originalStyle(style: string | undefined): string | undefined {
    let current = style;
    const seen = new Set<string>();
    while (current && !seen.has(current)) {
      seen.add(current);
      const patch = this.patchesByMarker.get(current);
      if (!patch) return current;
      current = patch.style;
    }
    return current;
  }
}

export async function patchDocxOpenTypeFeatures(
  buffer: ArrayBuffer,
  patches: readonly DocxOpenTypePatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  const byMarker = validatedPatches(patches);
  const archive = await JSZip.loadAsync(buffer);
  const applied = new Set<string>();
  for (const entry of Object.values(archive.files)) {
    if (entry.dir || !OPEN_TYPE_PART_PATTERN.test(entry.name)) continue;
    const document = parseXml(
      decodeXmlBytes(
        await entry.async('uint8array'),
        `generated DOCX ${entry.name}`,
      ),
      `generated DOCX ${entry.name}`,
    );
    let changed = false;
    for (const properties of descendants(document, 'rPr')) {
      if (!DOCX_WORDPROCESSING_NAMESPACES.has(properties.namespaceURI ?? '')) {
        continue;
      }
      const markerStyles = directChildren(properties, 'rStyle').filter(
        (style) => {
          if (!DOCX_WORDPROCESSING_NAMESPACES.has(style.namespaceURI ?? '')) {
            return false;
          }
          const marker = wordValue(style);
          return marker !== null && byMarker.has(marker);
        },
      );
      if (!markerStyles.length) continue;
      if (markerStyles.length !== 1) {
        throw new Error(
          'Generated DOCX OpenType marker is duplicated in one run.',
        );
      }
      const marker = wordValue(markerStyles[0]);
      const patch = marker ? byMarker.get(marker) : undefined;
      if (!marker || !patch) continue;
      if (
        directChildren(properties).some(
          (child) =>
            child.namespaceURI === DOCX_WORD_2010_NAMESPACE &&
            OPEN_TYPE_PROPERTY_ORDER.includes(
              child.localName as (typeof OPEN_TYPE_PROPERTY_ORDER)[number],
            ),
        )
      ) {
        throw new Error(
          'Generated DOCX OpenType marker unexpectedly has a native value.',
        );
      }
      if (patch.style) setWordValue(markerStyles[0], patch.style);
      else markerStyles[0].remove();
      appendDocxOpenTypeProperties(document, properties, patch.features);
      applied.add(marker);
      changed = true;
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  const missing = patches.filter((patch) => !applied.has(patch.marker));
  if (missing.length) {
    throw new Error(
      `DOCX OpenType markers were not emitted: ${missing
        .map((patch) => patch.marker)
        .join(', ')}.`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function validatedPatches(
  patches: readonly DocxOpenTypePatch[],
): Map<string, DocxOpenTypePatch> {
  const result = new Map<string, DocxOpenTypePatch>();
  for (const patch of patches) {
    const features = normalizeDocumentOpenTypeFeatures(patch.features);
    if (!patch.marker || !features || result.has(patch.marker)) {
      throw new Error('Document contains invalid OpenType patches.');
    }
    result.set(patch.marker, { ...patch, features });
  }
  return result;
}

export function appendDocxOpenTypeProperties(
  document: Document,
  properties: Element,
  source: unknown,
): void {
  const features = normalizeDocumentOpenTypeFeatures(source);
  if (!features) {
    throw new Error('Document contains invalid OpenType typography.');
  }
  const prefix = ensureWord2010Prefix(document.documentElement);
  if (features.ligatures !== undefined) {
    properties.append(
      valuedElement(document, prefix, 'ligatures', features.ligatures),
    );
  }
  if (features.numberForm !== undefined) {
    properties.append(
      valuedElement(document, prefix, 'numForm', features.numberForm),
    );
  }
  if (features.numberSpacing !== undefined) {
    properties.append(
      valuedElement(document, prefix, 'numSpacing', features.numberSpacing),
    );
  }
  if (features.stylisticSets !== undefined) {
    const stylisticSets = word2010Element(document, prefix, 'stylisticSets');
    for (const id of features.stylisticSets) {
      stylisticSets.append(
        valuedElement(document, prefix, 'styleSet', String(id), 'id'),
      );
    }
    properties.append(stylisticSets);
  }
  if (features.contextualAlternates !== undefined) {
    properties.append(
      valuedElement(
        document,
        prefix,
        'cntxtAlts',
        features.contextualAlternates ? '1' : '0',
      ),
    );
  }
}

function valuedElement(
  document: Document,
  prefix: string,
  name: string,
  value: string,
  attribute = 'val',
): Element {
  const element = word2010Element(document, prefix, name);
  element.setAttributeNS(
    DOCX_WORD_2010_NAMESPACE,
    `${prefix}:${attribute}`,
    value,
  );
  return element;
}

function word2010Element(
  document: Document,
  prefix: string,
  name: string,
): Element {
  return document.createElementNS(
    DOCX_WORD_2010_NAMESPACE,
    `${prefix}:${name}`,
  );
}

function ensureWord2010Prefix(root: Element): string {
  const existing = Array.from(root.attributes).find(
    (attribute) =>
      attribute.namespaceURI === XMLNS_NAMESPACE &&
      attribute.value === DOCX_WORD_2010_NAMESPACE,
  );
  const prefix = existing?.localName ?? availablePrefix(root, 'w14');
  if (!existing) {
    root.setAttributeNS(
      XMLNS_NAMESPACE,
      `xmlns:${prefix}`,
      DOCX_WORD_2010_NAMESPACE,
    );
  }
  ensureIgnorable(root, prefix);
  return prefix;
}

function availablePrefix(root: Element, preferred: string): string {
  let index = 0;
  while (true) {
    const prefix = index ? `${preferred}${index}` : preferred;
    const occupied = Array.from(root.attributes).some(
      (attribute) =>
        attribute.namespaceURI === XMLNS_NAMESPACE &&
        attribute.localName === prefix,
    );
    if (!occupied) return prefix;
    index += 1;
  }
}

function ensureIgnorable(root: Element, word2010Prefix: string): void {
  const existingPrefix = Array.from(root.attributes).find(
    (attribute) =>
      attribute.namespaceURI === XMLNS_NAMESPACE &&
      attribute.value === MARKUP_COMPATIBILITY_NAMESPACE,
  )?.localName;
  const prefix = existingPrefix ?? availablePrefix(root, 'mc');
  if (!existingPrefix) {
    root.setAttributeNS(
      XMLNS_NAMESPACE,
      `xmlns:${prefix}`,
      MARKUP_COMPATIBILITY_NAMESPACE,
    );
  }
  const value = root.getAttributeNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    'Ignorable',
  );
  const tokens = value?.trim().split(/\s+/).filter(Boolean) ?? [];
  if (!tokens.includes(word2010Prefix)) tokens.push(word2010Prefix);
  root.setAttributeNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    `${prefix}:Ignorable`,
    tokens.join(' '),
  );
}

function wordValue(element: Element): string | null {
  const values = Array.from(element.attributes).filter(
    (attribute) =>
      xmlAttributeLocalName(attribute) === 'val' &&
      DOCX_WORDPROCESSING_NAMESPACES.has(
        xmlAttributeNamespace(element, attribute) ?? '',
      ),
  );
  return values.length === 1 ? values[0].value : null;
}

function setWordValue(element: Element, value: string): void {
  const namespace = element.namespaceURI ?? '';
  const prefix = element.prefix || 'w';
  for (const attribute of Array.from(element.attributes)) {
    if (
      xmlAttributeLocalName(attribute) === 'val' &&
      DOCX_WORDPROCESSING_NAMESPACES.has(
        xmlAttributeNamespace(element, attribute) ?? '',
      )
    ) {
      element.removeAttributeNode(attribute);
    }
  }
  element.setAttributeNS(namespace, `${prefix}:val`, value);
}
