import JSZip from 'jszip';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { descendants, directChildren, parseXml } from './work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

export interface DocxHiddenTextPatch {
  marker: string;
  style?: string;
  value: boolean;
}

const HIDDEN_TEXT_PART_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/i;
const WORD_RUN_PROPERTY_ORDER = [
  'rStyle',
  'rFonts',
  'b',
  'bCs',
  'i',
  'iCs',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'noProof',
  'snapToGrid',
  'vanish',
  'webHidden',
  'color',
  'spacing',
  'w',
  'kern',
  'position',
  'sz',
  'szCs',
  'highlight',
  'u',
  'effect',
  'bdr',
  'shd',
  'fitText',
  'vertAlign',
  'rtl',
  'cs',
  'em',
  'lang',
  'eastAsianLayout',
  'specVanish',
  'rPrChange',
] as const;
const WORD_RUN_PROPERTY_RANK = new Map<string, number>(
  WORD_RUN_PROPERTY_ORDER.map((name, index) => [name, index]),
);
const VANISH_RANK = WORD_RUN_PROPERTY_RANK.get('vanish') ?? 0;

export class DocxHiddenTextPatchCollector {
  readonly patches: DocxHiddenTextPatch[] = [];
  private nextMarker = 1;
  private readonly occupied: string;
  private readonly patchesByMarker = new Map<string, DocxHiddenTextPatch>();
  private readonly markersByValueAndStyle = new Map<string, string>();

  constructor(source: string) {
    this.occupied = source;
  }

  marker(value: boolean, inheritedStyle?: string): string {
    const style = this.originalStyle(inheritedStyle);
    const key = JSON.stringify([value, style ?? null]);
    const existing = this.markersByValueAndStyle.get(key);
    if (existing) return existing;
    let marker = '';
    do {
      marker = `A3SOfficeHiddenText${value ? 'On' : 'Off'}${String(
        this.nextMarker,
      ).padStart(8, '0')}`;
      this.nextMarker += 1;
    } while (
      this.occupied.includes(marker) ||
      this.patches.some((patch) => patch.marker === marker)
    );
    const patch = { marker, value, ...(style ? { style } : {}) };
    this.patches.push(patch);
    this.patchesByMarker.set(marker, patch);
    this.markersByValueAndStyle.set(key, marker);
    return marker;
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

export async function patchDocxHiddenText(
  buffer: ArrayBuffer,
  patches: readonly DocxHiddenTextPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  const byMarker = new Map(patches.map((patch) => [patch.marker, patch]));
  const applied = new Set<string>();
  for (const entry of Object.values(archive.files)) {
    if (entry.dir || !HIDDEN_TEXT_PART_PATTERN.test(entry.name)) continue;
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
          'Generated DOCX hidden-text marker is duplicated in one run.',
        );
      }
      const marker = wordValue(markerStyles[0]);
      const patch = marker ? byMarker.get(marker) : undefined;
      if (!marker || !patch) continue;
      const existing = directChildren(properties, 'vanish').filter((element) =>
        DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
      );
      if (existing.length) {
        throw new Error(
          'Generated DOCX hidden-text marker unexpectedly has a native value.',
        );
      }
      if (patch.style) setWordValue(markerStyles[0], patch.style);
      else markerStyles[0].remove();
      insertHiddenTextProperty(document, properties, patch.value);
      applied.add(marker);
      changed = true;
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  const missing = patches.filter((patch) => !applied.has(patch.marker));
  if (missing.length) {
    throw new Error(
      `DOCX hidden-text markers were not emitted: ${missing
        .map((patch) => patch.marker)
        .join(', ')}.`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function insertHiddenTextProperty(
  document: Document,
  properties: Element,
  value: boolean,
): void {
  const namespace = properties.namespaceURI ?? '';
  const prefix = properties.prefix || 'w';
  const vanish = document.createElementNS(namespace, `${prefix}:vanish`);
  if (!value) vanish.setAttributeNS(namespace, `${prefix}:val`, '0');
  const next = directChildren(properties).find((child) => {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '')) {
      return true;
    }
    return (
      (WORD_RUN_PROPERTY_RANK.get(child.localName) ?? Infinity) > VANISH_RANK
    );
  });
  properties.insertBefore(vanish, next ?? null);
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
