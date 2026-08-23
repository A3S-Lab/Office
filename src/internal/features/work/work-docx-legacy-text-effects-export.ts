import JSZip from 'jszip';
import {
  DOCUMENT_LEGACY_TEXT_EFFECT_NAMES,
  normalizeDocumentLegacyTextEffects,
  type WorkDocumentLegacyTextEffects,
} from './work-document-legacy-text-effects';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { DOCX_RUN_PROPERTY_RANK } from './work-docx-run-property-order';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { descendants, directChildren, parseXml } from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

export interface DocxLegacyTextEffectsPatch {
  marker: string;
  style?: string;
  effects: WorkDocumentLegacyTextEffects;
}

const LEGACY_TEXT_EFFECTS_PART_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/i;

export class DocxLegacyTextEffectsPatchCollector {
  readonly patches: DocxLegacyTextEffectsPatch[] = [];
  private nextMarker = 1;
  private readonly occupied: string;
  private readonly patchesByMarker = new Map<
    string,
    DocxLegacyTextEffectsPatch
  >();
  private readonly markersByEffectsAndStyle = new Map<string, string>();

  constructor(source: string) {
    this.occupied = source;
  }

  marker(
    value: WorkDocumentLegacyTextEffects,
    inheritedStyle?: string,
  ): string {
    const effects = normalizeDocumentLegacyTextEffects(value);
    if (!effects || !Object.keys(effects).length) {
      throw new Error('Document contains invalid legacy text effects.');
    }
    const style = this.originalStyle(inheritedStyle);
    const key = JSON.stringify([
      DOCUMENT_LEGACY_TEXT_EFFECT_NAMES.map((name) => effects[name] ?? null),
      style ?? null,
    ]);
    const existing = this.markersByEffectsAndStyle.get(key);
    if (existing) return existing;
    let marker = '';
    do {
      marker = `A3SOfficeLegacyTextEffects${String(this.nextMarker).padStart(
        8,
        '0',
      )}`;
      this.nextMarker += 1;
    } while (
      this.occupied.includes(marker) ||
      this.patches.some((patch) => patch.marker === marker)
    );
    const patch = { marker, effects, ...(style ? { style } : {}) };
    this.patches.push(patch);
    this.patchesByMarker.set(marker, patch);
    this.markersByEffectsAndStyle.set(key, marker);
    return marker;
  }

  lookup(marker: string | undefined): DocxLegacyTextEffectsPatch | undefined {
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

export async function patchDocxLegacyTextEffects(
  buffer: ArrayBuffer,
  patches: readonly DocxLegacyTextEffectsPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  const byMarker = new Map<string, DocxLegacyTextEffectsPatch>();
  for (const patch of patches) {
    const effects = normalizeDocumentLegacyTextEffects(patch.effects);
    if (
      !patch.marker ||
      !effects ||
      !Object.keys(effects).length ||
      byMarker.has(patch.marker)
    ) {
      throw new Error('Document contains invalid legacy text-effect patches.');
    }
    byMarker.set(patch.marker, { ...patch, effects });
  }
  const archive = await JSZip.loadAsync(buffer);
  const applied = new Set<string>();
  for (const entry of Object.values(archive.files)) {
    if (entry.dir || !LEGACY_TEXT_EFFECTS_PART_PATTERN.test(entry.name)) {
      continue;
    }
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
          'Generated DOCX legacy text-effect marker is duplicated in one run.',
        );
      }
      const marker = wordValue(markerStyles[0]);
      const patch = marker ? byMarker.get(marker) : undefined;
      if (!marker || !patch) continue;
      if (
        DOCUMENT_LEGACY_TEXT_EFFECT_NAMES.some(
          (name) => directChildren(properties, name).length > 0,
        )
      ) {
        throw new Error(
          'Generated DOCX legacy text-effect marker unexpectedly has a native value.',
        );
      }
      if (patch.style) setWordValue(markerStyles[0], patch.style);
      else markerStyles[0].remove();
      for (const name of DOCUMENT_LEGACY_TEXT_EFFECT_NAMES) {
        const value = patch.effects[name];
        if (value === undefined) continue;
        insertLegacyTextEffectProperty(document, properties, name, value);
      }
      applied.add(marker);
      changed = true;
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  const missing = patches.filter((patch) => !applied.has(patch.marker));
  if (missing.length) {
    throw new Error(
      `DOCX legacy text-effect markers were not emitted: ${missing
        .map((patch) => patch.marker)
        .join(', ')}.`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function insertLegacyTextEffectProperty(
  document: Document,
  properties: Element,
  name: (typeof DOCUMENT_LEGACY_TEXT_EFFECT_NAMES)[number],
  value: boolean,
): void {
  const namespace = properties.namespaceURI ?? '';
  const prefix = properties.prefix || 'w';
  const effect = document.createElementNS(namespace, `${prefix}:${name}`);
  if (!value) effect.setAttributeNS(namespace, `${prefix}:val`, '0');
  const rank = DOCX_RUN_PROPERTY_RANK.get(name) ?? 0;
  const next = directChildren(properties).find((child) => {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '')) {
      return true;
    }
    return (DOCX_RUN_PROPERTY_RANK.get(child.localName) ?? Infinity) > rank;
  });
  properties.insertBefore(effect, next ?? null);
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
