import JSZip from 'jszip';
import {
  DOCUMENT_CHARACTER_SPACING_MAX_TWIPS,
  normalizeDocumentCharacterSpacingTwips,
} from './work-document-character-spacing';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { descendants, directChildren, parseXml } from './work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

export const DOCX_EXPLICIT_ZERO_CHARACTER_SPACING_SENTINEL =
  DOCUMENT_CHARACTER_SPACING_MAX_TWIPS + 1;

const CHARACTER_SPACING_PART_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i;

export function docxCharacterSpacingValue(value: unknown): number | undefined {
  const spacing = normalizeDocumentCharacterSpacingTwips(value);
  if (spacing === null) return undefined;
  return spacing === 0
    ? DOCX_EXPLICIT_ZERO_CHARACTER_SPACING_SENTINEL
    : spacing;
}

export function docxCharacterSpacingTwipsFromProperties(
  properties: Element | null | undefined,
): number | undefined {
  if (!properties) return undefined;
  const spacing = directChildren(properties, 'spacing').filter((element) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  if (spacing.length !== 1) return undefined;
  return (
    normalizeDocumentCharacterSpacingTwips(wordValue(spacing[0])) ?? undefined
  );
}

export async function patchDocxExplicitZeroCharacterSpacing(
  buffer: ArrayBuffer,
): Promise<ArrayBuffer> {
  const archive = await JSZip.loadAsync(buffer);
  let replacements = 0;
  for (const entry of Object.values(archive.files)) {
    if (entry.dir || !CHARACTER_SPACING_PART_PATTERN.test(entry.name)) {
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
    for (const spacing of descendants(document, 'spacing')) {
      if (
        !DOCX_WORDPROCESSING_NAMESPACES.has(spacing.namespaceURI ?? '') ||
        wordValue(spacing) !==
          String(DOCX_EXPLICIT_ZERO_CHARACTER_SPACING_SENTINEL)
      ) {
        continue;
      }
      setWordValue(spacing, '0');
      replacements += 1;
      changed = true;
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  if (!replacements) {
    throw new Error(
      'Generated DOCX explicit-zero character spacing was not emitted.',
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function wordValue(element: Element): string | null {
  const values = Array.from(element.attributes).filter(
    (item) =>
      xmlAttributeLocalName(item) === 'val' &&
      DOCX_WORDPROCESSING_NAMESPACES.has(
        xmlAttributeNamespace(element, item) ?? '',
      ),
  );
  return values.length === 1 ? values[0].value : null;
}

function setWordValue(element: Element, value: string): void {
  const existing = Array.from(element.attributes).find(
    (candidate) =>
      candidate.localName === 'val' &&
      DOCX_WORDPROCESSING_NAMESPACES.has(candidate.namespaceURI ?? ''),
  );
  const namespace = element.namespaceURI ?? '';
  const prefix = element.prefix || 'w';
  if (existing) {
    element.setAttributeNS(existing.namespaceURI, existing.name, value);
  } else {
    element.setAttributeNS(namespace, `${prefix}:val`, value);
  }
}
