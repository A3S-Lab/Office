import JSZip from 'jszip';
import {
  documentFontNameFromCssFamily,
  documentScriptFontFamily,
  documentScriptFontSlotFromHint,
  normalizeDocumentScriptFontSlot,
  normalizeDocumentScriptFonts,
  serializeDocumentScriptFonts,
  type WorkDocumentScriptFonts,
} from './work-document-script-fonts';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { descendants, parseXml } from './work-ooxml-package';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

export interface DocxRunFontsPatch {
  marker: string;
  fonts: WorkDocumentScriptFonts;
}

const RUN_FONT_PART_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/i;
const WORD_FONT_ATTRIBUTE_NAMES = new Set([
  'ascii',
  'hAnsi',
  'eastAsia',
  'cs',
  'asciiTheme',
  'hAnsiTheme',
  'eastAsiaTheme',
  'cstheme',
  'hint',
]);

export class DocxRunFontsPatchCollector {
  readonly patches: DocxRunFontsPatch[] = [];
  private nextMarker = 1;
  private readonly occupied: string;
  private readonly markersByFonts = new Map<string, string>();

  constructor(source: string) {
    this.occupied = source;
  }

  marker(
    source: unknown,
    slot: unknown,
    currentFontFamily: string,
  ): string | null {
    const fonts = normalizeDocumentScriptFonts(source);
    const normalizedSlot =
      normalizeDocumentScriptFontSlot(slot) ??
      documentScriptFontSlotFromHint(fonts?.hint);
    const expected = fonts
      ? documentScriptFontFamily(fonts, normalizedSlot)
      : undefined;
    const current = documentFontNameFromCssFamily(currentFontFamily);
    if (
      !fonts ||
      (expected
        ? documentFontNameFromCssFamily(expected)?.toLocaleLowerCase() !==
          current?.toLocaleLowerCase()
        : Boolean(currentFontFamily.trim()))
    ) {
      return null;
    }
    const key = serializeDocumentScriptFonts(fonts);
    if (!key) return null;
    const existing = this.markersByFonts.get(key);
    if (existing) return existing;
    let marker = '';
    do {
      marker = `A3SOfficeRunFonts${String(this.nextMarker).padStart(8, '0')}`;
      this.nextMarker += 1;
    } while (
      this.occupied.includes(marker) ||
      this.patches.some((patch) => patch.marker === marker)
    );
    this.patches.push({ marker, fonts });
    this.markersByFonts.set(key, marker);
    return marker;
  }
}

export async function patchDocxRunFonts(
  buffer: ArrayBuffer,
  patches: readonly DocxRunFontsPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  const byMarker = new Map(patches.map((patch) => [patch.marker, patch.fonts]));
  const applied = new Set<string>();
  for (const entry of Object.values(archive.files)) {
    if (entry.dir || !RUN_FONT_PART_PATTERN.test(entry.name)) continue;
    const document = parseXml(
      decodeXmlBytes(
        await entry.async('uint8array'),
        `generated DOCX ${entry.name}`,
      ),
      `generated DOCX ${entry.name}`,
    );
    let changedPart = false;
    for (const element of descendants(document, 'rFonts')) {
      if (!DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')) {
        continue;
      }
      const marker = wordFontAttribute(element, 'ascii');
      if (!marker) continue;
      const fonts = byMarker.get(marker);
      if (!fonts) continue;
      replaceWordFontAttributes(element, fonts);
      applied.add(marker);
      changedPart = true;
    }
    if (changedPart) {
      archive.file(entry.name, serializeUtf8Xml(document));
    }
  }
  const missing = patches.filter((patch) => !applied.has(patch.marker));
  if (missing.length) {
    throw new Error(
      `DOCX run-font markers were not emitted: ${missing
        .map((patch) => patch.marker)
        .join(', ')}.`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function replaceWordFontAttributes(
  element: Element,
  fonts: WorkDocumentScriptFonts,
): void {
  for (const item of Array.from(element.attributes)) {
    if (
      item.namespaceURI !== XMLNS_NAMESPACE &&
      WORD_FONT_ATTRIBUTE_NAMES.has(xmlAttributeLocalName(item)) &&
      DOCX_WORDPROCESSING_NAMESPACES.has(
        xmlAttributeNamespace(element, item) ?? '',
      )
    ) {
      element.removeAttributeNode(item);
    }
  }
  const values = [
    ['ascii', fonts.ascii?.name],
    ['hAnsi', fonts.highAnsi?.name],
    ['eastAsia', fonts.eastAsia?.name],
    ['cs', fonts.complexScript?.name],
    ['asciiTheme', fonts.ascii?.theme],
    ['hAnsiTheme', fonts.highAnsi?.theme],
    ['eastAsiaTheme', fonts.eastAsia?.theme],
    ['cstheme', fonts.complexScript?.theme],
    ['hint', fonts.hint],
  ] as const;
  for (const [name, value] of values) {
    if (value) setWordFontAttribute(element, name, value);
  }
}

function wordFontAttribute(element: Element, name: string): string | null {
  const values = Array.from(element.attributes).filter(
    (item) =>
      xmlAttributeLocalName(item) === name &&
      DOCX_WORDPROCESSING_NAMESPACES.has(
        xmlAttributeNamespace(element, item) ?? '',
      ),
  );
  return values.length === 1 ? (values[0]?.value ?? null) : null;
}

function setWordFontAttribute(
  element: Element,
  name: string,
  value: string,
): void {
  const namespace = element.namespaceURI ?? '';
  const prefix = element.prefix || 'w';
  element.setAttributeNS(namespace, `${prefix}:${name}`, value);
}
