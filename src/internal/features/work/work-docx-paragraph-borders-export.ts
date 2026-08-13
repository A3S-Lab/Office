import type { IParagraphPropertiesOptions } from 'docx';
import JSZip from 'jszip';
import {
  DOCUMENT_PARAGRAPH_BORDER_EDGES,
  type DocumentParagraphBorder,
  type DocumentParagraphBorders,
  normalizeDocumentParagraphBorders,
  parseDocumentParagraphBordersElement,
} from './work-document-paragraph-borders';
import {
  directChildren,
  parseXml,
  xmlNamespacePrefix,
} from './work-ooxml-package';
import { serializeUtf8Xml } from './work-ooxml-xml';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

type ParagraphBorderOptions = Partial<
  Pick<IParagraphPropertiesOptions, 'border'>
>;

interface DocxParagraphBorderPatch {
  marker: string;
  borders: DocumentParagraphBorders;
}

export class DocxParagraphBorderPatchCollector {
  readonly patches: DocxParagraphBorderPatch[] = [];
  private nextMarker = 1;
  private readonly usedColors: Set<string>;

  constructor(sourceHtml: string) {
    this.usedColors = sourceColors(sourceHtml);
  }

  marker(source: unknown): string | null {
    const borders = normalizeDocumentParagraphBorders(source);
    if (!borders) return null;
    let marker = '';
    do {
      const value = 0xd00000 + this.nextMarker;
      if (value > 0xffffff)
        throw new Error('DOCX paragraph border marker space is exhausted.');
      marker = value.toString(16).padStart(6, '0').toUpperCase();
      this.nextMarker += 1;
    } while (this.usedColors.has(marker));
    this.usedColors.add(marker);
    this.patches.push({ marker, borders });
    return marker;
  }
}

export function documentParagraphBordersDocxOptions(
  element: HTMLElement,
  collector?: DocxParagraphBorderPatchCollector,
): ParagraphBorderOptions {
  const borders = parseDocumentParagraphBordersElement(element);
  const marker = collector?.marker(borders);
  if (!marker) return {};
  return {
    border: {
      top: {
        style: 'single',
        color: marker,
        size: 2,
        space: 0,
      },
    },
  };
}

export async function patchDocxParagraphBorders(
  buffer: ArrayBuffer,
  patches: readonly DocxParagraphBorderPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  const byMarker = new Map(patches.map((patch) => [patch.marker, patch]));
  const applied = new Set<string>();
  const entries = Object.values(archive.files).filter(
    (entry) =>
      !entry.dir &&
      /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/i.test(
        entry.name,
      ),
  );
  for (const entry of entries) {
    const document = parseXml(await entry.async('text'), entry.name);
    let changed = false;
    for (const element of Array.from(document.getElementsByTagName('*'))) {
      if (
        element.localName !== 'pBdr' ||
        element.namespaceURI !== WORD_NAMESPACE
      ) {
        continue;
      }
      const marker = directChildren(element)
        .map((edge) => wordAttribute(edge, 'color')?.toUpperCase())
        .find((color) => color && byMarker.has(color));
      const patch = marker ? byMarker.get(marker) : undefined;
      if (!patch || !marker) continue;
      replaceParagraphBorders(document, element, patch.borders);
      applied.add(marker);
      changed = true;
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  const missing = patches.filter((patch) => !applied.has(patch.marker));
  if (missing.length) {
    throw new Error(
      `Generated DOCX is missing ${missing.length} paragraph border marker(s).`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function replaceParagraphBorders(
  document: Document,
  container: Element,
  borders: DocumentParagraphBorders,
): void {
  container.replaceChildren();
  const prefix =
    xmlNamespacePrefix(document.documentElement, WORD_NAMESPACE) ?? 'w';
  for (const edge of DOCUMENT_PARAGRAPH_BORDER_EDGES) {
    const border = borders[edge];
    if (!border) continue;
    const element = document.createElementNS(
      WORD_NAMESPACE,
      `${prefix}:${edge}`,
    );
    setWordAttribute(document, element, 'val', border.style);
    setBorderColor(document, element, border);
    if (border.size !== undefined)
      setWordAttribute(document, element, 'sz', String(border.size));
    if (border.space !== undefined)
      setWordAttribute(document, element, 'space', String(border.space));
    if (border.shadow !== undefined)
      setWordAttribute(document, element, 'shadow', border.shadow ? '1' : '0');
    if (border.frame !== undefined)
      setWordAttribute(document, element, 'frame', border.frame ? '1' : '0');
    container.append(element);
  }
}

function setBorderColor(
  document: Document,
  element: Element,
  border: DocumentParagraphBorder,
): void {
  const color = border.color;
  if (!color) return;
  setWordAttribute(
    document,
    element,
    'color',
    color.value === 'auto' ? 'auto' : color.value.slice(1).toUpperCase(),
  );
  if (!color.theme) return;
  setWordAttribute(document, element, 'themeColor', color.theme.theme);
  if (color.theme.tint)
    setWordAttribute(document, element, 'themeTint', color.theme.tint);
  if (color.theme.shade)
    setWordAttribute(document, element, 'themeShade', color.theme.shade);
}

function sourceColors(source: string): Set<string> {
  const colors = new Set<string>();
  for (const match of source.matchAll(/#([0-9a-f]{6})\b/gi)) {
    if (match[1]) colors.add(match[1].toUpperCase());
  }
  for (const match of source.matchAll(
    /rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/gi,
  )) {
    const channels = match.slice(1, 4).map(Number);
    if (channels.some((channel) => channel < 0 || channel > 255)) continue;
    colors.add(
      channels
        .map((channel) => channel.toString(16).padStart(2, '0'))
        .join('')
        .toUpperCase(),
    );
  }
  return colors;
}

function wordAttribute(element: Element, name: string): string | null {
  return (
    element.getAttributeNS(WORD_NAMESPACE, name) ??
    element.getAttribute(`w:${name}`)
  );
}

function setWordAttribute(
  document: Document,
  element: Element,
  name: string,
  value: string,
): void {
  const prefix =
    xmlNamespacePrefix(document.documentElement, WORD_NAMESPACE) ?? 'w';
  element.setAttributeNS(WORD_NAMESPACE, `${prefix}:${name}`, value);
}
