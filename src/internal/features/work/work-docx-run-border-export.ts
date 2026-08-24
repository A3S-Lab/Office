import type { IRunOptions } from 'docx';
import JSZip from 'jszip';
import {
  type DocumentRunBorder,
  normalizeDocumentRunBorder,
} from './work-document-run-border';
import { parseXml, xmlNamespacePrefix } from './work-ooxml-package';
import { serializeUtf8Xml } from './work-ooxml-xml';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export interface DocxRunBorderPatch {
  marker: string;
  border: DocumentRunBorder;
}

export class DocxRunBorderPatchCollector {
  readonly patches: DocxRunBorderPatch[] = [];
  private nextMarker = 1;
  private readonly usedColors: Set<string>;

  constructor(sourceHtml: string) {
    this.usedColors = sourceColors(sourceHtml);
  }

  marker(source: unknown): string | null {
    const border = normalizeDocumentRunBorder(source);
    if (!border) return null;
    let marker = '';
    do {
      const value = 0xc00000 + this.nextMarker;
      if (value > 0xcfffff) {
        throw new Error('DOCX character border marker space is exhausted.');
      }
      marker = value.toString(16).padStart(6, '0').toUpperCase();
      this.nextMarker += 1;
    } while (this.usedColors.has(marker));
    this.usedColors.add(marker);
    this.patches.push({ marker, border });
    return marker;
  }

  hasMarker(value: string | null | undefined): value is string {
    const marker = value?.toUpperCase();
    return Boolean(
      marker && this.patches.some((patch) => patch.marker === marker),
    );
  }
}

export function documentRunBorderDocxOptions(
  source: unknown,
  collector: DocxRunBorderPatchCollector,
): Pick<IRunOptions, 'border'> {
  const marker = collector.marker(source);
  return marker
    ? {
        border: {
          style: 'single',
          color: marker,
          size: 2,
          space: 0,
        },
      }
    : {};
}

export async function patchDocxRunBorders(
  buffer: ArrayBuffer,
  patches: readonly DocxRunBorderPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  const byMarker = new Map(patches.map((patch) => [patch.marker, patch]));
  const applied = new Set<string>();
  for (const entry of Object.values(archive.files)) {
    if (
      entry.dir ||
      !/^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/i.test(
        entry.name,
      )
    ) {
      continue;
    }
    const document = parseXml(await entry.async('text'), entry.name);
    let changed = false;
    for (const element of Array.from(document.getElementsByTagName('*'))) {
      if (
        element.localName !== 'bdr' ||
        element.namespaceURI !== WORD_NAMESPACE ||
        element.parentElement?.localName !== 'rPr'
      ) {
        continue;
      }
      const marker = wordAttribute(element, 'color')?.toUpperCase();
      const patch = marker ? byMarker.get(marker) : undefined;
      if (!marker || !patch) continue;
      replaceRunBorderAttributes(document, element, patch.border);
      applied.add(marker);
      changed = true;
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  const missing = patches.filter((patch) => !applied.has(patch.marker));
  if (missing.length) {
    throw new Error(
      `Generated DOCX is missing ${missing.length} character border marker(s).`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

export function setDocxRunBorderAttributes(
  document: Document,
  element: Element,
  border: DocumentRunBorder,
): void {
  setWordAttribute(document, element, 'val', border.style);
  const color = border.color;
  if (color) {
    setWordAttribute(
      document,
      element,
      'color',
      color.value === 'auto' ? 'auto' : color.value.slice(1).toUpperCase(),
    );
    if (color.theme) {
      setWordAttribute(document, element, 'themeColor', color.theme.theme);
      if (color.theme.tint) {
        setWordAttribute(document, element, 'themeTint', color.theme.tint);
      }
      if (color.theme.shade) {
        setWordAttribute(document, element, 'themeShade', color.theme.shade);
      }
    }
  }
  if (border.size !== undefined) {
    setWordAttribute(document, element, 'sz', String(border.size));
  }
  if (border.space !== undefined) {
    setWordAttribute(document, element, 'space', String(border.space));
  }
  if (border.shadow !== undefined) {
    setWordAttribute(document, element, 'shadow', border.shadow ? '1' : '0');
  }
  if (border.frame !== undefined) {
    setWordAttribute(document, element, 'frame', border.frame ? '1' : '0');
  }
}

function replaceRunBorderAttributes(
  document: Document,
  element: Element,
  border: DocumentRunBorder,
): void {
  for (const attribute of Array.from(element.attributes)) {
    if (attribute.name !== 'xmlns' && !attribute.name.startsWith('xmlns:')) {
      element.removeAttributeNode(attribute);
    }
  }
  element.replaceChildren();
  setDocxRunBorderAttributes(document, element, border);
}

function sourceColors(source: string): Set<string> {
  const colors = new Set<string>();
  for (const match of source.matchAll(/#([0-9a-f]{6})\b/gi)) {
    if (match[1]) colors.add(match[1].toUpperCase());
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
