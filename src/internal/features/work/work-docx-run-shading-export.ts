import type { IRunOptions } from 'docx';
import JSZip from 'jszip';
import {
  type DocumentRunShading,
  normalizeDocumentRunShading,
} from './work-document-run-shading';
import { parseXml, xmlNamespacePrefix } from './work-ooxml-package';
import { serializeUtf8Xml } from './work-ooxml-xml';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export interface DocxRunShadingPatch {
  marker: string;
  shading: DocumentRunShading;
}

export class DocxRunShadingPatchCollector {
  readonly patches: DocxRunShadingPatch[] = [];
  private nextMarker = 1;
  private readonly usedColors: Set<string>;

  constructor(sourceHtml: string) {
    this.usedColors = sourceColors(sourceHtml);
  }

  marker(source: unknown): string | null {
    const shading = normalizeDocumentRunShading(source);
    if (!shading) return null;
    let marker = '';
    do {
      const value = 0xd00000 + this.nextMarker;
      if (value > 0xdfffff) {
        throw new Error('DOCX character shading marker space is exhausted.');
      }
      marker = value.toString(16).padStart(6, '0').toUpperCase();
      this.nextMarker += 1;
    } while (this.usedColors.has(marker));
    this.usedColors.add(marker);
    this.patches.push({ marker, shading });
    return marker;
  }

  hasMarker(value: string | null | undefined): value is string {
    const marker = value?.toUpperCase();
    return Boolean(
      marker && this.patches.some((patch) => patch.marker === marker),
    );
  }
}

export function documentRunShadingDocxOptions(
  source: unknown,
  collector: DocxRunShadingPatchCollector,
): Pick<IRunOptions, 'shading'> {
  const marker = collector.marker(source);
  return marker
    ? {
        shading: {
          type: 'clear',
          color: 'auto',
          fill: marker,
        },
      }
    : {};
}

export async function patchDocxRunShading(
  buffer: ArrayBuffer,
  patches: readonly DocxRunShadingPatch[],
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
        element.localName !== 'shd' ||
        element.namespaceURI !== WORD_NAMESPACE ||
        element.parentElement?.localName !== 'rPr'
      ) {
        continue;
      }
      const marker = wordAttribute(element, 'fill')?.toUpperCase();
      const patch = marker ? byMarker.get(marker) : undefined;
      if (!marker || !patch) continue;
      replaceRunShadingAttributes(document, element, patch.shading);
      applied.add(marker);
      changed = true;
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  const missing = patches.filter((patch) => !applied.has(patch.marker));
  if (missing.length) {
    throw new Error(
      `Generated DOCX is missing ${missing.length} character shading marker(s).`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

export function setDocxRunShadingAttributes(
  document: Document,
  element: Element,
  source: unknown,
): boolean {
  const shading = normalizeDocumentRunShading(source);
  if (!shading) return false;
  setWordAttribute(document, element, 'val', shading.pattern);
  setShadingColorAttributes(document, element, shading.color, 'color');
  setShadingColorAttributes(document, element, shading.fill, 'fill');
  return true;
}

function replaceRunShadingAttributes(
  document: Document,
  element: Element,
  shading: DocumentRunShading,
): void {
  for (const attribute of Array.from(element.attributes)) {
    if (attribute.name !== 'xmlns' && !attribute.name.startsWith('xmlns:')) {
      element.removeAttributeNode(attribute);
    }
  }
  element.replaceChildren();
  setDocxRunShadingAttributes(document, element, shading);
}

function setShadingColorAttributes(
  document: Document,
  element: Element,
  color: DocumentRunShading['color'],
  kind: 'color' | 'fill',
): void {
  if (!color) return;
  setWordAttribute(
    document,
    element,
    kind,
    color.value === 'auto' ? 'auto' : color.value.slice(1).toUpperCase(),
  );
  if (!color.theme) return;
  const themeName = kind === 'color' ? 'themeColor' : 'themeFill';
  const tintName = kind === 'color' ? 'themeTint' : 'themeFillTint';
  const shadeName = kind === 'color' ? 'themeShade' : 'themeFillShade';
  setWordAttribute(document, element, themeName, color.theme.theme);
  if (color.theme.tint) {
    setWordAttribute(document, element, tintName, color.theme.tint);
  }
  if (color.theme.shade) {
    setWordAttribute(document, element, shadeName, color.theme.shade);
  }
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
    element.getAttributeNS(element.namespaceURI ?? WORD_NAMESPACE, name) ??
    element.getAttribute(`${element.prefix ?? 'w'}:${name}`)
  );
}

function setWordAttribute(
  document: Document,
  element: Element,
  name: string,
  value: string,
): void {
  const namespace = element.namespaceURI ?? WORD_NAMESPACE;
  const prefix =
    element.prefix ??
    xmlNamespacePrefix(document.documentElement, namespace) ??
    'w';
  element.setAttributeNS(namespace, `${prefix}:${name}`, value);
}
