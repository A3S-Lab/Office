import JSZip from 'jszip';
import { parseXml, xmlNamespacePrefix } from './work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export interface DocxThemeColorReference {
  theme: string;
  resolved: string;
  tint?: string;
  shade?: string;
}

type DocxThemePatchKind = 'color' | 'fill' | 'border' | 'shadingColor';

interface DocxThemePatch {
  kind: DocxThemePatchKind;
  marker: string;
  reference: DocxThemeColorReference;
}

export class DocxThemePatchCollector {
  readonly patches: DocxThemePatch[] = [];
  private nextMarker = 1;
  private readonly usedColors: Set<string>;

  constructor(sourceHtml: string) {
    this.usedColors = sourceColors(sourceHtml);
  }

  marker(
    kind: DocxThemePatchKind,
    reference: DocxThemeColorReference | null,
    currentColor: string | null,
  ): string | null {
    if (!reference || normalizeColor(currentColor) !== reference.resolved) {
      return null;
    }
    let marker = '';
    do {
      marker = (0xf00000 + this.nextMarker)
        .toString(16)
        .padStart(6, '0')
        .toUpperCase();
      this.nextMarker += 1;
    } while (this.usedColors.has(marker));
    this.usedColors.add(marker);
    this.patches.push({ kind, marker, reference });
    return marker;
  }
}

export function serializeDocxThemeReference(
  reference: DocxThemeColorReference | null,
): string | undefined {
  return reference ? JSON.stringify(reference) : undefined;
}

export function parseDocxThemeReference(
  value: string | null | undefined,
): DocxThemeColorReference | null {
  if (!value) return null;
  try {
    const parsed = JSON.parse(value) as Record<string, unknown>;
    const theme = typeof parsed.theme === 'string' ? parsed.theme.trim() : '';
    const resolved = normalizeColor(
      typeof parsed.resolved === 'string' ? parsed.resolved : null,
    );
    const tint = byteHex(parsed.tint);
    const shade = byteHex(parsed.shade);
    if (!theme || !resolved) return null;
    return {
      theme,
      resolved,
      ...(tint ? { tint } : {}),
      ...(shade ? { shade } : {}),
    };
  } catch {
    return null;
  }
}

export async function patchDocxThemeReferences(
  buffer: ArrayBuffer,
  patches: readonly DocxThemePatch[],
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
    for (const element of Array.from(document.getElementsByTagName('*'))) {
      for (const target of themePatchTargets(element.localName)) {
        const marker = wordAttribute(
          element,
          target.directAttribute,
        )?.toUpperCase();
        const patch = marker ? byMarker.get(marker) : undefined;
        if (!patch || patch.kind !== target.kind) continue;
        setWordAttribute(
          document,
          element,
          target.directAttribute,
          patch.reference.resolved.slice(1).toUpperCase(),
        );
        setWordAttribute(
          document,
          element,
          target.themeAttribute,
          patch.reference.theme,
        );
        setOptionalWordAttribute(
          document,
          element,
          target.tintAttribute,
          patch.reference.tint,
        );
        setOptionalWordAttribute(
          document,
          element,
          target.shadeAttribute,
          patch.reference.shade,
        );
        changed = true;
      }
    }
    if (changed)
      archive.file(entry.name, new XMLSerializer().serializeToString(document));
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function themePatchTargets(localName: string): Array<{
  kind: DocxThemePatchKind;
  directAttribute: string;
  themeAttribute: string;
  tintAttribute: string;
  shadeAttribute: string;
}> {
  if (localName === 'color') {
    return [
      {
        kind: 'color',
        directAttribute: 'val',
        themeAttribute: 'themeColor',
        tintAttribute: 'themeTint',
        shadeAttribute: 'themeShade',
      },
    ];
  }
  if (localName === 'shd') {
    return [
      {
        kind: 'fill',
        directAttribute: 'fill',
        themeAttribute: 'themeFill',
        tintAttribute: 'themeFillTint',
        shadeAttribute: 'themeFillShade',
      },
      {
        kind: 'shadingColor',
        directAttribute: 'color',
        themeAttribute: 'themeColor',
        tintAttribute: 'themeTint',
        shadeAttribute: 'themeShade',
      },
    ];
  }
  return ['top', 'right', 'bottom', 'left', 'start', 'end'].includes(localName)
    ? [
        {
          kind: 'border',
          directAttribute: 'color',
          themeAttribute: 'themeColor',
          tintAttribute: 'themeTint',
          shadeAttribute: 'themeShade',
        },
      ]
    : [];
}

function normalizeColor(value: string | null): string | null {
  const normalized = value?.trim().toLowerCase();
  if (!normalized || !/^#[0-9a-f]{6}$/.test(normalized)) return null;
  return normalized;
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

function byteHex(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  const normalized = value.trim().toUpperCase();
  return /^[0-9A-F]{2}$/.test(normalized) ? normalized : undefined;
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

function setOptionalWordAttribute(
  document: Document,
  element: Element,
  name: string,
  value: string | undefined,
): void {
  if (value) setWordAttribute(document, element, name, value);
  else element.removeAttributeNS(WORD_NAMESPACE, name);
}
