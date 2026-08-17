import JSZip from 'jszip';
import {
  DOCUMENT_PARAGRAPH_BORDER_EDGES,
  type DocumentParagraphBorders,
} from './work-document-paragraph-borders';
import type { DocumentParagraphShading } from './work-document-paragraph-shading';
import type { DocumentTabStop } from './work-document-tab-stops';
import { parseDocumentParagraphFormatting } from './work-document-paragraph-format-changes';
import { setDocxBorderAttributes } from './work-docx-paragraph-borders-export';
import {
  descendants,
  directChildren,
  parseXml,
  xmlNamespacePrefix,
} from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

interface DocxParagraphFormattingChangePatch {
  marker: string;
  id: number;
  author: string;
  date: string;
  before: string;
}

interface ParagraphFormattingSnapshot {
  textAlign: string | null;
  paragraphDirection: string | null;
  indentLevel: number;
  rightIndent: number;
  firstLineIndent: number;
  spaceBefore: number | null;
  spaceAfter: number | null;
  lineHeight: string | null;
  lineRule: string | null;
  keepLines: boolean | null;
  keepWithNext: boolean | null;
  pageBreakBefore: boolean | null;
  widowControl: boolean | null;
  contextualSpacing: boolean | null;
  outlineLevel: number | null;
  tabStops: DocumentTabStop[] | null;
  paragraphBorders: DocumentParagraphBorders | null;
  paragraphShading: DocumentParagraphShading | null;
  defaultCollapsed: boolean | null;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WORD_2012_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2012/wordml';
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';
const MAX_PARAGRAPH_FORMATTING_CHANGE_PATCHES = 65_536;
const PARAGRAPH_PART_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i;

export class DocxParagraphFormattingChangePatchCollector {
  readonly patches: DocxParagraphFormattingChangePatch[] = [];

  register(element: HTMLElement, id: number): string | null {
    if (
      element.dataset.changeKind !== 'paragraph-formatting' ||
      !element.hasAttribute('data-document-change')
    ) {
      return null;
    }
    const key = element.dataset.changeId?.trim() ?? '';
    const author = element.dataset.changeAuthor?.trim() ?? '';
    const date = normalizedRevisionDate(element.dataset.changeDate);
    const before = element.dataset.changeBefore ?? '';
    if (
      !key ||
      !author ||
      author.length > 255 ||
      !parseDocumentParagraphFormatting(before)
    ) {
      throw new Error(
        'Document contains an invalid paragraph-formatting revision.',
      );
    }
    if (this.patches.length >= MAX_PARAGRAPH_FORMATTING_CHANGE_PATCHES) {
      throw new Error(
        'Document exceeds the paragraph-formatting revision limit.',
      );
    }
    const marker = `__A3S_WORK_PARAGRAPH_FORMAT_CHANGE_EXPORT_${this.patches.length + 1}__`;
    this.patches.push({ marker, id, author, date, before });
    return marker;
  }
}

export async function patchDocxParagraphFormattingChanges(
  buffer: ArrayBuffer,
  patches: readonly DocxParagraphFormattingChangePatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  if (patches.length > MAX_PARAGRAPH_FORMATTING_CHANGE_PATCHES) {
    throw new Error(
      'Document exceeds the paragraph-formatting revision limit.',
    );
  }
  const archive = await JSZip.loadAsync(buffer);
  const byMarker = new Map(patches.map((patch) => [patch.marker, patch]));
  const applied = new Set<string>();
  for (const entry of Object.values(archive.files)) {
    if (entry.dir || !PARAGRAPH_PART_PATTERN.test(entry.name)) continue;
    const document = parseXml(
      decodeXmlBytes(
        await entry.async('uint8array'),
        `generated DOCX ${entry.name}`,
      ),
      `generated DOCX ${entry.name}`,
    );
    let changed = false;
    for (const paragraph of descendants(document, 'p').filter(
      (element) => element.namespaceURI === WORD_NAMESPACE,
    )) {
      const marker = paragraphFormattingMarker(paragraph, byMarker);
      if (!marker) continue;
      setParagraphFormattingChange(document, paragraph, marker.patch);
      marker.run.remove();
      applied.add(marker.patch.marker);
      changed = true;
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  const missing = patches.filter((patch) => !applied.has(patch.marker));
  if (missing.length) {
    throw new Error(
      `DOCX paragraph-formatting revision markers were not emitted: ${missing
        .map((patch) => patch.marker)
        .join(', ')}.`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function paragraphFormattingMarker(
  paragraph: Element,
  patches: ReadonlyMap<string, DocxParagraphFormattingChangePatch>,
): { patch: DocxParagraphFormattingChangePatch; run: Element } | null {
  const matches = directChildren(paragraph, 'r').flatMap((run) => {
    if (run.namespaceURI !== WORD_NAMESPACE) return [];
    const texts = directChildren(run, 't').filter(
      (text) => text.namespaceURI === WORD_NAMESPACE,
    );
    if (texts.length !== 1 || !runHasOnlyMarkerText(run, texts[0])) return [];
    const patch = patches.get(texts[0]?.textContent ?? '');
    return patch ? [{ patch, run }] : [];
  });
  if (matches.length > 1) {
    throw new Error('Generated DOCX paragraph has duplicate revision markers.');
  }
  return matches[0] ?? null;
}

function runHasOnlyMarkerText(run: Element, markerText: Element | undefined) {
  return Boolean(
    markerText &&
      directChildren(run).every(
        (child) =>
          child === markerText ||
          (child.localName === 'rPr' && child.namespaceURI === WORD_NAMESPACE),
      ),
  );
}

function setParagraphFormattingChange(
  document: Document,
  paragraph: Element,
  patch: DocxParagraphFormattingChangePatch,
): void {
  const propertyNodes = directChildren(paragraph, 'pPr').filter(
    (element) => element.namespaceURI === WORD_NAMESPACE,
  );
  if (propertyNodes.length > 1) {
    throw new Error('Generated DOCX paragraph contains duplicate properties.');
  }
  const prefix =
    xmlNamespacePrefix(document.documentElement, WORD_NAMESPACE) ?? 'w';
  const properties =
    propertyNodes[0] ?? insertParagraphProperties(document, paragraph, prefix);
  if (
    directChildren(properties, 'pPrChange').some(
      (element) => element.namespaceURI === WORD_NAMESPACE,
    )
  ) {
    throw new Error(
      'Generated DOCX paragraph already contains a formatting revision.',
    );
  }
  const change = wordElement(document, prefix, 'pPrChange');
  setWordAttribute(document, change, 'id', String(patch.id));
  setWordAttribute(document, change, 'author', patch.author);
  setWordAttribute(document, change, 'date', patch.date);
  const before = wordElement(document, prefix, 'pPr');
  appendParagraphFormattingProperties(document, before, patch.before, prefix);
  change.append(before);
  properties.append(change);
}

function appendParagraphFormattingProperties(
  document: Document,
  properties: Element,
  serialized: string,
  prefix: string,
): void {
  const parsed = parseDocumentParagraphFormatting(serialized);
  if (!parsed) {
    throw new Error(
      'Document contains an invalid paragraph-formatting revision.',
    );
  }
  const formatting = parsed as unknown as ParagraphFormattingSnapshot;
  appendOnOff(
    document,
    properties,
    prefix,
    'keepNext',
    formatting.keepWithNext,
  );
  appendOnOff(document, properties, prefix, 'keepLines', formatting.keepLines);
  appendOnOff(
    document,
    properties,
    prefix,
    'pageBreakBefore',
    formatting.pageBreakBefore,
  );
  appendOnOff(
    document,
    properties,
    prefix,
    'widowControl',
    formatting.widowControl,
  );
  appendBorders(document, properties, prefix, formatting.paragraphBorders);
  appendShading(document, properties, prefix, formatting.paragraphShading);
  appendTabStops(document, properties, prefix, formatting.tabStops);
  if (formatting.paragraphDirection) {
    appendOnOff(
      document,
      properties,
      prefix,
      'bidi',
      formatting.paragraphDirection === 'rtl',
    );
  }
  appendSpacing(document, properties, prefix, formatting);
  appendIndent(document, properties, prefix, formatting);
  appendOnOff(
    document,
    properties,
    prefix,
    'contextualSpacing',
    formatting.contextualSpacing,
  );
  if (formatting.textAlign) {
    appendValuedProperty(
      document,
      properties,
      prefix,
      'jc',
      formatting.textAlign === 'justify' ? 'both' : formatting.textAlign,
    );
  }
  if (formatting.outlineLevel !== null) {
    appendValuedProperty(
      document,
      properties,
      prefix,
      'outlineLvl',
      String(formatting.outlineLevel),
    );
  }
  appendDefaultCollapsed(
    document,
    properties,
    prefix,
    formatting.defaultCollapsed,
  );
}

function appendSpacing(
  document: Document,
  properties: Element,
  prefix: string,
  formatting: ParagraphFormattingSnapshot,
): void {
  const spacing = wordElement(document, prefix, 'spacing');
  if (formatting.spaceBefore !== null) {
    setWordAttribute(
      document,
      spacing,
      'before',
      String(Math.round(formatting.spaceBefore * 20)),
    );
  }
  if (formatting.spaceAfter !== null) {
    setWordAttribute(
      document,
      spacing,
      'after',
      String(Math.round(formatting.spaceAfter * 20)),
    );
  }
  if (formatting.lineHeight && formatting.lineRule) {
    const line = docxLineHeight(formatting.lineHeight, formatting.lineRule);
    if (line !== null) {
      setWordAttribute(document, spacing, 'line', String(line));
      setWordAttribute(document, spacing, 'lineRule', formatting.lineRule);
    }
  }
  if (spacing.attributes.length) properties.append(spacing);
}

function docxLineHeight(value: string, rule: string): number | null {
  if (rule === 'auto') {
    const multiple = Number(value.replace(/%$/, ''));
    if (!Number.isFinite(multiple)) return null;
    return Math.round((value.endsWith('%') ? multiple / 100 : multiple) * 240);
  }
  const length = /^(\d+(?:\.\d+)?)(px|pt)?$/i.exec(value);
  if (!length?.[1]) return null;
  const points =
    length[2]?.toLowerCase() === 'px'
      ? Number(length[1]) * 0.75
      : Number(length[1]);
  return Math.round(points * 20);
}

function appendIndent(
  document: Document,
  properties: Element,
  prefix: string,
  formatting: ParagraphFormattingSnapshot,
): void {
  const indent = wordElement(document, prefix, 'ind');
  setWordAttribute(
    document,
    indent,
    'left',
    String(Math.round(formatting.indentLevel * 24 * 15)),
  );
  setWordAttribute(
    document,
    indent,
    'right',
    String(Math.round(formatting.rightIndent * 15)),
  );
  if (formatting.firstLineIndent < 0) {
    setWordAttribute(
      document,
      indent,
      'hanging',
      String(Math.round(Math.abs(formatting.firstLineIndent) * 15)),
    );
  } else {
    setWordAttribute(
      document,
      indent,
      'firstLine',
      String(Math.round(formatting.firstLineIndent * 15)),
    );
  }
  properties.append(indent);
}

function appendTabStops(
  document: Document,
  properties: Element,
  prefix: string,
  tabStops: DocumentTabStop[] | null,
): void {
  if (!tabStops?.length) return;
  const tabs = wordElement(document, prefix, 'tabs');
  for (const stop of tabStops) {
    const tab = wordElement(document, prefix, 'tab');
    setWordAttribute(document, tab, 'val', stop.alignment);
    setWordAttribute(
      document,
      tab,
      'pos',
      String(Math.round(stop.position * 15)),
    );
    if (stop.leader !== 'none') {
      setWordAttribute(document, tab, 'leader', stop.leader);
    }
    tabs.append(tab);
  }
  properties.append(tabs);
}

function appendBorders(
  document: Document,
  properties: Element,
  prefix: string,
  borders: DocumentParagraphBorders | null,
): void {
  if (!borders) return;
  const container = wordElement(document, prefix, 'pBdr');
  for (const edge of DOCUMENT_PARAGRAPH_BORDER_EDGES) {
    const border = borders[edge];
    if (!border) continue;
    const element = wordElement(document, prefix, edge);
    setDocxBorderAttributes(document, element, border);
    container.append(element);
  }
  if (container.children.length) properties.append(container);
}

function appendShading(
  document: Document,
  properties: Element,
  prefix: string,
  shading: DocumentParagraphShading | null,
): void {
  if (!shading) return;
  const element = wordElement(document, prefix, 'shd');
  setWordAttribute(document, element, 'val', shading.pattern);
  appendShadingColor(document, element, 'color', shading.color);
  appendShadingColor(document, element, 'fill', shading.fill);
  properties.append(element);
}

function appendShadingColor(
  document: Document,
  element: Element,
  kind: 'color' | 'fill',
  color: DocumentParagraphShading['color'],
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
  setWordAttribute(document, element, themeName, color.theme.theme);
  if (color.theme.tint) {
    setWordAttribute(document, element, `${themeName}Tint`, color.theme.tint);
  }
  if (color.theme.shade) {
    setWordAttribute(document, element, `${themeName}Shade`, color.theme.shade);
  }
}

function appendDefaultCollapsed(
  document: Document,
  properties: Element,
  wordPrefix: string,
  value: boolean | null,
): void {
  if (value === null) return;
  const element = document.createElementNS(
    WORD_2012_NAMESPACE,
    'w15:collapsed',
  );
  element.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:w15', WORD_2012_NAMESPACE);
  element.setAttributeNS(
    WORD_NAMESPACE,
    `${wordPrefix}:val`,
    value ? '1' : '0',
  );
  properties.append(element);
}

function appendOnOff(
  document: Document,
  properties: Element,
  prefix: string,
  name: string,
  value: boolean | null,
): void {
  if (value === null) return;
  appendValuedProperty(document, properties, prefix, name, value ? '1' : '0');
}

function appendValuedProperty(
  document: Document,
  properties: Element,
  prefix: string,
  name: string,
  value: string,
): void {
  const element = wordElement(document, prefix, name);
  setWordAttribute(document, element, 'val', value);
  properties.append(element);
}

function insertParagraphProperties(
  document: Document,
  paragraph: Element,
  prefix: string,
): Element {
  const properties = wordElement(document, prefix, 'pPr');
  paragraph.insertBefore(properties, paragraph.firstChild);
  return properties;
}

function wordElement(
  document: Document,
  prefix: string,
  name: string,
): Element {
  return document.createElementNS(WORD_NAMESPACE, `${prefix}:${name}`);
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

function normalizedRevisionDate(value: string | undefined): string {
  const time = Date.parse(value?.trim() ?? '');
  return Number.isFinite(time)
    ? new Date(time).toISOString()
    : new Date().toISOString();
}
