import JSZip from 'jszip';
import {
  normalizeRepeatingSectionItemProperties,
  normalizeRepeatingSectionProperties,
  repeatingSectionItemPropertiesFromElement,
  repeatingSectionPropertiesFromElement,
  type WorkDocumentRepeatingSectionItemProperties,
  type WorkDocumentRepeatingSectionProperties,
} from './work-document-repeating-section';
import { ensureIgnorableContentControlNamespace } from './work-docx-note-comment-content-control-xml';
import { XML_NAMESPACE } from './work-docx-settings-xml';
import {
  descendants,
  parseXml,
} from './work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const WORD_NAMESPACES = new Set([WORD_NAMESPACE, STRICT_WORD_NAMESPACE]);
const STORY_PATTERN =
  /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/i;
const WORD_2012_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2012/wordml';

export interface DocxRepeatingSectionPatch {
  startMarker: string;
  endMarker: string;
  properties: WorkDocumentRepeatingSectionProperties;
  items: Array<{
    startMarker: string;
    endMarker: string;
    properties: WorkDocumentRepeatingSectionItemProperties;
  }>;
}

export class DocxRepeatingSectionPatchCollector {
  readonly patches: DocxRepeatingSectionPatch[] = [];
  private nextMarker = 1;

  register(
    properties: WorkDocumentRepeatingSectionProperties,
    items: readonly WorkDocumentRepeatingSectionItemProperties[],
  ): DocxRepeatingSectionPatch {
    const index = this.nextMarker;
    this.nextMarker += 1;
    const patch: DocxRepeatingSectionPatch = {
      startMarker: `__A3S_WORK_REPEATING_SECTION_START_${index}__`,
      endMarker: `__A3S_WORK_REPEATING_SECTION_END_${index}__`,
      properties: normalizeRepeatingSectionProperties(properties),
      items: items.map((item, itemIndex) => ({
        startMarker: `__A3S_WORK_REPEATING_SECTION_ITEM_START_${index}_${itemIndex + 1}__`,
        endMarker: `__A3S_WORK_REPEATING_SECTION_ITEM_END_${index}_${itemIndex + 1}__`,
        properties: normalizeRepeatingSectionItemProperties(item),
      })),
    };
    this.patches.push(patch);
    return patch;
  }
}

/** Emits temporary marker paragraphs for a repeating-section HTML block. */
export function docxRepeatingSectionParagraphs(
  element: HTMLElement,
  docx: typeof import('docx'),
  collector: DocxRepeatingSectionPatchCollector,
  itemParagraphs: Array<InstanceType<typeof docx.Paragraph>>,
): Array<InstanceType<typeof docx.Paragraph>> {
  const properties = repeatingSectionPropertiesFromElement(element);
  const itemElements = Array.from(
    element.querySelectorAll<HTMLElement>(
      ':scope > [data-document-repeating-section-item]',
    ),
  );
  if (!itemElements.length || itemElements.length !== itemParagraphs.length) {
    throw new Error(
      'Document repeating sections require one paragraph body per item.',
    );
  }
  const itemProperties = itemElements.map((item) =>
    repeatingSectionItemPropertiesFromElement(item),
  );
  const patch = collector.register(properties, itemProperties);
  const paragraphs: Array<InstanceType<typeof docx.Paragraph>> = [
    new docx.Paragraph({
      children: [new docx.TextRun(patch.startMarker)],
    }),
  ];
  for (const [index, itemParagraph] of itemParagraphs.entries()) {
    paragraphs.push(
      new docx.Paragraph({
        children: [new docx.TextRun(patch.items[index]!.startMarker)],
      }),
      itemParagraph,
      new docx.Paragraph({
        children: [new docx.TextRun(patch.items[index]!.endMarker)],
      }),
    );
  }
  paragraphs.push(
    new docx.Paragraph({
      children: [new docx.TextRun(patch.endMarker)],
    }),
  );
  return paragraphs;
}

/**
 * Replaces generated repeating-section marker paragraphs with native body-level
 * `w15:repeatingSection` / `w15:repeatingSectionItem` controls.
 */
export async function patchDocxRepeatingSections(
  buffer: ArrayBuffer,
  patches: readonly DocxRepeatingSectionPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  const entries = Object.values(archive.files).filter(
    (entry) =>
      !entry.dir &&
      /(^|\/)word\/document\.xml$/i.test(entry.name.replace(/\\/g, '/')),
  );
  if (!entries.length) {
    throw new Error(
      'Generated DOCX is missing word/document.xml for repeating-section patches.',
    );
  }
  const documents = await Promise.all(
    entries.map(async (entry) => ({
      entry,
      document: parseXml(
        await entry.async('text'),
        `generated DOCX ${entry.name}`,
      ),
    })),
  );
  const usedIds = new Set<number>();
  for (const { document } of documents) {
    for (const id of descendants(document, 'id')) {
      const value = wordAttribute(id, 'val');
      const parsed = value === null ? null : Number(value);
      if (
        parsed !== null &&
        Number.isSafeInteger(parsed) &&
        parsed >= -2_147_483_648 &&
        parsed <= 2_147_483_647
      ) {
        usedIds.add(parsed);
      }
    }
  }
  let nextId = 1;
  const allocateId = (preferred: number | null): number => {
    if (
      preferred !== null &&
      !usedIds.has(preferred) &&
      Number.isSafeInteger(preferred)
    ) {
      usedIds.add(preferred);
      return preferred;
    }
    while (usedIds.has(nextId)) nextId += 1;
    const id = nextId;
    usedIds.add(id);
    nextId += 1;
    return id;
  };

  for (const { entry, document } of documents) {
    for (const patch of patches) {
      const startParagraph = findMarkerParagraph(document, patch.startMarker);
      if (!startParagraph) {
        throw new Error(
          `Generated DOCX did not emit repeating-section marker ${patch.startMarker}.`,
        );
      }
      const built = buildRepeatingSection(
        document,
        patch,
        allocateId,
        startParagraph,
      );
      if (!built) {
        throw new Error(
          `Generated DOCX repeating-section markers were incomplete for ${patch.startMarker}.`,
        );
      }
    }
    ensureIgnorableContentControlNamespace(
      document.documentElement,
      WORD_2012_NAMESPACE,
      'w15',
    );
    archive.file(entry.name, new XMLSerializer().serializeToString(document));
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function findMarkerParagraph(
  document: Document,
  marker: string,
): Element | null {
  for (const text of descendants(document, 't')) {
    if (!(text.textContent ?? '').includes(marker)) continue;
    let current: Element | null = text;
    while (current) {
      if (isWordParagraph(current)) return current;
      current = current.parentElement;
    }
  }
  return null;
}

function isWordParagraph(element: Element): boolean {
  return (
    element.localName === 'p' &&
    WORD_NAMESPACES.has(element.namespaceURI ?? '')
  );
}

function buildRepeatingSection(
  document: Document,
  patch: DocxRepeatingSectionPatch,
  allocateId: (preferred: number | null) => number,
  startParagraph: Element,
): Element | null {
  const parent = startParagraph.parentNode;
  if (!parent) return null;
  const siblings = Array.from(parent.childNodes);
  const startIndex = siblings.indexOf(startParagraph);
  if (startIndex < 0) return null;

  let cursor = startIndex + 1;
  const itemBodies: Element[] = [];
  for (const item of patch.items) {
    const itemStart = siblings[cursor];
    if (
      !itemStart ||
      itemStart.nodeType !== Node.ELEMENT_NODE ||
      !(itemStart as Element).textContent?.includes(item.startMarker)
    ) {
      return null;
    }
    cursor += 1;
    const body = siblings[cursor];
    if (!body || body.nodeType !== Node.ELEMENT_NODE) return null;
    itemBodies.push(body as Element);
    cursor += 1;
    const itemEnd = siblings[cursor];
    if (
      !itemEnd ||
      itemEnd.nodeType !== Node.ELEMENT_NODE ||
      !(itemEnd as Element).textContent?.includes(item.endMarker)
    ) {
      return null;
    }
    cursor += 1;
  }
  const endParagraph = siblings[cursor];
  if (
    !endParagraph ||
    endParagraph.nodeType !== Node.ELEMENT_NODE ||
    !(endParagraph as Element).textContent?.includes(patch.endMarker)
  ) {
    return null;
  }

  const context = wordContext(document);
  const section = wordElement(document, context, 'sdt');
  const sectionPr = wordElement(document, context, 'sdtPr');
  appendSectionProperties(
    document,
    context,
    sectionPr,
    patch.properties,
    allocateId(patch.properties.nativeId),
  );
  const sectionContent = wordElement(document, context, 'sdtContent');
  for (const [index, body] of itemBodies.entries()) {
    const item = wordElement(document, context, 'sdt');
    const itemPr = wordElement(document, context, 'sdtPr');
    const itemId = wordElement(document, context, 'id');
    setWordAttribute(
      itemId,
      context,
      'val',
      String(allocateId(patch.items[index]!.properties.nativeId)),
    );
    itemPr.append(itemId);
    const repeatingItem = namespacedElement(
      document,
      context,
      WORD_2012_NAMESPACE,
      'w15',
      'repeatingSectionItem',
    );
    itemPr.append(repeatingItem);
    const itemContent = wordElement(document, context, 'sdtContent');
    itemContent.append(cleanWordParagraph(document, context, body));
    item.append(itemPr, itemContent);
    sectionContent.append(item);
  }
  section.append(sectionPr, sectionContent);

  const removeUntil = cursor;
  for (let index = removeUntil; index >= startIndex; index -= 1) {
    const node = siblings[index];
    if (node) parent.removeChild(node);
  }
  parent.insertBefore(section, siblings[removeUntil + 1] ?? null);
  return section;
}

/** Rebuilds a transitional Word paragraph from plain text, dropping foreign attrs. */
function cleanWordParagraph(
  document: Document,
  context: string,
  source: Element,
): Element {
  const paragraph = wordElement(document, context, 'p');
  const run = wordElement(document, context, 'r');
  const text = wordElement(document, context, 't');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = source.textContent ?? '';
  run.append(text);
  paragraph.append(run);
  return paragraph;
}

function appendSectionProperties(
  document: Document,
  context: string,
  propertiesElement: Element,
  properties: WorkDocumentRepeatingSectionProperties,
  nativeId: number,
): void {
  const id = wordElement(document, context, 'id');
  setWordAttribute(id, context, 'val', String(nativeId));
  propertiesElement.append(id);
  if (properties.alias) {
    const alias = wordElement(document, context, 'alias');
    setWordAttribute(alias, context, 'val', properties.alias);
    propertiesElement.append(alias);
  }
  if (properties.tag) {
    const tag = wordElement(document, context, 'tag');
    setWordAttribute(tag, context, 'val', properties.tag);
    propertiesElement.append(tag);
  }
  if (properties.lock !== 'unlocked') {
    const lock = wordElement(document, context, 'lock');
    setWordAttribute(lock, context, 'val', properties.lock);
    propertiesElement.append(lock);
  }
  const repeating = namespacedElement(
    document,
    context,
    WORD_2012_NAMESPACE,
    'w15',
    'repeatingSection',
  );
  propertiesElement.append(repeating);
  if (properties.appearance !== 'boundingBox') {
    const appearance = namespacedElement(
      document,
      context,
      WORD_2012_NAMESPACE,
      'w15',
      'appearance',
    );
    setNamespacedAttribute(
      appearance,
      context,
      WORD_2012_NAMESPACE,
      'w15',
      'val',
      properties.appearance,
    );
    propertiesElement.append(appearance);
  }
  if (properties.color) {
    const color = namespacedElement(
      document,
      context,
      WORD_2012_NAMESPACE,
      'w15',
      'color',
    );
    setNamespacedAttribute(
      color,
      context,
      WORD_2012_NAMESPACE,
      'w15',
      'val',
      properties.color.replace(/^#/, '').toUpperCase(),
    );
    propertiesElement.append(color);
  }
}

function wordContext(document: Document): string {
  const root = document.documentElement;
  return WORD_NAMESPACES.has(root.namespaceURI ?? '')
    ? root.namespaceURI === STRICT_WORD_NAMESPACE
      ? STRICT_WORD_NAMESPACE
      : WORD_NAMESPACE
    : WORD_NAMESPACE;
}

function wordElement(
  document: Document,
  context: string,
  localName: string,
): Element {
  return document.createElementNS(context, `w:${localName}`);
}

function namespacedElement(
  document: Document,
  _context: string,
  namespace: string,
  prefix: string,
  localName: string,
): Element {
  return document.createElementNS(namespace, `${prefix}:${localName}`);
}

function setWordAttribute(
  element: Element,
  context: string,
  localName: string,
  value: string,
): void {
  element.setAttributeNS(context, `w:${localName}`, value);
}

function setNamespacedAttribute(
  element: Element,
  _context: string,
  namespace: string,
  prefix: string,
  localName: string,
  value: string,
): void {
  element.setAttributeNS(namespace, `${prefix}:${localName}`, value);
}

function wordAttribute(element: Element, localName: string): string | null {
  for (const attribute of Array.from(element.attributes)) {
    if (
      attribute.localName === localName &&
      WORD_NAMESPACES.has(attribute.namespaceURI ?? element.namespaceURI ?? '')
    ) {
      return attribute.value;
    }
  }
  return element.getAttribute(`w:${localName}`);
}
