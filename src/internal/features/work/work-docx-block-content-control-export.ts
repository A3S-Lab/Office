import JSZip from 'jszip';
import {
  blockContentControlPropertiesFromElement,
  normalizeBlockContentControlProperties,
  type WorkDocumentBlockContentControlProperties,
} from './work-document-block-content-control';
import { ensureIgnorableContentControlNamespace } from './work-docx-note-comment-content-control-xml';
import { XML_NAMESPACE } from './work-docx-settings-xml';
import {
  descendants,
  parseXml,
  xmlNamespacePrefix,
} from './work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const WORD_NAMESPACES = new Set([WORD_NAMESPACE, STRICT_WORD_NAMESPACE]);
const WORD_2012_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2012/wordml';

export interface DocxBlockContentControlPatch {
  startMarker: string;
  endMarker: string;
  properties: WorkDocumentBlockContentControlProperties;
}

export class DocxBlockContentControlPatchCollector {
  readonly patches: DocxBlockContentControlPatch[] = [];
  private nextMarker = 1;

  register(
    properties: WorkDocumentBlockContentControlProperties,
  ): DocxBlockContentControlPatch {
    const index = this.nextMarker;
    this.nextMarker += 1;
    const patch: DocxBlockContentControlPatch = {
      startMarker: `__A3S_WORK_BLOCK_CONTENT_CONTROL_START_${index}__`,
      endMarker: `__A3S_WORK_BLOCK_CONTENT_CONTROL_END_${index}__`,
      properties: normalizeBlockContentControlProperties(properties),
    };
    this.patches.push(patch);
    return patch;
  }
}

/** Emits temporary marker paragraphs around a block content-control body. */
export function docxBlockContentControlParagraphs(
  element: HTMLElement,
  docx: typeof import('docx'),
  collector: DocxBlockContentControlPatchCollector,
  children: import('docx').Paragraph[],
): import('docx').Paragraph[] {
  const properties = blockContentControlPropertiesFromElement(element);
  const patch = collector.register(properties);
  return [
    new docx.Paragraph({
      children: [new docx.TextRun(patch.startMarker)],
    }),
    ...children,
    new docx.Paragraph({
      children: [new docx.TextRun(patch.endMarker)],
    }),
  ];
}

export async function patchDocxBlockContentControls(
  buffer: ArrayBuffer,
  patches: readonly DocxBlockContentControlPatch[],
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
      'Generated DOCX is missing word/document.xml for block content-control patches.',
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
          `Generated DOCX did not emit block content-control marker ${patch.startMarker}.`,
        );
      }
      if (!buildBlockContentControl(document, patch, allocateId, startParagraph)) {
        throw new Error(
          `Generated DOCX block content-control markers were incomplete for ${patch.startMarker}.`,
        );
      }
    }
    if (
      patches.some(
        (patch) =>
          patch.properties.appearance !== 'boundingBox' || patch.properties.color,
      )
    ) {
      ensureIgnorableContentControlNamespace(
        document.documentElement,
        WORD_2012_NAMESPACE,
        'w15',
      );
    }
    archive.file(entry.name, new XMLSerializer().serializeToString(document));
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function buildBlockContentControl(
  document: Document,
  patch: DocxBlockContentControlPatch,
  allocateId: (preferred: number | null) => number,
  startParagraph: Element,
): boolean {
  const parent = startParagraph.parentNode;
  if (!parent) return false;
  const endParagraph = findMarkerParagraph(document, patch.endMarker);
  if (!endParagraph || endParagraph.parentNode !== parent) return false;
  const bodyNodes: Element[] = [];
  let current: ChildNode | null = startParagraph.nextSibling;
  while (current && current !== endParagraph) {
    if (current.nodeType === 1) {
      const element = current as Element;
      if (isWordParagraph(element) || isWordSdt(element)) {
        bodyNodes.push(element);
      } else if ((element.textContent ?? '').trim()) {
        return false;
      }
    } else if (
      current.nodeType === 3 &&
      (current.textContent ?? '').trim()
    ) {
      return false;
    }
    current = current.nextSibling;
  }
  if (bodyNodes.length < 1) return false;
  const control = wordElement(document, startParagraph, 'sdt');
  const propertiesElement = wordElement(document, startParagraph, 'sdtPr');
  appendProperties(
    document,
    startParagraph,
    propertiesElement,
    patch.properties,
    allocateId(patch.properties.nativeId),
  );
  const content = wordElement(document, startParagraph, 'sdtContent');
  for (const body of bodyNodes) {
    if (isWordSdt(body)) {
      content.append(body);
      continue;
    }
    content.append(cleanWordParagraph(document, startParagraph, body));
  }
  control.append(propertiesElement, content);
  parent.insertBefore(control, startParagraph);
  startParagraph.remove();
  for (const body of bodyNodes) {
    if (body.parentNode === parent) body.remove();
  }
  endParagraph.remove();
  return true;
}

function isWordSdt(element: Element): boolean {
  return (
    WORD_NAMESPACES.has(element.namespaceURI ?? '') &&
    element.localName === 'sdt'
  );
}

function appendProperties(
  document: Document,
  context: Element,
  propertiesElement: Element,
  properties: WorkDocumentBlockContentControlProperties,
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
  if (properties.bindingStoreItemId && properties.bindingXPath) {
    const dataBinding = wordElement(document, context, 'dataBinding');
    setWordAttribute(
      dataBinding,
      context,
      'storeItemID',
      properties.bindingStoreItemId,
    );
    setWordAttribute(dataBinding, context, 'xpath', properties.bindingXPath);
    if (properties.bindingPrefixMappings) {
      setWordAttribute(
        dataBinding,
        context,
        'prefixMappings',
        properties.bindingPrefixMappings,
      );
    }
    propertiesElement.append(dataBinding);
  }
  if (properties.type === 'text') {
    propertiesElement.append(wordElement(document, context, 'text'));
  } else {
    propertiesElement.append(wordElement(document, context, 'richText'));
  }
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
      properties.color.slice(1).toUpperCase(),
    );
    propertiesElement.append(color);
  }
}

function cleanWordParagraph(
  document: Document,
  context: Element,
  source: Element,
): Element {
  const paragraph = wordElement(document, context, 'p');
  for (const child of Array.from(source.children)) {
    if (!WORD_NAMESPACES.has(child.namespaceURI ?? '')) continue;
    if (child.localName === 'pPr') continue;
    if (child.localName !== 'r') continue;
    const run = wordElement(document, context, 'r');
    for (const runChild of Array.from(child.children)) {
      if (!WORD_NAMESPACES.has(runChild.namespaceURI ?? '')) continue;
      if (runChild.localName === 'rPr') continue;
      if (runChild.localName === 't') {
        const text = wordElement(document, context, 't');
        const value = runChild.textContent ?? '';
        text.textContent = value;
        if (/^\s|\s$/u.test(value)) {
          text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
        }
        run.append(text);
      }
    }
    if (run.children.length) paragraph.append(run);
  }
  if (!paragraph.children.length) {
    const run = wordElement(document, context, 'r');
    const text = wordElement(document, context, 't');
    text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
    text.textContent = '';
    run.append(text);
    paragraph.append(run);
  }
  return paragraph;
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

function wordElement(
  document: Document,
  context: Element,
  localName: string,
): Element {
  const namespace = context.namespaceURI ?? WORD_NAMESPACE;
  const prefix = xmlNamespacePrefix(context, namespace) ?? 'w';
  return document.createElementNS(namespace, `${prefix}:${localName}`);
}

function namespacedElement(
  document: Document,
  context: Element,
  namespaceUri: string,
  prefix: string,
  localName: string,
): Element {
  const existing = xmlNamespacePrefix(context, namespaceUri) ?? prefix;
  return document.createElementNS(namespaceUri, `${existing}:${localName}`);
}

function setWordAttribute(
  element: Element,
  context: Element,
  localName: string,
  value: string,
): void {
  const namespace = context.namespaceURI ?? WORD_NAMESPACE;
  const prefix = xmlNamespacePrefix(context, namespace) ?? 'w';
  element.setAttributeNS(namespace, `${prefix}:${localName}`, value);
}

function setNamespacedAttribute(
  element: Element,
  context: Element,
  namespaceUri: string,
  prefix: string,
  localName: string,
  value: string,
): void {
  const existing = xmlNamespacePrefix(context, namespaceUri) ?? prefix;
  element.setAttributeNS(namespaceUri, `${existing}:${localName}`, value);
}

function wordAttribute(element: Element, localName: string): string | null {
  for (const attribute of Array.from(element.attributes)) {
    const name = attribute.localName || attribute.name.replace(/^.*:/, '');
    if (name !== localName) continue;
    const namespace = attribute.namespaceURI ?? element.namespaceURI ?? '';
    if (WORD_NAMESPACES.has(namespace) || !attribute.namespaceURI) {
      return attribute.value;
    }
  }
  return null;
}
