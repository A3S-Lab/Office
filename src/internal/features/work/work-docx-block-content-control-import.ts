import {
  blockContentControlDomAttributes,
  DOCUMENT_BLOCK_CONTENT_CONTROL_MAX_NEST_DEPTH,
  normalizeBlockContentControlProperties,
  type WorkDocumentBlockContentControlProperties,
} from './work-document-block-content-control';
import {
  CONTENT_CONTROL_WORD_2012_NAMESPACE,
  hasContentControlRelationshipReference,
  isDocxWordElement,
} from './work-docx-note-comment-content-control-xml';
import {
  XML_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { descendants, xmlNamespacePrefix } from './work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const WORD_NAMESPACES = new Set([WORD_NAMESPACE, STRICT_WORD_NAMESPACE]);
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MAX_IMPORTED_BLOCK_CONTROLS = 1024;
const MAX_BLOCK_CONTENT_PARAGRAPHS = 64;
const PROPERTY_NAMES = new Set([
  'alias',
  'dataBinding',
  'id',
  'lock',
  'richText',
  'tag',
  'text',
]);
const LOCK_VALUES = new Set([
  'contentLocked',
  'sdtContentLocked',
  'sdtLocked',
  'unlocked',
]);
const APPEARANCE_VALUES = new Set(['boundingBox', 'hidden', 'tags']);
const RUN_CONTENT_NAMES = new Set([
  'br',
  'cr',
  'noBreakHyphen',
  'softHyphen',
  'tab',
  't',
]);

export interface ImportedDocxBlockContentControlMarker {
  start: string;
  end: string;
  properties: WorkDocumentBlockContentControlProperties;
  text: string;
}

export interface ImportedDocxBlockContentControlMarkers {
  controls: ImportedDocxBlockContentControlMarker[];
  unsupported: number;
}

export interface DocxBlockContentControlInspection {
  supported: number;
  unsupported: number;
}

/**
 * Inspects body-level multi-paragraph text/rich-text `w:sdt` controls.
 * Repeating sections and inline paragraph controls stay on their own paths.
 */
export function inspectDocxBlockContentControls(
  document: Document,
): DocxBlockContentControlInspection {
  let supported = 0;
  let unsupported = 0;
  for (const control of blockContentControlCandidates(document)) {
    if (
      readBlockContentControl(
        control,
        supported + 1,
        DOCUMENT_BLOCK_CONTENT_CONTROL_MAX_NEST_DEPTH,
      )
    ) {
      supported += 1;
    } else unsupported += 1;
  }
  return { supported, unsupported };
}

export function markDocxBlockContentControls(
  document: Document,
): ImportedDocxBlockContentControlMarkers {
  const controls: ImportedDocxBlockContentControlMarker[] = [];
  let unsupported = 0;
  for (const control of blockContentControlCandidates(document)) {
    const nested = markNestedBlockContentControls(
      document,
      control,
      controls.length + 1,
      DOCUMENT_BLOCK_CONTENT_CONTROL_MAX_NEST_DEPTH,
    );
    if (nested === null) {
      unsupported += 1;
      continue;
    }
    controls.push(...nested);
    const parsed = readLeafBlockContentControl(
      control,
      controls.length + 1,
      DOCUMENT_BLOCK_CONTENT_CONTROL_MAX_NEST_DEPTH,
    );
    if (!parsed) {
      unsupported += 1;
      continue;
    }
    if (controls.length >= MAX_IMPORTED_BLOCK_CONTROLS) {
      throw new Error('Imported DOCX exceeds the block content-control limit.');
    }
    const [start, end] = blockContentControlMarkers(document, controls.length + 1);
    if (!replaceBlockControlWithMarkers(document, control, start, end)) {
      unsupported += 1;
      continue;
    }
    controls.push({
      start,
      end,
      properties: parsed.properties,
      text: parsed.text,
    });
  }
  return { controls, unsupported };
}

export function applyImportedDocxBlockContentControlMarkers(
  document: Document,
  markers: ImportedDocxBlockContentControlMarkers,
): void {
  if (!markers.controls.length) return;
  // Innermost markers are registered before their parents; apply children first.
  for (const marker of [...markers.controls].reverse()) {
    const wrapper = document.createElement('div');
    for (const [name, value] of Object.entries(
      blockContentControlDomAttributes(marker.properties),
    )) {
      if (value !== undefined) wrapper.setAttribute(name, value);
    }
    wrapper.className = 'work-document-block-content-control';
    wrapper.setAttribute('role', 'group');
    wrapper.setAttribute(
      'aria-label',
      marker.properties.alias || marker.properties.tag || '块内容控件',
    );
    moveNodesBetween(document.body, marker.start, marker.end, wrapper);
    if (!wrapper.querySelector('p, [data-document-block-content-control]')) {
      const paragraph = document.createElement('p');
      paragraph.textContent = marker.text;
      wrapper.append(paragraph);
    }
  }
  document.body.normalize();
}

export function hasImportedDocxBlockContentControlMarkers(
  markers: ImportedDocxBlockContentControlMarkers,
): boolean {
  return markers.controls.length > 0;
}

interface ParsedBlockContentControl {
  properties: WorkDocumentBlockContentControlProperties;
  text: string;
}

function blockContentControlCandidates(document: Document): Element[] {
  return descendants(document, 'sdt').filter((control) => {
    if (!isDocxWordElement(control) || hasContentControlAncestor(control)) {
      return false;
    }
    if (!isBodyLevelControl(control)) return false;
    if (hasRepeatingSectionProperty(control)) return false;
    if (isStructuralContentControl(control)) return false;
    return true;
  });
}

function hasContentControlAncestor(control: Element): boolean {
  let current = control.parentElement;
  while (current) {
    if (isDocxWordElement(current) && current.localName === 'sdt') return true;
    current = current.parentElement;
  }
  return false;
}

function isBodyLevelControl(control: Element): boolean {
  const parent = control.parentElement;
  return Boolean(
    parent &&
      isDocxWordElement(parent) &&
      (parent.localName === 'body' || parent.localName === 'tc'),
  );
}

function hasRepeatingSectionProperty(control: Element): boolean {
  const properties = Array.from(control.children).find(
    (child) => isDocxWordElement(child) && child.localName === 'sdtPr',
  );
  if (!properties) return false;
  return Array.from(properties.children).some(
    (child) =>
      child.namespaceURI === CONTENT_CONTROL_WORD_2012_NAMESPACE &&
      child.localName === 'repeatingSection',
  );
}

function isStructuralContentControl(control: Element): boolean {
  const properties = Array.from(control.children).find(
    (child) => isDocxWordElement(child) && child.localName === 'sdtPr',
  );
  if (!properties) return false;
  return Array.from(properties.children).some(
    (child) =>
      isDocxWordElement(child) &&
      ['docPartObj', 'docPartList'].includes(child.localName),
  );
}

function markNestedBlockContentControls(
  document: Document,
  control: Element,
  nextSequence: number,
  maxChildDepth: number,
): ImportedDocxBlockContentControlMarker[] | null {
  const content = Array.from(control.children).find(
    (child) => isDocxWordElement(child) && child.localName === 'sdtContent',
  );
  if (!content) return null;
  if (maxChildDepth < 1) return [];
  const nested: ImportedDocxBlockContentControlMarker[] = [];
  let sequence = nextSequence;
  for (const child of Array.from(content.children)) {
    if (!isDocxWordElement(child) || child.localName !== 'sdt') continue;
    if (maxChildDepth >= 2) {
      const deeper = markNestedBlockContentControls(
        document,
        child,
        sequence,
        maxChildDepth - 1,
      );
      if (deeper === null) return null;
      nested.push(...deeper);
      sequence = nextSequence + nested.length;
    }
    const parsed = readLeafBlockContentControl(child, sequence, 0);
    if (!parsed) return null;
    if (nested.length + 1 >= MAX_IMPORTED_BLOCK_CONTROLS) {
      throw new Error('Imported DOCX exceeds the block content-control limit.');
    }
    const [start, end] = blockContentControlMarkers(document, sequence);
    if (!replaceBlockControlWithMarkers(document, child, start, end)) {
      return null;
    }
    nested.push({
      start,
      end,
      properties: parsed.properties,
      text: parsed.text,
    });
    sequence += 1;
  }
  return nested;
}

function readBlockContentControl(
  control: Element,
  sequence: number,
  remainingNestDepth: number,
): ParsedBlockContentControl | null {
  return readLeafBlockContentControl(control, sequence, remainingNestDepth);
}

function readLeafBlockContentControl(
  control: Element,
  sequence: number,
  remainingNestDepth: number,
): ParsedBlockContentControl | null {
  if (
    !isDocxWordElement(control) ||
    control.localName !== 'sdt' ||
    hasContentControlRelationshipReference(control) ||
    !validContainerAttributes(control)
  ) {
    return null;
  }
  const children = Array.from(control.children);
  const propertiesElements = children.filter(
    (child) => isDocxWordElement(child) && child.localName === 'sdtPr',
  );
  const contentElements = children.filter(
    (child) => isDocxWordElement(child) && child.localName === 'sdtContent',
  );
  if (
    propertiesElements.length !== 1 ||
    contentElements.length !== 1 ||
    children.some(
      (child) =>
        !isDocxWordElement(child) ||
        (child.localName !== 'sdtPr' &&
          child.localName !== 'sdtContent' &&
          child.localName !== 'sdtEndPr'),
    )
  ) {
    return null;
  }
  const properties = readProperties(propertiesElements[0]!, sequence);
  if (!properties) return null;
  const contentChildren = Array.from(contentElements[0]!.children);
  if (
    contentChildren.length < 1 ||
    contentChildren.length > MAX_BLOCK_CONTENT_PARAGRAPHS
  ) {
    return null;
  }
  const paragraphTexts: string[] = [];
  for (const child of contentChildren) {
    if (!isDocxWordElement(child)) return null;
    if (child.localName === 'p') {
      const runs = Array.from(child.children).filter(
        (runChild) => isDocxWordElement(runChild) && runChild.localName === 'r',
      );
      if (
        Array.from(child.children).some(
          (runChild) =>
            !isDocxWordElement(runChild) || runChild.localName !== 'r',
        )
      ) {
        return null;
      }
      if (runs.some((run) => !validRun(run))) return null;
      const text = runs.map(runText).join('');
      if (text.includes('__A3S_WORK_BLOCK_CONTENT_CONTROL_')) continue;
      paragraphTexts.push(text);
      continue;
    }
    if (child.localName === 'sdt') {
      if (remainingNestDepth < 1) return null;
      if (
        !readLeafBlockContentControl(child, sequence, remainingNestDepth - 1)
      ) {
        return null;
      }
      paragraphTexts.push('');
      continue;
    }
    return null;
  }
  return { properties, text: paragraphTexts.filter(Boolean).join('\n') };
}

function readProperties(
  properties: Element,
  sequence: number,
): WorkDocumentBlockContentControlProperties | null {
  if (!validContainerAttributes(properties)) return null;
  const groups = new Map<string, Element[]>();
  let appearance: WorkDocumentBlockContentControlProperties['appearance'] =
    'boundingBox';
  let color: string | null = null;
  let bindingStoreItemId = '';
  let bindingXPath = '';
  let bindingPrefixMappings = '';
  for (const child of Array.from(properties.children)) {
    const namespace = child.namespaceURI ?? '';
    if (isDocxWordElement(child)) {
      if (!PROPERTY_NAMES.has(child.localName)) return null;
      const matches = groups.get(child.localName) ?? [];
      matches.push(child);
      groups.set(child.localName, matches);
      continue;
    }
    if (namespace === CONTENT_CONTROL_WORD_2012_NAMESPACE) {
      if (child.localName === 'appearance' || child.localName === 'color') {
        const matches = groups.get(`w15:${child.localName}`) ?? [];
        matches.push(child);
        groups.set(`w15:${child.localName}`, matches);
        continue;
      }
      return null;
    }
    if (namespace === MARKUP_COMPATIBILITY_NAMESPACE) continue;
    return null;
  }
  if (Array.from(groups.values()).some((items) => items.length > 1)) return null;
  const textElement = groups.get('text')?.[0];
  const richTextElement = groups.get('richText')?.[0];
  if (textElement && richTextElement) return null;
  if (
    groups.has('comboBox') ||
    groups.has('date') ||
    groups.has('dropDownList')
  ) {
    return null;
  }
  if (textElement && !validLeaf(textElement, new Set(['multiLine']))) {
    return null;
  }
  if (richTextElement && !validLeaf(richTextElement, new Set())) return null;
  const id = groups.get('id')?.[0];
  if (id && !validLeaf(id, new Set(['val']))) return null;
  const nativeId = id ? parseNativeId(wordAttribute(id, 'val')) : null;
  if (id && nativeId === null) return null;
  const alias = readStringProperty(groups.get('alias')?.[0]);
  const tag = readStringProperty(groups.get('tag')?.[0]);
  if (alias === null || tag === null) return null;
  const lockElement = groups.get('lock')?.[0];
  if (lockElement && !validLeaf(lockElement, new Set(['val']))) return null;
  const lockValue = lockElement ? wordAttribute(lockElement, 'val') : null;
  const lock = lockValue ?? 'unlocked';
  if (!LOCK_VALUES.has(lock)) return null;
  const bindingElement = groups.get('dataBinding')?.[0];
  if (bindingElement) {
    const binding = readDataBinding(bindingElement);
    if (!binding) return null;
    bindingStoreItemId = binding.bindingStoreItemId;
    bindingXPath = binding.bindingXPath;
    bindingPrefixMappings = binding.bindingPrefixMappings;
  }
  const appearanceElement = groups.get('w15:appearance')?.[0];
  if (appearanceElement) {
    if (
      !validNamespacedLeaf(
        appearanceElement,
        new Set(['val']),
        CONTENT_CONTROL_WORD_2012_NAMESPACE,
      )
    ) {
      return null;
    }
    const value = namespacedAttribute(
      appearanceElement,
      'val',
      CONTENT_CONTROL_WORD_2012_NAMESPACE,
    );
    if (!value || !APPEARANCE_VALUES.has(value)) return null;
    appearance =
      value as WorkDocumentBlockContentControlProperties['appearance'];
  }
  const colorElement = groups.get('w15:color')?.[0];
  if (colorElement) {
    if (
      !validNamespacedLeaf(
        colorElement,
        new Set(['val']),
        CONTENT_CONTROL_WORD_2012_NAMESPACE,
      )
    ) {
      return null;
    }
    const value = namespacedAttribute(
      colorElement,
      'val',
      CONTENT_CONTROL_WORD_2012_NAMESPACE,
    );
    if (!value || !/^[0-9a-f]{6}$/i.test(value)) return null;
    color = `#${value.toLowerCase()}`;
  }
  return normalizeBlockContentControlProperties({
    id: `docx-block-content-control-${sequence}`,
    nativeId,
    type: textElement ? 'text' : 'richText',
    alias,
    tag,
    lock: lock as WorkDocumentBlockContentControlProperties['lock'],
    appearance,
    color,
    bindingStoreItemId,
    bindingXPath,
    bindingPrefixMappings,
  });
}

function readDataBinding(binding: Element): {
  bindingPrefixMappings: string;
  bindingStoreItemId: string;
  bindingXPath: string;
} | null {
  if (!isDocxWordElement(binding) || binding.localName !== 'dataBinding') {
    return null;
  }
  if (hasNonWhitespaceText(binding) || binding.children.length) return null;
  const allowed = new Set(['prefixMappings', 'storeItemID', 'xpath']);
  for (const attribute of Array.from(binding.attributes)) {
    const namespace = xmlAttributeNamespace(binding, attribute);
    if (namespace === 'http://www.w3.org/2000/xmlns/') continue;
    if (namespace === MARKUP_COMPATIBILITY_NAMESPACE) continue;
    if (
      !WORD_NAMESPACES.has(namespace ?? '') ||
      !allowed.has(xmlAttributeLocalName(attribute))
    ) {
      return null;
    }
  }
  const normalized = normalizeBlockContentControlProperties({
    bindingStoreItemId: wordAttribute(binding, 'storeItemID') ?? '',
    bindingXPath: wordAttribute(binding, 'xpath') ?? '',
    bindingPrefixMappings: wordAttribute(binding, 'prefixMappings') ?? '',
  });
  if (!normalized.bindingStoreItemId || !normalized.bindingXPath) return null;
  return {
    bindingStoreItemId: normalized.bindingStoreItemId,
    bindingXPath: normalized.bindingXPath,
    bindingPrefixMappings: normalized.bindingPrefixMappings,
  };
}

function replaceBlockControlWithMarkers(
  document: Document,
  control: Element,
  start: string,
  end: string,
): boolean {
  const parent = control.parentNode;
  if (!parent) return false;
  const content = Array.from(control.children).find(
    (child) => isDocxWordElement(child) && child.localName === 'sdtContent',
  );
  if (!content) return false;
  const paragraphs = Array.from(content.children).filter(
    (child) => isDocxWordElement(child) && child.localName === 'p',
  );
  if (
    paragraphs.length < 1 ||
    paragraphs.length > MAX_BLOCK_CONTENT_PARAGRAPHS
  ) {
    return false;
  }
  parent.insertBefore(markerParagraph(document, control, start), control);
  for (const paragraph of paragraphs) {
    parent.insertBefore(paragraph, control);
  }
  parent.insertBefore(markerParagraph(document, control, end), control);
  control.remove();
  return true;
}

function markerParagraph(
  document: Document,
  context: Element,
  value: string,
): Element {
  const namespace = context.namespaceURI ?? WORD_NAMESPACE;
  const prefix = xmlNamespacePrefix(context, namespace) ?? 'w';
  const paragraph = document.createElementNS(namespace, `${prefix}:p`);
  const run = document.createElementNS(namespace, `${prefix}:r`);
  const text = document.createElementNS(namespace, `${prefix}:t`);
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = value;
  run.append(text);
  paragraph.append(run);
  return paragraph;
}

function blockContentControlMarkers(
  document: Document,
  initialIndex: number,
): [string, string] {
  const occupiedText = document.documentElement.textContent ?? '';
  let index = initialIndex;
  for (let attempts = 0; attempts < 1_000_000; attempts += 1) {
    const start = `__A3S_WORK_BLOCK_CONTENT_CONTROL_START_${index}__`;
    const end = `__A3S_WORK_BLOCK_CONTENT_CONTROL_END_${index}__`;
    if (!occupiedText.includes(start) && !occupiedText.includes(end)) {
      return [start, end];
    }
    index += 1;
  }
  throw new Error('Imported DOCX block content-control marker space is exhausted.');
}

function moveNodesBetween(
  root: HTMLElement,
  start: string,
  end: string,
  replacement: HTMLElement,
): void {
  const startElement = findTextParagraph(root, start);
  const endElement = findTextParagraph(root, end);
  if (!startElement || !endElement || !startElement.parentNode) return;
  const parent = startElement.parentNode;
  const collected: ChildNode[] = [];
  let current: ChildNode | null = startElement.nextSibling;
  while (current && current !== endElement) {
    const next = current.nextSibling;
    collected.push(current);
    current = next;
  }
  for (const node of collected) {
    if (node.nodeType !== Node.ELEMENT_NODE) continue;
    const element = node as Element;
    if (
      element.localName === 'p' ||
      element.hasAttribute('data-document-block-content-control')
    ) {
      replacement.append(node);
    }
  }
  parent.insertBefore(replacement, startElement);
  startElement.remove();
  endElement.remove();
  for (const node of collected) {
    if (node.parentNode && node.parentNode !== replacement) {
      node.parentNode.removeChild(node);
    }
  }
}

function findTextParagraph(root: HTMLElement, marker: string): Element | null {
  const walker = root.ownerDocument.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let current = walker.nextNode();
  while (current) {
    if ((current.textContent ?? '').includes(marker)) {
      let element: Element | null =
        current.parentElement instanceof Element ? current.parentElement : null;
      while (element && element !== root) {
        if (element.localName === 'p') return element;
        element = element.parentElement;
      }
    }
    current = walker.nextNode();
  }
  return null;
}

function validRun(run: Element): boolean {
  if (!isDocxWordElement(run) || run.localName !== 'r') return false;
  for (const child of Array.from(run.children)) {
    if (!isDocxWordElement(child)) return false;
    if (child.localName === 'rPr') continue;
    if (!RUN_CONTENT_NAMES.has(child.localName)) return false;
    if (child.children.length) return false;
  }
  return true;
}

function runText(run: Element): string {
  let value = '';
  for (const child of Array.from(run.children)) {
    if (!isDocxWordElement(child)) continue;
    if (child.localName === 't') value += child.textContent ?? '';
    else if (child.localName === 'tab') value += '\t';
    else if (child.localName === 'br' || child.localName === 'cr') value += '\n';
    else if (child.localName === 'noBreakHyphen') value += '\u2011';
    else if (child.localName === 'softHyphen') value += '\u00ad';
  }
  return value;
}

function validContainerAttributes(element: Element): boolean {
  for (const attribute of Array.from(element.attributes)) {
    const namespace = xmlAttributeNamespace(element, attribute);
    if (namespace === 'http://www.w3.org/2000/xmlns/') continue;
    if (namespace === MARKUP_COMPATIBILITY_NAMESPACE) continue;
    return false;
  }
  return true;
}

function validLeaf(element: Element, allowed: Set<string>): boolean {
  if (hasNonWhitespaceText(element) || element.children.length) return false;
  for (const attribute of Array.from(element.attributes)) {
    const namespace = xmlAttributeNamespace(element, attribute);
    if (namespace === 'http://www.w3.org/2000/xmlns/') continue;
    if (namespace === MARKUP_COMPATIBILITY_NAMESPACE) continue;
    if (
      !WORD_NAMESPACES.has(namespace ?? '') ||
      !allowed.has(xmlAttributeLocalName(attribute))
    ) {
      return false;
    }
  }
  return true;
}

function validNamespacedLeaf(
  element: Element,
  allowed: Set<string>,
  namespaceUri: string,
): boolean {
  if (hasNonWhitespaceText(element) || element.children.length) return false;
  for (const attribute of Array.from(element.attributes)) {
    const namespace = xmlAttributeNamespace(element, attribute);
    if (namespace === 'http://www.w3.org/2000/xmlns/') continue;
    if (namespace === MARKUP_COMPATIBILITY_NAMESPACE) continue;
    if (
      namespace !== namespaceUri ||
      !allowed.has(xmlAttributeLocalName(attribute))
    ) {
      return false;
    }
  }
  return true;
}

function readStringProperty(element: Element | undefined): string | null {
  if (!element) return '';
  if (!validLeaf(element, new Set(['val']))) return null;
  return wordAttribute(element, 'val') ?? '';
}

function parseNativeId(value: string | null): number | null {
  if (value === null || value === '') return null;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) &&
    parsed >= -2_147_483_648 &&
    parsed <= 2_147_483_647
    ? parsed
    : null;
}

function wordAttribute(element: Element, localName: string): string | null {
  for (const attribute of Array.from(element.attributes)) {
    if (
      xmlAttributeLocalName(attribute) === localName &&
      WORD_NAMESPACES.has(xmlAttributeNamespace(element, attribute) ?? '')
    ) {
      return attribute.value;
    }
  }
  return null;
}

function namespacedAttribute(
  element: Element,
  localName: string,
  namespaceUri: string,
): string | null {
  for (const attribute of Array.from(element.attributes)) {
    if (
      xmlAttributeLocalName(attribute) === localName &&
      xmlAttributeNamespace(element, attribute) === namespaceUri
    ) {
      return attribute.value;
    }
  }
  return null;
}

function hasNonWhitespaceText(element: Element): boolean {
  return Array.from(element.childNodes).some(
    (node) =>
      node.nodeType === Node.TEXT_NODE &&
      (node.textContent ?? '').trim().length > 0,
  );
}
