import {
  contentControlDomAttributes,
  normalizeDocumentContentControlProperties,
  type WorkDocumentContentControlProperties,
} from './work-document-content-control';
import {
  CONTENT_CONTROL_WORD_2012_NAMESPACE,
  hasContentControlRelationshipReference,
  hasUnsupportedContentControlSemanticChild,
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
const MAX_IMPORTED_CONTENT_CONTROLS = 4096;
const MAX_CONTENT_CONTROL_TEXT = 1_000_000;
const RUN_CONTENT_NAMES = new Set([
  'br',
  'cr',
  'noBreakHyphen',
  'softHyphen',
  'tab',
  't',
]);
const PROPERTY_NAMES = new Set([
  'alias',
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

export interface ImportedDocxContentControlMarker {
  start: string;
  end: string;
  properties: WorkDocumentContentControlProperties;
}

export interface ImportedDocxContentControlMarkers {
  controls: ImportedDocxContentControlMarker[];
  unsupported: number;
}

export interface DocxContentControlInspection {
  supported: number;
  unsupported: number;
}

/**
 * Inspects body-level inline `w:sdt` controls without changing the package.
 * Structural controls such as TOC/INDEX are left to their dedicated
 * importers, so they are not reported as unsupported form controls here.
 */
export function inspectDocxContentControls(
  document: Document,
): DocxContentControlInspection {
  let supported = 0;
  let unsupported = 0;
  for (const control of descendants(document, 'sdt')) {
    if (hasContentControlAncestor(control)) continue;
    if (isStructuralContentControl(control)) continue;
    if (isDirectParagraphControl(control) && readContentControl(control, 0))
      supported += 1;
    else unsupported += 1;
  }
  return { supported, unsupported };
}

/**
 * Moves the bounded inline control body into ordinary runs and surrounds it
 * with text markers. Mammoth can then convert the runs normally; the markers
 * are wrapped back into an editable HTML node after conversion.
 */
export function markDocxContentControls(
  document: Document,
): ImportedDocxContentControlMarkers {
  const controls: ImportedDocxContentControlMarker[] = [];
  let unsupported = 0;
  const candidates = descendants(document, 'sdt').filter(
    (control) =>
      !hasContentControlAncestor(control) &&
      !isStructuralContentControl(control),
  );
  for (const [index, control] of candidates.entries()) {
    if (!isDirectParagraphControl(control)) {
      unsupported += 1;
      continue;
    }
    const properties = readContentControl(control, index + 1);
    if (!properties) {
      unsupported += 1;
      continue;
    }
    if (controls.length >= MAX_IMPORTED_CONTENT_CONTROLS) {
      throw new Error('Imported DOCX exceeds the content-control limit.');
    }
    const [start, end] = contentControlMarkers(document, controls.length + 1);
    if (!replaceControlWithMarkers(document, control, start, end)) {
      unsupported += 1;
      continue;
    }
    controls.push({ start, end, properties });
  }
  return { controls, unsupported };
}

function contentControlMarkers(
  document: Document,
  initialIndex: number,
): [string, string] {
  const occupiedText = document.documentElement.textContent ?? '';
  let index = initialIndex;
  for (let attempts = 0; attempts < 1_000_000; attempts += 1) {
    const start = `__A3S_WORK_CONTENT_CONTROL_START_${index}__`;
    const end = `__A3S_WORK_CONTENT_CONTROL_END_${index}__`;
    if (!occupiedText.includes(start) && !occupiedText.includes(end)) {
      return [start, end];
    }
    index += 1;
  }
  throw new Error('Imported DOCX content-control marker space is exhausted.');
}

export function applyImportedDocxContentControlMarkers(
  document: Document,
  markers: ImportedDocxContentControlMarkers,
): void {
  if (!markers.controls.length) return;
  for (const marker of markers.controls) {
    const wrapper = document.createElement('span');
    for (const [name, value] of Object.entries(
      contentControlDomAttributes(marker.properties),
    )) {
      if (value !== undefined) wrapper.setAttribute(name, value);
    }
    wrapper.className = 'work-document-content-control';
    wrapper.setAttribute('role', 'textbox');
    wrapper.setAttribute(
      'aria-label',
      marker.properties.alias || marker.properties.tag || '内容控件',
    );
    if (
      marker.properties.lock === 'contentLocked' ||
      marker.properties.lock === 'sdtContentLocked'
    ) {
      wrapper.setAttribute('contenteditable', 'false');
    }
    replaceMarkerRange(document.body, marker.start, marker.end, wrapper);
  }
  document.body.normalize();
}

export function hasImportedDocxContentControlMarkers(
  markers: ImportedDocxContentControlMarkers,
): boolean {
  return markers.controls.length > 0;
}

interface ParsedContentControl {
  control: Element;
  content: Element;
  properties: Element;
}

function readContentControl(
  control: Element,
  sequence: number,
): WorkDocumentContentControlProperties | null {
  const definition = readDefinition(control);
  if (!definition) return null;
  const metadata = readProperties(definition.properties);
  if (!metadata) return null;
  const runs = Array.from(definition.content.children);
  if (
    !runs.length ||
    runs.some((run) => !isDocxWordElement(run) || run.localName !== 'r') ||
    hasUnsupportedContentControlSemanticChild(definition.content) ||
    runs.some((run) => !validRun(run))
  ) {
    // A missing run is represented as an empty editable span below.
    if (runs.length) return null;
  }
  const text = runs.map(runText).join('');
  if (text.length > MAX_CONTENT_CONTROL_TEXT) {
    throw new Error('Imported DOCX content control text is too large.');
  }
  if (metadata.type === 'text' && !metadata.multiLine && text.includes('\n')) {
    return null;
  }
  const nativeId = metadata.nativeId;
  return normalizeDocumentContentControlProperties({
    // Native w:id values are expected to be unique, but malformed documents
    // do occur in the wild. Keep the editor identity tied to the source
    // occurrence rather than the untrusted native value so duplicate IDs
    // cannot merge lock/selection state in the browser model.
    id: `docx-content-control-${sequence}`,
    nativeId,
    type: metadata.type,
    alias: metadata.alias,
    tag: metadata.tag,
    lock: metadata.lock,
    multiLine: metadata.multiLine,
    appearance: metadata.appearance,
    color: metadata.color,
  });
}

function readDefinition(control: Element): ParsedContentControl | null {
  if (
    !isDocxWordElement(control) ||
    control.localName !== 'sdt' ||
    !isDirectParagraphControl(control) ||
    hasContentControlRelationshipReference(control) ||
    hasUnsupportedContentControlSemanticChild(control) ||
    !validControlAttributes(control)
  ) {
    return null;
  }
  const children = Array.from(control.children);
  const properties = children.filter(
    (child) => isDocxWordElement(child) && child.localName === 'sdtPr',
  );
  const content = children.filter(
    (child) => isDocxWordElement(child) && child.localName === 'sdtContent',
  );
  const endProperties = children.filter(
    (child) => isDocxWordElement(child) && child.localName === 'sdtEndPr',
  );
  if (
    properties.length !== 1 ||
    content.length !== 1 ||
    endProperties.length > 1 ||
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
  if (
    [control, properties[0], content[0], endProperties[0]].some(
      (element) => element && hasNonWhitespaceText(element),
    )
  ) {
    return null;
  }
  if (
    endProperties[0] &&
    (endProperties[0].children.length > 0 ||
      !validContainerAttributes(endProperties[0]))
  ) {
    return null;
  }
  if (
    !validContainerAttributes(properties[0]) ||
    !validContainerAttributes(content[0])
  ) {
    return null;
  }
  return { control, content: content[0], properties: properties[0] };
}

function readProperties(properties: Element): {
  alias: string;
  appearance: 'boundingBox' | 'hidden' | 'tags';
  color: string | null;
  lock: 'contentLocked' | 'sdtContentLocked' | 'sdtLocked' | 'unlocked';
  multiLine: boolean;
  nativeId: number | null;
  tag: string;
  type: 'richText' | 'text';
} | null {
  if (hasContentControlRelationshipReference(properties)) return null;
  const children = Array.from(properties.children);
  const groups = new Map<string, Element[]>();
  let appearance: 'boundingBox' | 'hidden' | 'tags' = 'boundingBox';
  let color: string | null = null;
  for (const child of children) {
    const namespace = child.namespaceURI ?? '';
    if (isDocxWordElement(child)) {
      if (!PROPERTY_NAMES.has(child.localName)) {
        // These include dataBinding, placeholders, and active form controls.
        return null;
      }
      const matches = groups.get(child.localName) ?? [];
      matches.push(child);
      groups.set(child.localName, matches);
      continue;
    }
    if (namespace === CONTENT_CONTROL_WORD_2012_NAMESPACE) {
      if (child.localName !== 'appearance' && child.localName !== 'color') {
        return null;
      }
      const matches = groups.get(`w15:${child.localName}`) ?? [];
      matches.push(child);
      groups.set(`w15:${child.localName}`, matches);
      continue;
    }
    if (namespace === MARKUP_COMPATIBILITY_NAMESPACE) continue;
    return null;
  }
  if (Array.from(groups.values()).some((items) => items.length > 1))
    return null;
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
  const textElement = groups.get('text')?.[0];
  const richTextElement = groups.get('richText')?.[0];
  if (textElement && richTextElement) return null;
  if (textElement && !validLeaf(textElement, new Set(['multiLine']))) {
    return null;
  }
  if (richTextElement && !validLeaf(richTextElement, new Set())) return null;
  const multilineValue = textElement
    ? wordAttribute(textElement, 'multiLine')
    : null;
  if (
    multilineValue !== null &&
    !['0', '1', 'false', 'off', 'on', 'true'].includes(multilineValue)
  ) {
    return null;
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
    appearance = value as 'boundingBox' | 'hidden' | 'tags';
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
  return {
    alias: alias ?? '',
    appearance,
    color,
    lock: lock as
      | 'contentLocked'
      | 'sdtContentLocked'
      | 'sdtLocked'
      | 'unlocked',
    multiLine:
      multilineValue !== null && ['1', 'on', 'true'].includes(multilineValue),
    nativeId,
    tag: tag ?? '',
    type: textElement ? 'text' : 'richText',
  };
}

function replaceControlWithMarkers(
  document: Document,
  control: Element,
  start: string,
  end: string,
): boolean {
  const paragraph = control.parentElement;
  const content = Array.from(control.children).find(
    (child) => isDocxWordElement(child) && child.localName === 'sdtContent',
  );
  if (!paragraph || !content) return false;
  const startRun = markerRun(document, paragraph, start);
  const endRun = markerRun(document, paragraph, end);
  paragraph.insertBefore(startRun, control);
  const runs = Array.from(content.children);
  for (const run of runs) paragraph.insertBefore(run, control);
  paragraph.insertBefore(endRun, control);
  control.remove();
  return true;
}

function markerRun(
  document: Document,
  context: Element,
  value: string,
): Element {
  const namespace = context.namespaceURI ?? WORD_NAMESPACE;
  const prefix = xmlNamespacePrefix(context, namespace) ?? 'w';
  const run = document.createElementNS(namespace, `${prefix}:r`);
  const text = document.createElementNS(namespace, `${prefix}:t`);
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = value;
  run.append(text);
  return run;
}

function replaceMarkerRange(
  root: HTMLElement,
  start: string,
  end: string,
  replacement: HTMLElement,
): void {
  const startElement = replaceMarkerWithBoundary(root, start, 'start');
  const endElement = replaceMarkerWithBoundary(root, end, 'end');
  if (!startElement || !endElement) return;
  if (startElement.parentNode === endElement.parentNode) {
    const parent = startElement.parentNode;
    let current = startElement.nextSibling;
    while (current && current !== endElement) {
      const next = current.nextSibling;
      replacement.append(current);
      current = next;
    }
    if (current === endElement) {
      startElement.replaceWith(replacement);
      endElement.remove();
      return;
    }
    // The end marker was detached unexpectedly; restore a safe visible tree.
    parent?.append(replacement);
    startElement.remove();
    endElement.remove();
    return;
  }
  const range = root.ownerDocument.createRange();
  try {
    range.setStartAfter(startElement);
    range.setEndBefore(endElement);
    const content = range.extractContents();
    replacement.append(content);
    startElement.replaceWith(replacement);
    endElement.remove();
  } catch {
    // A malformed conversion can place boundaries in different block trees.
    // Remove the sentinels and leave the visible text on Mammoth's normal
    // compatibility path rather than manufacturing a cross-paragraph node.
    startElement.remove();
    endElement.remove();
  }
}

function replaceMarkerWithBoundary(
  root: ParentNode,
  marker: string,
  kind: 'start' | 'end',
): HTMLElement | null {
  const node = textNodes(root).find((candidate) =>
    candidate.data.includes(marker),
  );
  if (!node?.parentNode) return null;
  const ownerDocument = root.ownerDocument;
  if (!ownerDocument) return null;
  const offset = node.data.indexOf(marker);
  const before = node.data.slice(0, offset);
  const after = node.data.slice(offset + marker.length);
  const boundary = ownerDocument.createElement('span');
  boundary.dataset.contentControlBoundary = kind;
  node.parentNode.insertBefore(ownerDocument.createTextNode(before), node);
  node.parentNode.insertBefore(boundary, node);
  if (after) {
    node.parentNode.insertBefore(ownerDocument.createTextNode(after), node);
  }
  node.remove();
  return boundary;
}

function textNodes(root: ParentNode): Text[] {
  const walker = root.ownerDocument?.createTreeWalker(
    root,
    NodeFilter.SHOW_TEXT,
  );
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}

function validControlAttributes(element: Element): boolean {
  return validContainerAttributes(element);
}

function validRun(run: Element): boolean {
  if (hasNonWhitespaceText(run)) return false;
  let sawContent = false;
  let properties = 0;
  for (const child of Array.from(run.children)) {
    if (!isDocxWordElement(child)) return false;
    if (child.localName === 'rPr') {
      properties += 1;
      if (properties > 1 || sawContent) return false;
      if (hasContentControlRelationshipReference(child)) return false;
      continue;
    }
    if (!RUN_CONTENT_NAMES.has(child.localName)) return false;
    sawContent = true;
    if (
      child.localName === 't' &&
      (child.textContent?.length ?? 0) > MAX_CONTENT_CONTROL_TEXT
    ) {
      return false;
    }
  }
  return true;
}

function runText(run: Element): string {
  let value = '';
  for (const child of Array.from(run.children)) {
    if (!isDocxWordElement(child)) continue;
    if (child.localName === 't') value += child.textContent ?? '';
    else if (child.localName === 'tab') value += '\t';
    else if (child.localName === 'br' || child.localName === 'cr')
      value += '\n';
    else if (child.localName === 'noBreakHyphen') value += '\u2011';
    else if (child.localName === 'softHyphen') value += '\u00ad';
  }
  return value;
}

function isDirectParagraphControl(control: Element): boolean {
  return (
    isDocxWordElement(control) &&
    Boolean(
      control.parentElement &&
        isDocxWordElement(control.parentElement) &&
        control.parentElement.localName === 'p',
    )
  );
}

function hasContentControlAncestor(control: Element): boolean {
  let current = control.parentElement;
  while (current) {
    if (isDocxWordElement(current) && current.localName === 'sdt') return true;
    current = current.parentElement;
  }
  return false;
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

function wordAttribute(element: Element, localName: string): string | null {
  const matches = Array.from(element.attributes).filter(
    (item) =>
      xmlAttributeLocalName(item) === localName &&
      xmlAttributeNamespace(element, item) === element.namespaceURI &&
      WORD_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  return matches.length === 1 ? matches[0].value : null;
}

function namespacedAttribute(
  element: Element,
  localName: string,
  namespace: string,
): string | null {
  const matches = Array.from(element.attributes).filter(
    (item) =>
      xmlAttributeLocalName(item) === localName &&
      xmlAttributeNamespace(element, item) === namespace,
  );
  return matches.length === 1 ? matches[0].value : null;
}

function readStringProperty(element: Element | undefined): string | null {
  if (!element) return '';
  if (!validLeaf(element, new Set(['val']))) return null;
  const value = wordAttribute(element, 'val');
  if (value === null || /[\u0000-\u001f\u007f]/u.test(value)) return null;
  return value.slice(0, 255);
}

function validContainerAttributes(element: Element): boolean {
  return Array.from(element.attributes).every((item) => {
    const namespace = xmlAttributeNamespace(element, item);
    return (
      namespace === 'http://www.w3.org/2000/xmlns/' ||
      namespace === MARKUP_COMPATIBILITY_NAMESPACE
    );
  });
}

function hasNonWhitespaceText(element: Element): boolean {
  return Array.from(element.childNodes).some(
    (child) =>
      child.nodeType === Node.TEXT_NODE && Boolean(child.textContent?.trim()),
  );
}

function validLeaf(
  element: Element,
  allowedAttributes: ReadonlySet<string>,
): boolean {
  return (
    element.children.length === 0 &&
    Array.from(element.childNodes).every(
      (child) =>
        child.nodeType !== Node.TEXT_NODE || !(child.textContent ?? '').trim(),
    ) &&
    Array.from(element.attributes).every((item) => {
      const namespace = xmlAttributeNamespace(element, item);
      if (namespace === 'http://www.w3.org/2000/xmlns/') return true;
      if (namespace === MARKUP_COMPATIBILITY_NAMESPACE) return true;
      return (
        namespace === element.namespaceURI &&
        WORD_NAMESPACES.has(element.namespaceURI ?? '') &&
        allowedAttributes.has(xmlAttributeLocalName(item))
      );
    })
  );
}

function validNamespacedLeaf(
  element: Element,
  allowedAttributes: ReadonlySet<string>,
  namespace: string,
): boolean {
  return (
    element.children.length === 0 &&
    Array.from(element.childNodes).every(
      (child) =>
        child.nodeType !== Node.TEXT_NODE || !(child.textContent ?? '').trim(),
    ) &&
    Array.from(element.attributes).every((item) => {
      const attributeNamespace = xmlAttributeNamespace(element, item);
      return (
        attributeNamespace === 'http://www.w3.org/2000/xmlns/' ||
        (attributeNamespace === namespace &&
          allowedAttributes.has(xmlAttributeLocalName(item)))
      );
    })
  );
}

function parseNativeId(value: string | null): number | null {
  if (value === null || !/^-?(0|[1-9][0-9]{0,9})$/u.test(value)) return null;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) &&
    parsed >= -2_147_483_648 &&
    parsed <= 2_147_483_647
    ? parsed
    : null;
}
