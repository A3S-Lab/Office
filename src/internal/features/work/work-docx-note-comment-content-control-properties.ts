import {
  DOCX_WORDPROCESSING_NAMESPACES,
  type DocxIgnorableExtensionPair,
} from './work-docx-ignorable-extension-preservation';
import { isKnownOoxmlNamespace } from './work-docx-note-comment-hyperlink-content';
import { cloneDocxNoteCommentRunProperties } from './work-docx-note-comment-run-properties';
import {
  CONTENT_CONTROL_MC_NAMESPACE as MARKUP_COMPATIBILITY_NAMESPACE,
  CONTENT_CONTROL_RELATIONSHIP_NAMESPACES as RELATIONSHIP_NAMESPACES,
  CONTENT_CONTROL_WORD_2012_NAMESPACE as WORD_2012_NAMESPACE,
  createNamespacedElement as namespacedElement,
  createWordElement as wordElement,
  hasContentControlRelationshipReference as hasRelationshipReference,
  hasOnlyPassiveContentControlAttributes as hasOnlyPassiveAttributes,
  hasUnsupportedContentControlSemanticChild as hasUnsupportedSemanticChild,
  isContentControlSemanticNamespace as isSemanticNamespace,
  isDocxWordElement as isWordElement,
  setNamespacedContentControlAttribute as setNamespacedAttribute,
  setWordContentControlAttribute as setWordAttribute,
  wordDirectChildren,
} from './work-docx-note-comment-content-control-xml';
import {
  XMLNS_NAMESPACE,
  XML_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';

const STATIC_CORE_PROPERTIES = new Set([
  'alias',
  'id',
  'label',
  'lock',
  'placeholder',
  'richText',
  'rPr',
  'tabIndex',
  'tag',
  'temporary',
  'text',
]);
const UNSUPPORTED_CONTROL_TYPES = new Set([
  'bibliography',
  'citation',
  'comboBox',
  'date',
  'docPartList',
  'docPartObj',
  'dropDownList',
  'equation',
  'group',
  'picture',
]);
const LOCK_VALUES = new Set([
  'contentLocked',
  'sdtContentLocked',
  'sdtLocked',
  'unlocked',
]);
const ON_OFF_VALUES = new Set(['0', '1', 'false', 'off', 'on', 'true']);
const APPEARANCE_VALUES = new Set(['boundingBox', 'hidden', 'tags']);
const MAX_STRING_LENGTH = 255;
const MAX_INT32 = 2_147_483_647;
const MIN_INT32 = -2_147_483_648;

export interface StaticContentControlDefinition {
  content: Element;
  control: Element;
  endProperties: Element | null;
  id: string | null;
  multiLine: boolean;
  properties: Element;
  type: 'richText' | 'text';
}

export interface StaticContentControlShell {
  content: Element;
  control: Element;
  pairs: DocxIgnorableExtensionPair[];
}

export function readStaticContentControlDefinition(
  control: Element,
): StaticContentControlDefinition | null {
  if (!isWordElement(control) || control.localName !== 'sdt') return null;
  if (!hasOnlyPassiveAttributes(control)) return null;
  const wordChildren = Array.from(control.children).filter(isWordElement);
  const properties = wordChildren.filter(
    (child) => child.localName === 'sdtPr',
  );
  const endProperties = wordChildren.filter(
    (child) => child.localName === 'sdtEndPr',
  );
  const content = wordChildren.filter(
    (child) => child.localName === 'sdtContent',
  );
  if (
    properties.length !== 1 ||
    endProperties.length > 1 ||
    content.length !== 1 ||
    wordChildren.some(
      (child) =>
        child.localName !== 'sdtPr' &&
        child.localName !== 'sdtEndPr' &&
        child.localName !== 'sdtContent',
    ) ||
    hasUnsupportedSemanticChild(control)
  ) {
    return null;
  }
  const parsed = readProperties(properties[0]);
  if (!parsed || !validEndProperties(endProperties[0] ?? null)) return null;
  return {
    content: content[0],
    control,
    endProperties: endProperties[0] ?? null,
    id: parsed.id,
    multiLine: parsed.multiLine,
    properties: properties[0],
    type: parsed.type,
  };
}

export function createStaticContentControlShell(
  document: Document,
  source: StaticContentControlDefinition,
  assignedId: string | null,
  targetContext: Element,
  kind: 'comment' | 'note',
): StaticContentControlShell | null {
  const namespace = targetContext.namespaceURI;
  if (!namespace || !DOCX_WORDPROCESSING_NAMESPACES.has(namespace)) return null;
  const control = wordElement(document, targetContext, 'sdt');
  const properties = wordElement(document, targetContext, 'sdtPr');
  const content = wordElement(document, targetContext, 'sdtContent');
  const pairs: DocxIgnorableExtensionPair[] = [
    { generated: control, source: source.control },
    { generated: properties, source: source.properties },
    { generated: content, source: source.content },
  ];
  for (const property of Array.from(source.properties.children)) {
    const clone = cloneStaticProperty(
      document,
      property,
      assignedId,
      targetContext,
      kind,
    );
    if (!clone) continue;
    properties.append(clone);
    pairs.push({ generated: clone, source: property });
  }
  control.append(properties);
  if (source.endProperties) {
    const endProperties = wordElement(document, targetContext, 'sdtEndPr');
    const sourceRunProperties = wordDirectChildren(
      source.endProperties,
      'rPr',
    )[0];
    if (sourceRunProperties) {
      const runProperties = cloneDocxNoteCommentRunProperties(
        document,
        sourceRunProperties,
        targetContext,
        kind,
      );
      if (runProperties) {
        endProperties.append(runProperties);
        pairs.push({
          generated: runProperties,
          source: sourceRunProperties,
        });
      }
    }
    control.append(endProperties);
    pairs.push({ generated: endProperties, source: source.endProperties });
  }
  control.append(content);
  return { content, control, pairs };
}

interface ParsedStaticProperties {
  id: string | null;
  multiLine: boolean;
  type: 'richText' | 'text';
}

function readProperties(properties: Element): ParsedStaticProperties | null {
  if (!hasOnlyPassiveAttributes(properties)) return null;
  const groups = new Map<string, Element[]>();
  let type: 'richText' | 'text' = 'richText';
  let typeCount = 0;
  for (const child of Array.from(properties.children)) {
    const namespace = child.namespaceURI ?? '';
    if (isWordElement(child)) {
      if (
        child.localName === 'dataBinding' ||
        child.localName === 'showingPlcHdr' ||
        UNSUPPORTED_CONTROL_TYPES.has(child.localName) ||
        !STATIC_CORE_PROPERTIES.has(child.localName)
      ) {
        return null;
      }
      const matches = groups.get(child.localName) ?? [];
      matches.push(child);
      groups.set(child.localName, matches);
      if (child.localName === 'richText' || child.localName === 'text') {
        type = child.localName;
        typeCount += 1;
      }
      continue;
    }
    if (namespace === WORD_2012_NAMESPACE) {
      if (child.localName !== 'appearance' && child.localName !== 'color') {
        return null;
      }
      const key = `{${namespace}}${child.localName}`;
      const matches = groups.get(key) ?? [];
      matches.push(child);
      groups.set(key, matches);
      continue;
    }
    if (isSemanticNamespace(namespace)) return null;
  }
  if (
    typeCount > 1 ||
    Array.from(groups.values()).some((matches) => matches.length > 1)
  ) {
    return null;
  }
  for (const [key, matches] of groups) {
    if (!validStaticProperty(key, matches[0])) return null;
  }
  const idElement = groups.get('id')?.[0];
  const id = idElement ? normalizeInt32(wordAttribute(idElement, 'val')) : null;
  if (idElement && id === null) return null;
  const text = groups.get('text')?.[0];
  const multiLineValue = text ? wordAttribute(text, 'multiLine') : null;
  return {
    id,
    multiLine: Boolean(
      multiLineValue && ['1', 'on', 'true'].includes(multiLineValue),
    ),
    type,
  };
}

function validStaticProperty(key: string, property: Element): boolean {
  if (key === 'rPr') return !hasRelationshipReference(property);
  if (key === 'placeholder') return !hasRelationshipReference(property);
  if (key === 'temporary') return validLeaf(property, new Set());
  if (key === 'richText') return validLeaf(property, new Set());
  if (key === 'text') {
    const value = wordAttribute(property, 'multiLine');
    return (
      validLeaf(property, new Set(['multiLine'])) &&
      (value === null || ON_OFF_VALUES.has(value))
    );
  }
  if (key === 'alias' || key === 'tag') {
    const value = wordAttribute(property, 'val');
    return Boolean(
      validLeaf(property, new Set(['val'])) &&
        value !== null &&
        value.length <= MAX_STRING_LENGTH &&
        !/[\u0000-\u001f\u007f]/u.test(value),
    );
  }
  if (key === 'id') {
    return (
      validLeaf(property, new Set(['val'])) &&
      normalizeInt32(wordAttribute(property, 'val')) !== null
    );
  }
  if (key === 'label' || key === 'tabIndex') {
    const value = wordAttribute(property, 'val');
    return (
      validLeaf(property, new Set(['val'])) &&
      value !== null &&
      /^(0|[1-9][0-9]{0,9})$/u.test(value) &&
      Number(value) <= MAX_INT32
    );
  }
  if (key === 'lock') {
    const value = wordAttribute(property, 'val');
    return (
      validLeaf(property, new Set(['val'])) &&
      value !== null &&
      LOCK_VALUES.has(value)
    );
  }
  if (key === `{${WORD_2012_NAMESPACE}}appearance`) {
    const value = namespacedAttribute(property, 'val', WORD_2012_NAMESPACE);
    return (
      validLeaf(property, new Set(['val']), WORD_2012_NAMESPACE) &&
      value !== null &&
      APPEARANCE_VALUES.has(value)
    );
  }
  if (key === `{${WORD_2012_NAMESPACE}}color`) {
    const value = namespacedAttribute(property, 'val', WORD_2012_NAMESPACE);
    return (
      validLeaf(property, new Set(['val']), WORD_2012_NAMESPACE) &&
      Boolean(value && /^[0-9A-Fa-f]{6}$/u.test(value))
    );
  }
  return false;
}

function cloneStaticProperty(
  document: Document,
  source: Element,
  assignedId: string | null,
  targetContext: Element,
  kind: 'comment' | 'note',
): Element | null {
  if (isWordElement(source)) {
    if (
      source.localName === 'placeholder' ||
      source.localName === 'temporary'
    ) {
      return null;
    }
    if (source.localName === 'rPr') {
      return cloneDocxNoteCommentRunProperties(
        document,
        source,
        targetContext,
        kind,
      );
    }
    const clone = wordElement(document, targetContext, source.localName);
    for (const item of Array.from(source.attributes)) {
      const namespace = xmlAttributeNamespace(source, item);
      const localName = xmlAttributeLocalName(item);
      if (DOCX_WORDPROCESSING_NAMESPACES.has(namespace ?? '')) {
        const value = source.localName === 'id' ? assignedId : item.value;
        if (value !== null)
          setWordAttribute(clone, targetContext, localName, value);
      } else if (namespace === XML_NAMESPACE) {
        clone.setAttributeNS(XML_NAMESPACE, `xml:${localName}`, item.value);
      } else if (!namespace) {
        clone.setAttribute(localName, item.value);
      }
    }
    return clone;
  }
  if (source.namespaceURI === WORD_2012_NAMESPACE) {
    const clone = namespacedElement(
      document,
      targetContext,
      WORD_2012_NAMESPACE,
      'w15',
      source.localName,
    );
    const value = namespacedAttribute(source, 'val', WORD_2012_NAMESPACE);
    if (value !== null) {
      setNamespacedAttribute(
        clone,
        targetContext,
        WORD_2012_NAMESPACE,
        'w15',
        'val',
        value,
      );
    }
    return clone;
  }
  return null;
}

function validEndProperties(element: Element | null): boolean {
  if (!element) return true;
  if (
    !hasOnlyPassiveAttributes(element) ||
    hasUnsupportedSemanticChild(element)
  ) {
    return false;
  }
  const wordChildren = Array.from(element.children).filter(isWordElement);
  return (
    wordChildren.length <= 1 &&
    wordChildren.every(
      (child) => child.localName === 'rPr' && !hasRelationshipReference(child),
    )
  );
}

function validLeaf(
  element: Element,
  allowedAttributes: ReadonlySet<string>,
  namespace = element.namespaceURI ?? '',
): boolean {
  return (
    element.children.length === 0 &&
    Array.from(element.childNodes).every(
      (child) =>
        child.nodeType !== Node.TEXT_NODE || !(child.textContent ?? '').trim(),
    ) &&
    Array.from(element.attributes).every((item) => {
      const itemNamespace = xmlAttributeNamespace(element, item);
      if (itemNamespace === XMLNS_NAMESPACE) return true;
      if (itemNamespace === namespace) {
        return allowedAttributes.has(xmlAttributeLocalName(item));
      }
      return Boolean(
        itemNamespace &&
          itemNamespace !== MARKUP_COMPATIBILITY_NAMESPACE &&
          !RELATIONSHIP_NAMESPACES.has(itemNamespace) &&
          !isKnownOoxmlNamespace(itemNamespace),
      );
    })
  );
}

function normalizeInt32(value: string | null): string | null {
  if (!value || !/^-?(0|[1-9][0-9]{0,9})$/u.test(value)) return null;
  const number = Number(value);
  return Number.isSafeInteger(number) &&
    number >= MIN_INT32 &&
    number <= MAX_INT32
    ? String(number)
    : null;
}

function wordAttribute(element: Element, localName: string): string | null {
  return uniqueAttribute(element, localName, DOCX_WORDPROCESSING_NAMESPACES);
}

function namespacedAttribute(
  element: Element,
  localName: string,
  namespace: string,
): string | null {
  return uniqueAttribute(element, localName, new Set([namespace]));
}

function uniqueAttribute(
  element: Element,
  localName: string,
  namespaces: ReadonlySet<string>,
): string | null {
  const matches = Array.from(element.attributes).filter(
    (item) =>
      xmlAttributeLocalName(item) === localName &&
      namespaces.has(xmlAttributeNamespace(element, item) ?? ''),
  );
  return matches.length === 1 ? matches[0].value : null;
}
