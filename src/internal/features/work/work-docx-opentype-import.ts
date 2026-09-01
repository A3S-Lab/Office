import {
  normalizeDocumentOpenTypeLigatures,
  normalizeDocumentOpenTypeNumberForm,
  normalizeDocumentOpenTypeNumberSpacing,
  normalizeDocumentOpenTypeStylisticSets,
  type WorkDocumentOpenTypeFeatures,
} from './work-document-opentype';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { directChildren } from './work-ooxml-package';

export interface ResolvedDocxOpenTypeFeatures {
  features: WorkDocumentOpenTypeFeatures | null;
  invalidCount: number;
  spoofedCount: number;
}

export const DOCX_WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';

const OPEN_TYPE_PROPERTY_NAMES = [
  'ligatures',
  'numForm',
  'numSpacing',
  'stylisticSets',
  'cntxtAlts',
] as const;
type OpenTypePropertyName = (typeof OPEN_TYPE_PROPERTY_NAMES)[number];
const OPEN_TYPE_PROPERTY_NAME_SET = new Set<string>(OPEN_TYPE_PROPERTY_NAMES);
const MAX_STYLISTIC_SET_ENTRIES = 4_096;

export function resolveDocxOpenTypeFeatures(
  propertySources: readonly Element[],
): ResolvedDocxOpenTypeFeatures {
  const features: WorkDocumentOpenTypeFeatures = {};
  let invalidCount = 0;
  let spoofedCount = 0;
  for (const properties of propertySources) {
    const candidates = new Map<OpenTypePropertyName, Element[]>();
    for (const child of directChildren(properties)) {
      if (!OPEN_TYPE_PROPERTY_NAME_SET.has(child.localName)) continue;
      if (child.namespaceURI !== DOCX_WORD_2010_NAMESPACE) {
        spoofedCount += 1;
        continue;
      }
      const name = child.localName as OpenTypePropertyName;
      const values = candidates.get(name) ?? [];
      values.push(child);
      candidates.set(name, values);
    }
    for (const name of OPEN_TYPE_PROPERTY_NAMES) {
      const values = candidates.get(name) ?? [];
      if (!values.length) continue;
      if (values.length !== 1 || !values[0]) {
        invalidCount += 1;
        continue;
      }
      const parsed = parseProperty(name, values[0]);
      if (parsed === null) {
        invalidCount += 1;
        continue;
      }
      if (parsed.name === 'ligatures') features.ligatures = parsed.value;
      else if (parsed.name === 'numForm') features.numberForm = parsed.value;
      else if (parsed.name === 'numSpacing') {
        features.numberSpacing = parsed.value;
      } else if (parsed.name === 'stylisticSets') {
        features.stylisticSets = parsed.value;
      } else features.contextualAlternates = parsed.value;
    }
  }
  return {
    features:
      invalidCount || spoofedCount || !Object.keys(features).length
        ? null
        : features,
    invalidCount,
    spoofedCount,
  };
}

type ParsedOpenTypeProperty =
  | {
      name: 'ligatures';
      value: NonNullable<WorkDocumentOpenTypeFeatures['ligatures']>;
    }
  | {
      name: 'numForm';
      value: NonNullable<WorkDocumentOpenTypeFeatures['numberForm']>;
    }
  | {
      name: 'numSpacing';
      value: NonNullable<WorkDocumentOpenTypeFeatures['numberSpacing']>;
    }
  | { name: 'stylisticSets'; value: number[] }
  | { name: 'cntxtAlts'; value: boolean };

function parseProperty(
  name: OpenTypePropertyName,
  element: Element,
): ParsedOpenTypeProperty | null {
  if (name === 'ligatures') {
    const attributes = word2010LeafAttributes(element, new Set(['val']));
    const value =
      attributes?.size === 1
        ? normalizeDocumentOpenTypeLigatures(attributes.get('val'))
        : null;
    return value === null ? null : { name, value };
  }
  if (name === 'numForm') {
    const attributes = word2010LeafAttributes(element, new Set(['val']));
    const value =
      attributes?.size === 1
        ? normalizeDocumentOpenTypeNumberForm(attributes.get('val'))
        : null;
    return value === null ? null : { name, value };
  }
  if (name === 'numSpacing') {
    const attributes = word2010LeafAttributes(element, new Set(['val']));
    const value =
      attributes?.size === 1
        ? normalizeDocumentOpenTypeNumberSpacing(attributes.get('val'))
        : null;
    return value === null ? null : { name, value };
  }
  if (name === 'cntxtAlts') {
    const attributes = word2010LeafAttributes(element, new Set(['val']));
    if (!attributes) return null;
    const value = attributes.has('val')
      ? strictOnOff(attributes.get('val'))
      : true;
    return value === null ? null : { name, value };
  }
  const value = parseStylisticSets(element);
  return value === null ? null : { name, value };
}

function parseStylisticSets(element: Element): number[] | null {
  const attributes = word2010Attributes(element, new Set());
  if (!attributes || attributes.size || hasNonWhitespaceText(element)) {
    return null;
  }
  const children = directChildren(element);
  if (children.length > MAX_STYLISTIC_SET_ENTRIES) return null;
  const raw: number[] = [];
  for (const child of children) {
    if (
      child.localName !== 'styleSet' ||
      child.namespaceURI !== DOCX_WORD_2010_NAMESPACE
    ) {
      return null;
    }
    const childAttributes = word2010LeafAttributes(
      child,
      new Set(['id', 'val']),
    );
    if (!childAttributes?.has('id')) return null;
    const id = boundedInteger(childAttributes.get('id'), 1, 20);
    const enabled = childAttributes.has('val')
      ? strictOnOff(childAttributes.get('val'))
      : true;
    if (id === null || enabled === null) return null;
    if (enabled) raw.push(id);
  }
  return normalizeDocumentOpenTypeStylisticSets(raw);
}

function word2010LeafAttributes(
  element: Element,
  allowed: ReadonlySet<string>,
): Map<string, string> | null {
  if (directChildren(element).length || hasNonWhitespaceText(element)) {
    return null;
  }
  return word2010Attributes(element, allowed);
}

function word2010Attributes(
  element: Element,
  allowed: ReadonlySet<string>,
): Map<string, string> | null {
  const result = new Map<string, string>();
  for (const attribute of Array.from(element.attributes)) {
    if (attribute.namespaceURI === XMLNS_NAMESPACE) continue;
    const name = xmlAttributeLocalName(attribute);
    if (
      !allowed.has(name) ||
      xmlAttributeNamespace(element, attribute) !== DOCX_WORD_2010_NAMESPACE ||
      result.has(name)
    ) {
      return null;
    }
    result.set(name, attribute.value);
  }
  return result;
}

function hasNonWhitespaceText(element: Element): boolean {
  return Array.from(element.childNodes).some(
    (node) => node.nodeType !== 1 && (node.textContent ?? '').trim(),
  );
}

function strictOnOff(value: string | undefined): boolean | null {
  if (value === '1' || value === 'true') return true;
  if (value === '0' || value === 'false') return false;
  return null;
}

function boundedInteger(
  value: string | undefined,
  minimum: number,
  maximum: number,
): number | null {
  const source = value?.trim() ?? '';
  if (!/^[+-]?\d+$/.test(source)) return null;
  const number = Number(source);
  return Number.isSafeInteger(number) && number >= minimum && number <= maximum
    ? number
    : null;
}
