import {
  normalizeDocumentLanguageTag,
  normalizeDocumentProofingLanguages,
  type WorkDocumentProofingLanguages,
} from './work-document-proofing';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { directChildren } from './work-ooxml-package';

export interface DocxProofingLanguageOptions {
  value?: string;
  eastAsia?: string;
  bidirectional?: string;
}

export type DocxProofingLanguagesInspection =
  | { status: 'absent'; spoofedCount: number }
  | { status: 'invalid'; spoofedCount: number }
  | {
      status: 'valid';
      value: WorkDocumentProofingLanguages;
      spoofedCount: number;
    };

export type DocxNoProofInspection =
  | { status: 'absent'; spoofedCount: number }
  | { status: 'invalid'; spoofedCount: number }
  | { status: 'valid'; value: boolean; spoofedCount: number };

export interface ResolvedDocxProofing {
  languages: WorkDocumentProofingLanguages | undefined;
  noProof: boolean | undefined;
  invalidCount: number;
  spoofedCount: number;
}

const LANGUAGE_ATTRIBUTES = new Map<
  string,
  keyof WorkDocumentProofingLanguages
>([
  ['val', 'latin'],
  ['eastAsia', 'eastAsia'],
  ['bidi', 'bidi'],
] as const);

export function inspectDocxProofingLanguages(
  properties: Element | null | undefined,
): DocxProofingLanguagesInspection {
  if (!properties) return { status: 'absent', spoofedCount: 0 };
  const matches = directChildren(properties, 'lang');
  if (!matches.length) return { status: 'absent', spoofedCount: 0 };
  const native = matches.filter((element) => isWordElement(element));
  const spoofedCount = matches.length - native.length;
  if (matches.length !== 1 || native.length !== 1) {
    return { status: 'invalid', spoofedCount };
  }
  const element = native[0];
  if (!element || element.childNodes.length) {
    return { status: 'invalid', spoofedCount };
  }
  const languages: WorkDocumentProofingLanguages = {};
  for (const attribute of Array.from(element.attributes)) {
    if (isNamespaceDeclaration(attribute)) continue;
    const name = wordAttributeLocalName(element, attribute);
    const slot = name ? LANGUAGE_ATTRIBUTES.get(name) : undefined;
    if (!slot) {
      return { status: 'invalid', spoofedCount };
    }
    const language = normalizeDocumentLanguageTag(attribute.value);
    if (!language) return { status: 'invalid', spoofedCount };
    languages[slot] = language;
  }
  return Object.keys(languages).length
    ? { status: 'valid', value: languages, spoofedCount }
    : { status: 'invalid', spoofedCount };
}

export function inspectDocxNoProof(
  properties: Element | null | undefined,
): DocxNoProofInspection {
  if (!properties) return { status: 'absent', spoofedCount: 0 };
  const matches = directChildren(properties, 'noProof');
  if (!matches.length) return { status: 'absent', spoofedCount: 0 };
  const native = matches.filter((element) => isWordElement(element));
  const spoofedCount = matches.length - native.length;
  if (matches.length !== 1 || native.length !== 1) {
    return { status: 'invalid', spoofedCount };
  }
  const element = native[0];
  if (!element || element.childNodes.length) {
    return { status: 'invalid', spoofedCount };
  }
  const attributes = Array.from(element.attributes).filter(
    (attribute) => !isNamespaceDeclaration(attribute),
  );
  if (
    attributes.length > 1 ||
    attributes.some(
      (attribute) => wordAttributeLocalName(element, attribute) !== 'val',
    )
  ) {
    return { status: 'invalid', spoofedCount };
  }
  if (!attributes.length) {
    return { status: 'valid', value: true, spoofedCount };
  }
  const value = onOffValue(attributes[0]?.value);
  return value === undefined
    ? { status: 'invalid', spoofedCount }
    : { status: 'valid', value, spoofedCount };
}

export function resolveDocxProofing(
  propertySources: readonly Element[],
): ResolvedDocxProofing {
  let languages: WorkDocumentProofingLanguages | undefined;
  let noProof: boolean | undefined;
  let invalidCount = 0;
  let spoofedCount = 0;
  for (const properties of propertySources) {
    const languageInspection = inspectDocxProofingLanguages(properties);
    spoofedCount += languageInspection.spoofedCount;
    if (languageInspection.status === 'valid') {
      languages = { ...languages, ...languageInspection.value };
    } else if (languageInspection.status === 'invalid') {
      invalidCount += 1;
    }

    const noProofInspection = inspectDocxNoProof(properties);
    spoofedCount += noProofInspection.spoofedCount;
    if (noProofInspection.status === 'valid') {
      noProof = noProofInspection.value;
    } else if (noProofInspection.status === 'invalid') {
      invalidCount += 1;
    }
  }
  return { languages, noProof, invalidCount, spoofedCount };
}

export function documentProofingLanguageDocxOptions(
  source: unknown,
): DocxProofingLanguageOptions | undefined {
  const languages = normalizeDocumentProofingLanguages(source);
  if (!languages) return undefined;
  return {
    ...(languages.latin ? { value: languages.latin } : {}),
    ...(languages.eastAsia ? { eastAsia: languages.eastAsia } : {}),
    ...(languages.bidi ? { bidirectional: languages.bidi } : {}),
  };
}

function isWordElement(element: Element): boolean {
  return DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '');
}

function isNamespaceDeclaration(attribute: Attr): boolean {
  return (
    attribute.name === 'xmlns' ||
    attribute.prefix === 'xmlns' ||
    attribute.namespaceURI === 'http://www.w3.org/2000/xmlns/'
  );
}

function wordAttributeLocalName(
  element: Element,
  attribute: Attr,
): string | null {
  if (attribute.namespaceURI === element.namespaceURI) {
    return attribute.localName;
  }
  const prefix = element.prefix ?? element.tagName.split(':')[0];
  const lexicalPrefix = prefix ? `${prefix}:` : '';
  return attribute.namespaceURI === null &&
    lexicalPrefix &&
    attribute.name.startsWith(lexicalPrefix)
    ? attribute.name.slice(lexicalPrefix.length)
    : null;
}

function onOffValue(source: string | undefined): boolean | undefined {
  const value = source?.trim().toLowerCase();
  if (value === '1' || value === 'true' || value === 'on') return true;
  if (value === '0' || value === 'false' || value === 'off') return false;
  return undefined;
}
