import type JSZip from 'jszip';
import { attribute, directChildren, parseXml } from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';
import {
  assertXmlRoot,
  cloneXmlElement,
  collectRetainedNamespaces,
  declareInheritedNamespaces,
  hasXmlAttribute,
  isStructurallyValidAlternateContent,
  XMLNS_NAMESPACE,
  XML_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
  xmlAttributePrefix,
  xmlDeclaredPrefix,
  xmlNamespaceUri,
} from './work-docx-settings-xml';

const SETTINGS_PATH = 'word/settings.xml';
const WORDPROCESSING_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://purl.oclc.org/ooxml/wordprocessingml/main',
]);
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
]);
const MAX_EXTENSION_ELEMENTS = 4_096;
const MAX_EXTENSION_DEPTH = 64;
const SOURCE_SETTING_DENYLIST = new Set([
  'attachedTemplate',
  'documentProtection',
  'mailMerge',
  'trackRevisions',
  'updateFields',
  'writeProtection',
]);

interface IgnorableNamespace {
  prefix: string;
  namespace: string;
}

interface SettingsMergeContext {
  generatedDocument: Document;
  generatedRoot: Element;
  sourceRoot: Element;
  sourceIgnorable: readonly IgnorableNamespace[];
  sourceIgnorableNamespaces: ReadonlySet<string>;
  retainedNamespaces: Set<string>;
  nextExtensionPrefix: number;
}

/**
 * Merges passive extension markup from source settings into generated settings.
 *
 * WordprocessingML settings and relationship-bound markup remain generated-only.
 * The source contribution is limited to ignorable extension attributes/elements
 * and structurally valid, non-conflicting mc:AlternateContent blocks.
 */
export async function preserveDocxSettingsExtensions(
  generated: JSZip,
  source: JSZip,
  generatedPath = SETTINGS_PATH,
  sourcePath = SETTINGS_PATH,
): Promise<void> {
  const generatedEntry = generated.file(generatedPath);
  const sourceEntry = source.file(sourcePath);
  if (!generatedEntry || !sourceEntry) return;

  const [generatedBytes, sourceBytes] = await Promise.all([
    generatedEntry.async('uint8array'),
    sourceEntry.async('uint8array'),
  ]);
  const generatedDocument = parseXml(
    decodeXmlBytes(generatedBytes, `generated DOCX ${generatedPath}`),
    `generated DOCX ${generatedPath}`,
  );
  const sourceDocument = parseXml(
    decodeXmlBytes(sourceBytes, `source DOCX ${sourcePath}`),
    `source DOCX ${sourcePath}`,
  );
  assertXmlRoot(
    generatedDocument.documentElement,
    'settings',
    WORDPROCESSING_NAMESPACES,
    'Generated DOCX word/settings.xml is not WordprocessingML settings.',
  );
  assertXmlRoot(
    sourceDocument.documentElement,
    'settings',
    WORDPROCESSING_NAMESPACES,
    'Registered source DOCX word/settings.xml is not WordprocessingML settings.',
  );

  const sourceIgnorable = ignorableNamespaces(sourceDocument.documentElement);
  const context: SettingsMergeContext = {
    generatedDocument,
    generatedRoot: generatedDocument.documentElement,
    sourceRoot: sourceDocument.documentElement,
    sourceIgnorable,
    sourceIgnorableNamespaces: new Set(
      sourceIgnorable.map((item) => item.namespace),
    ),
    retainedNamespaces: new Set(),
    nextExtensionPrefix: 1,
  };
  mergeKnownElement(
    generatedDocument.documentElement,
    sourceDocument.documentElement,
    context,
    0,
  );
  mergeIgnorableNamespaces(context);
  mergeCompatibilityQNameRules(context);
  generated.file(generatedPath, serializeUtf8Xml(generatedDocument));
}

function mergeKnownElement(
  generated: Element,
  source: Element,
  context: SettingsMergeContext,
  depth: number,
): void {
  mergePassiveAttributes(generated, source, context);
  if (depth >= MAX_EXTENSION_DEPTH) return;

  const generatedChildren = directChildren(generated);
  const sourceChildren = directChildren(source);
  const generatedGroups = groupChildren(generatedChildren);
  const sourceGroups = groupChildren(sourceChildren);

  for (const sourceChild of sourceChildren) {
    const key = semanticElementKey(sourceChild);
    const generatedMatches = generatedGroups.get(key) ?? [];
    const sourceMatches = sourceGroups.get(key) ?? [];
    if (generatedMatches.length === 1 && sourceMatches.length === 1) {
      mergeKnownElement(generatedMatches[0], sourceChild, context, depth + 1);
    }
  }

  const generatedSemanticNames = new Set(
    generatedChildren
      .filter((child) => isWordprocessingNamespace(child.namespaceURI))
      .map((child) => child.localName),
  );
  for (const [index, sourceChild] of sourceChildren.entries()) {
    const key = semanticElementKey(sourceChild);
    if ((generatedGroups.get(key)?.length ?? 0) > 0) continue;
    if (!isExtensionCandidate(sourceChild, context)) continue;
    if (!isSafeExtensionSubtree(sourceChild, generatedSemanticNames)) continue;

    const imported = cloneXmlElement(context.generatedDocument, sourceChild);
    declareInheritedNamespaces(imported, sourceChild);
    const anchor = followingGeneratedAnchor(
      sourceChildren,
      index,
      generatedGroups,
      sourceGroups,
    );
    generated.insertBefore(imported, anchor);
    collectRetainedNamespaces(sourceChild, context.retainedNamespaces);
  }
}

function mergePassiveAttributes(
  generated: Element,
  source: Element,
  context: SettingsMergeContext,
): void {
  for (const sourceAttribute of Array.from(source.attributes)) {
    const namespace = xmlAttributeNamespace(source, sourceAttribute);
    const localName = xmlAttributeLocalName(sourceAttribute);
    if (
      !namespace ||
      namespace === XMLNS_NAMESPACE ||
      namespace === XML_NAMESPACE ||
      namespace === MARKUP_COMPATIBILITY_NAMESPACE ||
      isWordprocessingNamespace(namespace) ||
      RELATIONSHIP_NAMESPACES.has(namespace) ||
      !context.sourceIgnorableNamespaces.has(namespace) ||
      hasXmlAttribute(generated, namespace, localName)
    ) {
      continue;
    }
    const prefix = ensureRootNamespace(
      context,
      xmlAttributePrefix(sourceAttribute),
      namespace,
    );
    generated.setAttributeNS(
      namespace,
      `${prefix}:${localName}`,
      sourceAttribute.value,
    );
    context.retainedNamespaces.add(namespace);
  }
}

function isExtensionCandidate(
  element: Element,
  context: SettingsMergeContext,
): boolean {
  const namespace = element.namespaceURI;
  if (!namespace || isWordprocessingNamespace(namespace)) return false;
  if (namespace === MARKUP_COMPATIBILITY_NAMESPACE) {
    return element.localName === 'AlternateContent';
  }
  return (
    namespace !== XMLNS_NAMESPACE &&
    namespace !== XML_NAMESPACE &&
    !RELATIONSHIP_NAMESPACES.has(namespace) &&
    context.sourceIgnorableNamespaces.has(namespace)
  );
}

function isSafeExtensionSubtree(
  root: Element,
  generatedSemanticNames: ReadonlySet<string>,
): boolean {
  const stack: Array<{ element: Element; depth: number }> = [
    { element: root, depth: 1 },
  ];
  let elementCount = 0;
  let alternateContentCount = 0;
  while (stack.length) {
    const current = stack.pop();
    if (!current) break;
    elementCount += 1;
    if (
      elementCount > MAX_EXTENSION_ELEMENTS ||
      current.depth > MAX_EXTENSION_DEPTH ||
      hasUnsafeExtensionAttribute(current.element)
    ) {
      return false;
    }
    if (isWordprocessingNamespace(current.element.namespaceURI)) {
      if (
        SOURCE_SETTING_DENYLIST.has(current.element.localName) ||
        generatedSemanticNames.has(current.element.localName)
      ) {
        return false;
      }
    }
    if (current.element.namespaceURI === MARKUP_COMPATIBILITY_NAMESPACE) {
      if (current.element.localName === 'AlternateContent') {
        alternateContentCount += 1;
        if (
          alternateContentCount > 1 ||
          !isStructurallyValidAlternateContent(current.element)
        ) {
          return false;
        }
      } else if (
        current.element.localName !== 'Choice' &&
        current.element.localName !== 'Fallback'
      ) {
        return false;
      }
    }
    for (const child of directChildren(current.element)) {
      stack.push({ element: child, depth: current.depth + 1 });
    }
  }
  return true;
}

function hasUnsafeExtensionAttribute(element: Element): boolean {
  return Array.from(element.attributes).some((item) => {
    const namespace = xmlAttributeNamespace(element, item) ?? '';
    return (
      RELATIONSHIP_NAMESPACES.has(namespace) ||
      (namespace === MARKUP_COMPATIBILITY_NAMESPACE &&
        xmlAttributeLocalName(item) === 'MustUnderstand')
    );
  });
}

function followingGeneratedAnchor(
  sourceChildren: readonly Element[],
  sourceIndex: number,
  generatedGroups: ReadonlyMap<string, readonly Element[]>,
  sourceGroups: ReadonlyMap<string, readonly Element[]>,
): Element | null {
  for (let index = sourceIndex + 1; index < sourceChildren.length; index += 1) {
    const key = semanticElementKey(sourceChildren[index]);
    const generatedMatches = generatedGroups.get(key) ?? [];
    if (
      generatedMatches.length === 1 &&
      (sourceGroups.get(key)?.length ?? 0) === 1
    ) {
      return generatedMatches[0];
    }
  }
  return null;
}

function groupChildren(children: readonly Element[]): Map<string, Element[]> {
  const groups = new Map<string, Element[]>();
  for (const child of children) {
    const key = semanticElementKey(child);
    const existing = groups.get(key);
    if (existing) existing.push(child);
    else groups.set(key, [child]);
  }
  return groups;
}

function semanticElementKey(element: Element): string {
  return isWordprocessingNamespace(element.namespaceURI)
    ? `word:${element.localName}`
    : `{${element.namespaceURI ?? ''}}${element.localName}`;
}

function ignorableNamespaces(root: Element): IgnorableNamespace[] {
  const result: IgnorableNamespace[] = [];
  const seen = new Set<string>();
  for (const prefix of (attribute(root, 'mc:Ignorable') ?? '')
    .trim()
    .split(/\s+/)
    .filter(Boolean)) {
    const namespace = xmlNamespaceUri(root, prefix);
    if (!namespace || seen.has(namespace)) continue;
    seen.add(namespace);
    result.push({ prefix, namespace });
  }
  return result;
}

function mergeIgnorableNamespaces(context: SettingsMergeContext): void {
  const tokens = (attribute(context.generatedRoot, 'mc:Ignorable') ?? '')
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  const usedTokens = new Set(tokens);
  const coveredNamespaces = new Set(
    tokens.flatMap((prefix) => {
      const namespace = xmlNamespaceUri(context.generatedRoot, prefix);
      return namespace ? [namespace] : [];
    }),
  );
  for (const item of context.sourceIgnorable) {
    if (
      !context.retainedNamespaces.has(item.namespace) ||
      coveredNamespaces.has(item.namespace)
    ) {
      continue;
    }
    const prefix = ensureRootNamespace(context, item.prefix, item.namespace);
    if (!usedTokens.has(prefix)) {
      tokens.push(prefix);
      usedTokens.add(prefix);
    }
    coveredNamespaces.add(item.namespace);
  }
  if (!tokens.length) return;
  const prefix = ensureRootNamespace(
    context,
    'mc',
    MARKUP_COMPATIBILITY_NAMESPACE,
  );
  context.generatedRoot.setAttributeNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    `${prefix}:Ignorable`,
    tokens.join(' '),
  );
}

function mergeCompatibilityQNameRules(context: SettingsMergeContext): void {
  for (const name of [
    'PreserveAttributes',
    'PreserveElements',
    'ProcessContent',
  ]) {
    const tokens = (attribute(context.generatedRoot, `mc:${name}`) ?? '')
      .trim()
      .split(/\s+/)
      .filter(Boolean);
    const used = new Set(tokens);
    for (const sourceToken of (
      attribute(context.sourceRoot, `mc:${name}`) ?? ''
    )
      .trim()
      .split(/\s+/)
      .filter(Boolean)) {
      const match =
        /^([A-Za-z_][A-Za-z0-9_.-]*):([A-Za-z_][A-Za-z0-9_.-]*|\*)$/.exec(
          sourceToken,
        );
      const namespace = match
        ? xmlNamespaceUri(context.sourceRoot, match[1])
        : null;
      if (!match || !namespace || !context.retainedNamespaces.has(namespace)) {
        continue;
      }
      const prefix = ensureRootNamespace(context, match[1], namespace);
      const token = `${prefix}:${match[2]}`;
      if (!used.has(token)) tokens.push(token);
      used.add(token);
    }
    if (tokens.length) {
      const prefix = ensureRootNamespace(
        context,
        'mc',
        MARKUP_COMPATIBILITY_NAMESPACE,
      );
      context.generatedRoot.setAttributeNS(
        MARKUP_COMPATIBILITY_NAMESPACE,
        `${prefix}:${name}`,
        tokens.join(' '),
      );
    }
  }
}

function ensureRootNamespace(
  context: SettingsMergeContext,
  preferredPrefix: string | null,
  namespace: string,
): string {
  const existing = xmlDeclaredPrefix(context.generatedRoot, namespace);
  if (existing) return existing;
  if (
    preferredPrefix &&
    isNamespacePrefix(preferredPrefix) &&
    !xmlNamespaceUri(context.generatedRoot, preferredPrefix)
  ) {
    context.generatedRoot.setAttributeNS(
      XMLNS_NAMESPACE,
      `xmlns:${preferredPrefix}`,
      namespace,
    );
    return preferredPrefix;
  }
  let prefix: string;
  do {
    prefix = `a3sExt${context.nextExtensionPrefix}`;
    context.nextExtensionPrefix += 1;
  } while (xmlNamespaceUri(context.generatedRoot, prefix));
  context.generatedRoot.setAttributeNS(
    XMLNS_NAMESPACE,
    `xmlns:${prefix}`,
    namespace,
  );
  return prefix;
}

function isWordprocessingNamespace(namespace: string | null): boolean {
  return Boolean(namespace && WORDPROCESSING_NAMESPACES.has(namespace));
}

function isNamespacePrefix(value: string): boolean {
  return /^[A-Za-z_][A-Za-z0-9_.-]*$/.test(value) && value !== 'xml';
}
