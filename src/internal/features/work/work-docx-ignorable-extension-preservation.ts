import { attribute, directChildren } from './work-ooxml-package';
import {
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

export const DOCX_WORDPROCESSING_NAMESPACES = new Set([
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

export type DocxExtensionDocumentRole = 'generated' | 'source';

export interface DocxIgnorableExtensionOptions {
  semanticKey?: (element: Element, role: DocxExtensionDocumentRole) => string;
  isAdditionalSemanticNamespace?: (namespace: string) => boolean;
  allowExtensionNamespace?: (namespace: string) => boolean;
  allowSemanticElement?: (
    element: Element,
    generatedSemanticNames: ReadonlySet<string>,
  ) => boolean;
}

export interface DocxIgnorableExtensionPair {
  generated: Element;
  source: Element;
}

interface IgnorableNamespace {
  prefix: string;
  namespace: string;
}

interface ExtensionMergeContext {
  generatedDocument: Document;
  generatedRoot: Element;
  sourceRoot: Element;
  sourceIgnorable: readonly IgnorableNamespace[];
  sourceIgnorableNamespaces: ReadonlySet<string>;
  retainedNamespaces: Set<string>;
  nextExtensionPrefix: number;
  options: DocxIgnorableExtensionOptions;
}

export function mergeDocxIgnorableExtensions(
  generatedDocument: Document,
  sourceDocument: Document,
  options: DocxIgnorableExtensionOptions = {},
): void {
  mergeDocxIgnorableExtensionsAtPairs(
    generatedDocument,
    sourceDocument,
    [
      {
        generated: generatedDocument.documentElement,
        source: sourceDocument.documentElement,
      },
    ],
    options,
  );
}

export function mergeDocxIgnorableExtensionsAtPairs(
  generatedDocument: Document,
  sourceDocument: Document,
  pairs: readonly DocxIgnorableExtensionPair[],
  options: DocxIgnorableExtensionOptions = {},
): void {
  if (!pairs.length) return;
  const sourceIgnorable = ignorableNamespaces(sourceDocument.documentElement);
  const context: ExtensionMergeContext = {
    generatedDocument,
    generatedRoot: generatedDocument.documentElement,
    sourceRoot: sourceDocument.documentElement,
    sourceIgnorable,
    sourceIgnorableNamespaces: new Set(
      sourceIgnorable.map((item) => item.namespace),
    ),
    retainedNamespaces: new Set(),
    nextExtensionPrefix: 1,
    options,
  };
  for (const pair of pairs) {
    mergeKnownElement(pair.generated, pair.source, context, 0);
  }
  mergeIgnorableNamespaces(context);
  mergeCompatibilityQNameRules(context);
}

function mergeKnownElement(
  generated: Element,
  source: Element,
  context: ExtensionMergeContext,
  depth: number,
): void {
  mergePassiveAttributes(generated, source, context);
  if (depth >= MAX_EXTENSION_DEPTH) return;

  const generatedChildren = directChildren(generated);
  const sourceChildren = directChildren(source);
  const generatedGroups = groupChildren(
    generatedChildren,
    'generated',
    context,
  );
  const sourceGroups = groupChildren(sourceChildren, 'source', context);

  for (const sourceChild of sourceChildren) {
    const key = semanticElementKey(sourceChild, 'source', context);
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
    const key = semanticElementKey(sourceChild, 'source', context);
    if ((generatedGroups.get(key)?.length ?? 0) > 0) continue;
    if (!isExtensionCandidate(sourceChild, context)) continue;
    if (!isSafeExtensionSubtree(sourceChild, generatedSemanticNames, context)) {
      continue;
    }

    const imported = cloneXmlElement(context.generatedDocument, sourceChild);
    declareInheritedNamespaces(imported, sourceChild);
    const anchor = followingGeneratedAnchor(
      sourceChildren,
      index,
      generatedGroups,
      sourceGroups,
      context,
    );
    generated.insertBefore(imported, anchor);
    collectRetainedNamespaces(sourceChild, context.retainedNamespaces);
  }
}

function mergePassiveAttributes(
  generated: Element,
  source: Element,
  context: ExtensionMergeContext,
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
      !isAllowedExtensionNamespace(namespace, context) ||
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
  context: ExtensionMergeContext,
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
    context.sourceIgnorableNamespaces.has(namespace) &&
    isAllowedExtensionNamespace(namespace, context)
  );
}

function isSafeExtensionSubtree(
  root: Element,
  generatedSemanticNames: ReadonlySet<string>,
  context: ExtensionMergeContext,
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
    const namespace = current.element.namespaceURI;
    if (
      elementCount > MAX_EXTENSION_ELEMENTS ||
      current.depth > MAX_EXTENSION_DEPTH ||
      RELATIONSHIP_NAMESPACES.has(namespace ?? '') ||
      hasUnsafeExtensionAttribute(
        current.element,
        generatedSemanticNames,
        context,
      )
    ) {
      return false;
    }
    if (
      isSemanticNamespace(namespace, context) &&
      !context.options.allowSemanticElement?.(
        current.element,
        generatedSemanticNames,
      )
    ) {
      return false;
    }
    if (namespace === MARKUP_COMPATIBILITY_NAMESPACE) {
      if (current.element.localName === 'AlternateContent') {
        alternateContentCount += 1;
        if (
          alternateContentCount > 1 ||
          !isStructurallyValidAlternateContent(current.element) ||
          hasUnsafeAlternateContentRequirements(current.element, context)
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

function hasUnsafeExtensionAttribute(
  element: Element,
  generatedSemanticNames: ReadonlySet<string>,
  context: ExtensionMergeContext,
): boolean {
  return Array.from(element.attributes).some((item) => {
    const namespace = xmlAttributeNamespace(element, item) ?? '';
    if (
      !namespace ||
      namespace === XMLNS_NAMESPACE ||
      namespace === XML_NAMESPACE
    ) {
      return false;
    }
    if (
      RELATIONSHIP_NAMESPACES.has(namespace) ||
      namespace === MARKUP_COMPATIBILITY_NAMESPACE
    ) {
      return true;
    }
    if (isSemanticNamespace(namespace, context)) {
      return !(
        isSemanticNamespace(element.namespaceURI, context) &&
        context.options.allowSemanticElement?.(element, generatedSemanticNames)
      );
    }
    return !isAllowedExtensionNamespace(namespace, context);
  });
}

function hasUnsafeAlternateContentRequirements(
  alternateContent: Element,
  context: ExtensionMergeContext,
): boolean {
  for (const choice of directChildren(alternateContent, 'Choice')) {
    for (const prefix of (attribute(choice, 'Requires') ?? '')
      .trim()
      .split(/\s+/)
      .filter(Boolean)) {
      const namespace = xmlNamespaceUri(choice, prefix);
      if (
        namespace &&
        (isSemanticNamespace(namespace, context) ||
          !isAllowedExtensionNamespace(namespace, context))
      ) {
        return true;
      }
    }
  }
  return false;
}

function followingGeneratedAnchor(
  sourceChildren: readonly Element[],
  sourceIndex: number,
  generatedGroups: ReadonlyMap<string, readonly Element[]>,
  sourceGroups: ReadonlyMap<string, readonly Element[]>,
  context: ExtensionMergeContext,
): Element | null {
  for (let index = sourceIndex + 1; index < sourceChildren.length; index += 1) {
    const key = semanticElementKey(sourceChildren[index], 'source', context);
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

function groupChildren(
  children: readonly Element[],
  role: DocxExtensionDocumentRole,
  context: ExtensionMergeContext,
): Map<string, Element[]> {
  const groups = new Map<string, Element[]>();
  for (const child of children) {
    const key = semanticElementKey(child, role, context);
    const existing = groups.get(key);
    if (existing) existing.push(child);
    else groups.set(key, [child]);
  }
  return groups;
}

function semanticElementKey(
  element: Element,
  role: DocxExtensionDocumentRole,
  context: ExtensionMergeContext,
): string {
  return (
    context.options.semanticKey?.(element, role) ??
    (isWordprocessingNamespace(element.namespaceURI)
      ? `word:${element.localName}`
      : `{${element.namespaceURI ?? ''}}${element.localName}`)
  );
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

function mergeIgnorableNamespaces(context: ExtensionMergeContext): void {
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
  let changed = false;
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
      changed = true;
    }
    coveredNamespaces.add(item.namespace);
  }
  if (changed) setCompatibilityAttribute(context, 'Ignorable', tokens);
}

function mergeCompatibilityQNameRules(context: ExtensionMergeContext): void {
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
    let changed = false;
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
      if (!used.has(token)) {
        tokens.push(token);
        changed = true;
      }
      used.add(token);
    }
    if (changed) setCompatibilityAttribute(context, name, tokens);
  }
}

function setCompatibilityAttribute(
  context: ExtensionMergeContext,
  localName: string,
  tokens: readonly string[],
): void {
  for (const item of Array.from(context.generatedRoot.attributes)) {
    if (
      xmlAttributeLocalName(item) === localName &&
      (xmlAttributeNamespace(context.generatedRoot, item) ===
        MARKUP_COMPATIBILITY_NAMESPACE ||
        item.name === `mc:${localName}`)
    ) {
      context.generatedRoot.removeAttributeNode(item);
    }
  }
  const prefix = ensureRootNamespace(
    context,
    'mc',
    MARKUP_COMPATIBILITY_NAMESPACE,
  );
  context.generatedRoot.setAttributeNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    `${prefix}:${localName}`,
    tokens.join(' '),
  );
}

function ensureRootNamespace(
  context: ExtensionMergeContext,
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
  return Boolean(namespace && DOCX_WORDPROCESSING_NAMESPACES.has(namespace));
}

function isSemanticNamespace(
  namespace: string | null,
  context: ExtensionMergeContext,
): boolean {
  return Boolean(
    isWordprocessingNamespace(namespace) ||
      (namespace && context.options.isAdditionalSemanticNamespace?.(namespace)),
  );
}

function isAllowedExtensionNamespace(
  namespace: string,
  context: ExtensionMergeContext,
): boolean {
  return context.options.allowExtensionNamespace?.(namespace) ?? true;
}

function isNamespacePrefix(value: string): boolean {
  return /^[A-Za-z_][A-Za-z0-9_.-]*$/.test(value) && value !== 'xml';
}
