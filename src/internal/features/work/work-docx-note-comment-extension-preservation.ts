import type JSZip from 'jszip';
import { normalizeDocxCommentId } from './work-docx-note-comment-identity';
import { preserveDocxNoteCommentHyperlinks } from './work-docx-note-comment-hyperlink-preservation';
import { preserveDocxNoteCommentRunContent } from './work-docx-note-comment-run-preservation';
import {
  DOCX_WORDPROCESSING_NAMESPACES,
  mergeDocxIgnorableExtensionsAtPairs,
  type DocxExtensionDocumentRole,
  type DocxIgnorableExtensionPair,
} from './work-docx-ignorable-extension-preservation';
import { descendants, directChildren, parseXml } from './work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const WORD_2012_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2012/wordml';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MAX_NOTE_COMMENT_SCOPES = 65_536;

interface LoadedPart {
  document: Document;
  path: string;
}

interface CommentIndex {
  byId: Map<string, Element[]>;
  idByParagraphId: Map<string, string[]>;
}

export async function preserveDocxNoteCommentExtensions(
  generated: JSZip,
  source: JSZip,
): Promise<void> {
  for (const config of [
    {
      path: 'word/footnotes.xml',
      root: 'footnotes',
      item: 'footnote',
      note: true,
    },
    {
      path: 'word/endnotes.xml',
      root: 'endnotes',
      item: 'endnote',
      note: true,
    },
    {
      path: 'word/comments.xml',
      root: 'comments',
      item: 'comment',
      note: false,
    },
  ] as const) {
    await preserveWordItemExtensions(generated, source, config);
  }
  await preserveCommentExtendedExtensions(generated, source);
}

async function preserveWordItemExtensions(
  generatedArchive: JSZip,
  sourceArchive: JSZip,
  config: {
    path: string;
    root: string;
    item: string;
    note: boolean;
  },
): Promise<void> {
  const [generated, source] = await Promise.all([
    loadPart(generatedArchive, config.path, config.root, 'generated'),
    loadPart(sourceArchive, config.path, config.root, 'source'),
  ]);
  if (!generated || !source) return;
  const generatedItems = wordItems(generated.document, config.item).filter(
    (item) => !config.note || isNormalNote(item),
  );
  const sourceItems = wordItems(source.document, config.item).filter(
    (item) => !config.note || isNormalNote(item),
  );
  assertScopeLimit(generatedItems.length, 'Generated');
  assertScopeLimit(sourceItems.length, 'Registered source');
  const pairs = uniqueIdentityPairs(generatedItems, sourceItems, (item) =>
    itemIdentity(item, config.note),
  );
  if (!pairs.length) return;
  await preserveDocxNoteCommentHyperlinks(
    generatedArchive,
    sourceArchive,
    generated.document,
    source.document,
    pairs,
    config.note ? 'note' : 'comment',
    config.path,
  );
  preserveDocxNoteCommentRunContent(
    generated.document,
    source.document,
    pairs,
    config.note ? 'note' : 'comment',
  );
  mergePairs(generated.document, source.document, pairs);
  generatedArchive.file(config.path, serializeUtf8Xml(generated.document));
}

async function preserveCommentExtendedExtensions(
  generatedArchive: JSZip,
  sourceArchive: JSZip,
): Promise<void> {
  const [generatedComments, sourceComments, generatedExtended, sourceExtended] =
    await Promise.all([
      loadPart(generatedArchive, 'word/comments.xml', 'comments', 'generated'),
      loadPart(sourceArchive, 'word/comments.xml', 'comments', 'source'),
      loadPart(
        generatedArchive,
        'word/commentsExtended.xml',
        'commentsEx',
        'generated',
        WORD_2012_NAMESPACE,
      ),
      loadPart(
        sourceArchive,
        'word/commentsExtended.xml',
        'commentsEx',
        'source',
        WORD_2012_NAMESPACE,
      ),
    ]);
  if (
    !generatedComments ||
    !sourceComments ||
    !generatedExtended ||
    !sourceExtended
  ) {
    return;
  }
  const generatedIndex = indexComments(generatedComments.document);
  const sourceIndex = indexComments(sourceComments.document);
  const generatedEntries = extendedEntries(generatedExtended.document);
  const sourceEntries = extendedEntries(sourceExtended.document);
  assertScopeLimit(generatedEntries.length, 'Generated');
  assertScopeLimit(sourceEntries.length, 'Registered source');
  const pairs = uniqueIdentityPairs(
    generatedEntries,
    sourceEntries,
    (item, role) => {
      const paragraphId = namespacedAttribute(
        item,
        'paraId',
        WORD_2012_NAMESPACE,
      );
      const index = role === 'generated' ? generatedIndex : sourceIndex;
      const ids = paragraphId ? index.idByParagraphId.get(paragraphId) : null;
      return ids?.length === 1 ? ids[0] : null;
    },
  ).filter(({ generated, source }) => {
    const id = commentIdForExtended(generated, generatedIndex);
    if (!id) return false;
    return (
      generatedIndex.byId.get(id)?.length === 1 &&
      sourceIndex.byId.get(id)?.length === 1 &&
      commentIdForExtended(source, sourceIndex) === id
    );
  });
  if (!pairs.length) return;
  mergePairs(generatedExtended.document, sourceExtended.document, pairs);
  generatedArchive.file(
    generatedExtended.path,
    serializeUtf8Xml(generatedExtended.document),
  );
}

async function loadPart(
  archive: JSZip,
  path: string,
  rootName: string,
  role: string,
  rootNamespace?: string,
): Promise<LoadedPart | null> {
  const entry = archive.file(path);
  if (!entry) return null;
  try {
    const document = parseXml(
      decodeXmlBytes(await entry.async('uint8array'), `${role} DOCX ${path}`),
      `${role} DOCX ${path}`,
    );
    const root = document.documentElement;
    const validNamespace = rootNamespace
      ? root.namespaceURI === rootNamespace
      : DOCX_WORDPROCESSING_NAMESPACES.has(root.namespaceURI ?? '');
    return root.localName === rootName && validNamespace
      ? { document, path }
      : null;
  } catch {
    return null;
  }
}

function indexComments(document: Document): CommentIndex {
  const byId = new Map<string, Element[]>();
  const idByParagraphId = new Map<string, string[]>();
  for (const comment of wordItems(document, 'comment')) {
    const nativeId = normalizeDocxCommentId(wordAttribute(comment, 'id'));
    const id = nativeId === null ? null : String(nativeId);
    if (!id) continue;
    addIndexValue(byId, id, comment);
    const paragraphs = descendants(comment, 'p').filter(isWordElement);
    const paragraphId = namespacedAttribute(
      paragraphs.at(-1) ?? comment,
      'paraId',
      WORD_2010_NAMESPACE,
    );
    if (paragraphId) addIndexValue(idByParagraphId, paragraphId, id);
  }
  return { byId, idByParagraphId };
}

function wordItems(document: Document, localName: string): Element[] {
  return directChildren(document.documentElement, localName).filter(
    isWordElement,
  );
}

function extendedEntries(document: Document): Element[] {
  return directChildren(document.documentElement, 'commentEx').filter(
    (item) => item.namespaceURI === WORD_2012_NAMESPACE,
  );
}

function uniqueIdentityPairs(
  generated: readonly Element[],
  source: readonly Element[],
  identity: (
    element: Element,
    role: DocxExtensionDocumentRole,
  ) => string | null,
): DocxIgnorableExtensionPair[] {
  const generatedById = groupByIdentity(generated, (item) =>
    identity(item, 'generated'),
  );
  const sourceById = groupByIdentity(source, (item) =>
    identity(item, 'source'),
  );
  const pairs: DocxIgnorableExtensionPair[] = [];
  for (const [id, sourceItems] of sourceById) {
    const generatedItems = generatedById.get(id) ?? [];
    if (sourceItems.length === 1 && generatedItems.length === 1) {
      pairs.push({ generated: generatedItems[0], source: sourceItems[0] });
    }
  }
  return pairs;
}

function groupByIdentity(
  elements: readonly Element[],
  identity: (element: Element) => string | null,
): Map<string, Element[]> {
  const result = new Map<string, Element[]>();
  for (const element of elements) {
    const id = identity(element);
    if (!id) continue;
    addIndexValue(result, id, element);
  }
  return result;
}

function mergePairs(
  generated: Document,
  source: Document,
  pairs: readonly DocxIgnorableExtensionPair[],
): void {
  mergeDocxIgnorableExtensionsAtPairs(generated, source, pairs, {
    semanticKey: extensionSemanticKey,
    isAdditionalSemanticNamespace: isKnownOoxmlNamespace,
    allowExtensionNamespace: (namespace) => !isKnownOoxmlNamespace(namespace),
    allowMatchedElementMerge: (_generated, _source, depth) => depth === 0,
  });
}

function commentIdForExtended(
  item: Element,
  index: CommentIndex,
): string | null {
  const paragraphId = namespacedAttribute(item, 'paraId', WORD_2012_NAMESPACE);
  const ids = paragraphId ? index.idByParagraphId.get(paragraphId) : null;
  return ids?.length === 1 ? ids[0] : null;
}

function wordAttribute(element: Element, name: string): string | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === name &&
        DOCX_WORDPROCESSING_NAMESPACES.has(
          xmlAttributeNamespace(element, item) ?? '',
        ),
    )?.value ?? null
  );
}

function namespacedAttribute(
  element: Element,
  name: string,
  namespace: string,
): string | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === name &&
        xmlAttributeNamespace(element, item) === namespace,
    )?.value ?? null
  );
}

function positiveDecimalIdentity(value: string | null): string | null {
  if (!value || !/^[1-9][0-9]{0,9}$/.test(value)) return null;
  const number = Number(value);
  return Number.isSafeInteger(number) && number <= 2_147_483_647
    ? String(number)
    : null;
}

function itemIdentity(element: Element, note: boolean): string | null {
  const value = wordAttribute(element, 'id');
  if (note) return positiveDecimalIdentity(value);
  const id = normalizeDocxCommentId(value);
  return id === null ? null : String(id);
}

function isNormalNote(element: Element): boolean {
  const type = wordAttribute(element, 'type')?.trim().toLowerCase();
  return !type || type === 'normal';
}

function extensionSemanticKey(
  element: Element,
  _role: DocxExtensionDocumentRole,
): string {
  return isKnownOoxmlNamespace(element.namespaceURI ?? '')
    ? `{semantic}${element.localName}`
    : `{${element.namespaceURI ?? ''}}${element.localName}`;
}

function isKnownOoxmlNamespace(namespace: string): boolean {
  if (namespace === MARKUP_COMPATIBILITY_NAMESPACE) return false;
  return (
    DOCX_WORDPROCESSING_NAMESPACES.has(namespace) ||
    namespace.startsWith('http://schemas.microsoft.com/office/') ||
    namespace.startsWith('http://schemas.openxmlformats.org/') ||
    namespace.startsWith('http://purl.oclc.org/ooxml/') ||
    namespace.startsWith('urn:schemas-microsoft-com:') ||
    namespace.startsWith('urn:microsoft-com:office:')
  );
}

function isWordElement(element: Element): boolean {
  return DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '');
}

function addIndexValue<T>(map: Map<string, T[]>, key: string, value: T): void {
  const matches = map.get(key) ?? [];
  matches.push(value);
  map.set(key, matches);
}

function assertScopeLimit(count: number, role: string): void {
  if (count > MAX_NOTE_COMMENT_SCOPES) {
    throw new Error(`${role} DOCX exceeds the note/comment-scope limit.`);
  }
}
