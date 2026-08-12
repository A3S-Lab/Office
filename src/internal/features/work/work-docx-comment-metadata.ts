import JSZip from 'jszip';
import { normalizeDocumentParagraphId } from './work-document-paragraph-identity';
import {
  type AssignedDocxCommentThread,
  normalizeDocxCommentId,
} from './work-docx-note-comment-identity';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  attribute,
  descendants,
  directChildren,
  parseXml,
} from './work-ooxml-package';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
  xmlDeclaredPrefix,
  xmlNamespaceUri,
} from './work-docx-settings-xml';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const CONTENT_TYPES_PATH = '[Content_Types].xml';
const COMMENTS_PATH = 'word/comments.xml';
const COMMENTS_EXTENDED_PATH = 'word/commentsExtended.xml';
const DOCUMENT_RELATIONSHIPS_PATH = 'word/_rels/document.xml.rels';
const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const WORD_2012_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2012/wordml';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const CONTENT_TYPES_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const PACKAGE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const COMMENTS_EXTENDED_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
const COMMENTS_EXTENDED_RELATIONSHIP =
  'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
const MAX_COMMENT_METADATA_RECORDS = 65_536;

interface CommentParagraph {
  id: number;
  paragraphId: string;
}

export async function patchDocxCommentMetadata(
  buffer: ArrayBuffer,
  threads: readonly AssignedDocxCommentThread[],
): Promise<ArrayBuffer> {
  if (!threads.length) return buffer;
  const recordCount = threads.reduce(
    (count, thread) => count + 1 + thread.replies.length,
    0,
  );
  if (recordCount > MAX_COMMENT_METADATA_RECORDS) {
    throw new Error('Document exceeds the comment-metadata limit.');
  }
  const archive = await JSZip.loadAsync(buffer);
  const paragraphs = await commentParagraphs(archive);
  const byId = uniqueCommentParagraphs(paragraphs);
  const document = parseXml(
    `<w15:commentsEx xmlns:w15="${WORD_2012_NAMESPACE}" xmlns:mc="${MARKUP_COMPATIBILITY_NAMESPACE}" mc:Ignorable="w15"/>`,
    'generated commentsExtended',
  );
  for (const thread of threads) {
    const root = requiredParagraph(byId, thread.id);
    appendCommentExtended(document, root.paragraphId, {
      done: thread.comment.resolved,
    });
    for (const assignedReply of thread.replies) {
      const reply = requiredParagraph(byId, assignedReply.id);
      appendCommentExtended(document, reply.paragraphId, {
        parentParagraphId: root.paragraphId,
      });
    }
  }
  if (byId.size !== recordCount) {
    throw new Error('Generated DOCX contains unexpected comment records.');
  }
  archive.file(COMMENTS_EXTENDED_PATH, serializeUtf8Xml(document));
  await ensureCommentMetadataRegistration(archive);
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function commentParagraphs(archive: JSZip): Promise<CommentParagraph[]> {
  const entry = archive.file(COMMENTS_PATH);
  if (!entry) throw new Error('Generated DOCX comments part is missing.');
  const document = parseXml(
    decodeXmlBytes(
      await entry.async('uint8array'),
      `generated DOCX ${COMMENTS_PATH}`,
    ),
    `generated DOCX ${COMMENTS_PATH}`,
  );
  if (
    document.documentElement.localName !== 'comments' ||
    !isWordElement(document.documentElement)
  ) {
    throw new Error('Generated DOCX comments part has an invalid root.');
  }
  const comments = directChildren(document.documentElement, 'comment').filter(
    isWordElement,
  );
  const paragraphIdCounts = new Map<string, number>();
  for (const paragraph of descendants(document, 'p').filter(isWordElement)) {
    const id = normalizeDocumentParagraphId(
      namespacedAttribute(paragraph, 'paraId', WORD_2010_NAMESPACE),
    );
    if (id) paragraphIdCounts.set(id, (paragraphIdCounts.get(id) ?? 0) + 1);
  }
  const used = new Set(
    Array.from(paragraphIdCounts).flatMap(([id, count]) =>
      count === 1 ? [id] : [],
    ),
  );
  const records: CommentParagraph[] = [];
  let changed = false;
  for (const comment of comments) {
    const id = normalizeDocxCommentId(wordAttribute(comment, 'id'));
    const paragraph = descendants(comment, 'p').filter(isWordElement).at(-1);
    if (id === null || !paragraph) continue;
    const current = normalizeDocumentParagraphId(
      namespacedAttribute(paragraph, 'paraId', WORD_2010_NAMESPACE),
    );
    const paragraphId =
      current && paragraphIdCounts.get(current) === 1
        ? current
        : nextParagraphId(id, used);
    used.add(paragraphId);
    const textId = normalizeDocumentParagraphId(
      namespacedAttribute(paragraph, 'textId', WORD_2010_NAMESPACE),
    );
    if (current !== paragraphId || !textId) {
      setWord2010Identity(
        document,
        paragraph,
        paragraphId,
        textId ?? paragraphId,
      );
      changed = true;
    }
    records.push({ id, paragraphId });
  }
  if (changed) archive.file(COMMENTS_PATH, serializeUtf8Xml(document));
  return records;
}

function uniqueCommentParagraphs(
  paragraphs: readonly CommentParagraph[],
): Map<number, CommentParagraph> {
  const byId = new Map<number, CommentParagraph>();
  const paragraphIds = new Set<string>();
  for (const paragraph of paragraphs) {
    if (byId.has(paragraph.id) || paragraphIds.has(paragraph.paragraphId)) {
      throw new Error('Generated DOCX comment identities are ambiguous.');
    }
    byId.set(paragraph.id, paragraph);
    paragraphIds.add(paragraph.paragraphId);
  }
  return byId;
}

function requiredParagraph(
  paragraphs: ReadonlyMap<number, CommentParagraph>,
  id: number,
): CommentParagraph {
  const paragraph = paragraphs.get(id);
  if (!paragraph) {
    throw new Error(`Generated DOCX comment ${id} has no stable paragraph.`);
  }
  return paragraph;
}

function appendCommentExtended(
  document: Document,
  paragraphId: string,
  options: { done?: boolean; parentParagraphId?: string },
): void {
  const element = document.createElementNS(
    WORD_2012_NAMESPACE,
    'w15:commentEx',
  );
  element.setAttributeNS(WORD_2012_NAMESPACE, 'w15:paraId', paragraphId);
  if (options.parentParagraphId) {
    element.setAttributeNS(
      WORD_2012_NAMESPACE,
      'w15:paraIdParent',
      options.parentParagraphId,
    );
  } else {
    element.setAttributeNS(
      WORD_2012_NAMESPACE,
      'w15:done',
      options.done ? '1' : '0',
    );
  }
  document.documentElement.append(element);
}

async function ensureCommentMetadataRegistration(
  archive: JSZip,
): Promise<void> {
  const [typesEntry, relationshipsEntry] = [
    archive.file(CONTENT_TYPES_PATH),
    archive.file(DOCUMENT_RELATIONSHIPS_PATH),
  ];
  if (!typesEntry || !relationshipsEntry) {
    throw new Error('Generated DOCX package registration is incomplete.');
  }
  const types = parseXml(
    decodeXmlBytes(
      await typesEntry.async('uint8array'),
      `generated DOCX ${CONTENT_TYPES_PATH}`,
    ),
    `generated DOCX ${CONTENT_TYPES_PATH}`,
  );
  const partName = `/${COMMENTS_EXTENDED_PATH}`;
  if (
    !directChildren(types.documentElement, 'Override').some(
      (item) => attribute(item, 'PartName') === partName,
    )
  ) {
    const override = types.createElementNS(CONTENT_TYPES_NAMESPACE, 'Override');
    override.setAttribute('PartName', partName);
    override.setAttribute('ContentType', COMMENTS_EXTENDED_CONTENT_TYPE);
    types.documentElement.append(override);
  }
  archive.file(CONTENT_TYPES_PATH, serializeUtf8Xml(types));

  const relationships = parseXml(
    decodeXmlBytes(
      await relationshipsEntry.async('uint8array'),
      `generated DOCX ${DOCUMENT_RELATIONSHIPS_PATH}`,
    ),
    `generated DOCX ${DOCUMENT_RELATIONSHIPS_PATH}`,
  );
  const items = directChildren(relationships.documentElement, 'Relationship');
  if (
    !items.some(
      (item) =>
        attribute(item, 'Type') === COMMENTS_EXTENDED_RELATIONSHIP &&
        attribute(item, 'Target') === 'commentsExtended.xml',
    )
  ) {
    const usedIds = new Set(
      items.map((item) => attribute(item, 'Id') ?? '').filter(Boolean),
    );
    const relationship = relationships.createElementNS(
      PACKAGE_RELATIONSHIPS_NAMESPACE,
      'Relationship',
    );
    relationship.setAttribute('Id', nextRelationshipId(usedIds));
    relationship.setAttribute('Type', COMMENTS_EXTENDED_RELATIONSHIP);
    relationship.setAttribute('Target', 'commentsExtended.xml');
    relationships.documentElement.append(relationship);
  }
  archive.file(DOCUMENT_RELATIONSHIPS_PATH, serializeUtf8Xml(relationships));
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

function nextRelationshipId(used: ReadonlySet<string>): string {
  let index = 1;
  while (used.has(`rId${index}`)) index += 1;
  return `rId${index}`;
}

function nextParagraphId(commentId: number, used: ReadonlySet<string>): string {
  let candidate = commentId >= 0 && commentId < 0x7fff_fffe ? commentId + 1 : 1;
  while (candidate <= 0x7fff_ffff) {
    const id = candidate.toString(16).toUpperCase().padStart(8, '0');
    if (!used.has(id)) return id;
    candidate += 1;
  }
  candidate = 1;
  while (candidate <= commentId) {
    const id = candidate.toString(16).toUpperCase().padStart(8, '0');
    if (!used.has(id)) return id;
    candidate += 1;
  }
  throw new Error('Generated DOCX exhausted comment paragraph identities.');
}

function setWord2010Identity(
  document: Document,
  paragraph: Element,
  paragraphId: string,
  textId: string,
): void {
  const root = document.documentElement;
  const prefix = ensureNamespacePrefix(root, 'w14', WORD_2010_NAMESPACE);
  removeNamespacedAttribute(paragraph, 'paraId', WORD_2010_NAMESPACE);
  removeNamespacedAttribute(paragraph, 'textId', WORD_2010_NAMESPACE);
  paragraph.setAttributeNS(
    WORD_2010_NAMESPACE,
    `${prefix}:paraId`,
    paragraphId,
  );
  paragraph.setAttributeNS(WORD_2010_NAMESPACE, `${prefix}:textId`, textId);
  const compatibilityPrefix = ensureNamespacePrefix(
    root,
    'mc',
    MARKUP_COMPATIBILITY_NAMESPACE,
  );
  const tokens = (root.getAttribute(`${compatibilityPrefix}:Ignorable`) ?? '')
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  if (
    !tokens.some(
      (token) => xmlNamespaceUri(root, token) === WORD_2010_NAMESPACE,
    )
  ) {
    tokens.push(prefix);
  }
  removeNamespacedAttribute(root, 'Ignorable', MARKUP_COMPATIBILITY_NAMESPACE);
  root.setAttributeNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    `${compatibilityPrefix}:Ignorable`,
    tokens.join(' '),
  );
}

function removeNamespacedAttribute(
  element: Element,
  localName: string,
  namespace: string,
): void {
  for (const item of Array.from(element.attributes)) {
    if (
      xmlAttributeLocalName(item) === localName &&
      xmlAttributeNamespace(element, item) === namespace
    ) {
      element.removeAttributeNode(item);
    }
  }
}

function ensureNamespacePrefix(
  root: Element,
  preferred: string,
  namespace: string,
): string {
  const existing = xmlDeclaredPrefix(root, namespace);
  if (existing) return existing;
  if (!xmlNamespaceUri(root, preferred)) {
    root.setAttributeNS(XMLNS_NAMESPACE, `xmlns:${preferred}`, namespace);
    return preferred;
  }
  let suffix = 1;
  let prefix = '';
  do {
    prefix = `a3s${suffix}`;
    suffix += 1;
  } while (xmlNamespaceUri(root, prefix));
  root.setAttributeNS(XMLNS_NAMESPACE, `xmlns:${prefix}`, namespace);
  return prefix;
}

function isWordElement(element: Element): boolean {
  return DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '');
}
