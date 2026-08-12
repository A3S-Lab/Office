import type JSZip from 'jszip';
import { normalizeDocumentParagraphId } from './work-document-paragraph-identity';
import { normalizeDocxCommentId } from './work-docx-note-comment-identity';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  attribute,
  descendants,
  directChildren,
  parseXml,
  resolvePartTarget,
} from './work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const COMMENTS_IDS_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2016/wordml/cid';
const COMMENTS_PATH = 'word/comments.xml';
const DOCUMENT_RELATIONSHIPS_PATH = 'word/_rels/document.xml.rels';
const COMMENTS_IDS_RELATIONSHIP =
  'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds';
const MAX_COMMENT_RECORDS = 65_536;

interface CommentRecord {
  id: string;
  paragraphId: string;
}

interface CommentIndex {
  byId: Map<string, CommentRecord[]>;
  byParagraphId: Map<string, CommentRecord[]>;
}

interface DurableRecord {
  comment: CommentRecord;
  durableId: string;
}

export async function preserveDocxCommentDurableIds(
  generated: JSZip,
  source: JSZip,
  sourcePath: string,
): Promise<boolean> {
  if (!(await hasCommentsIdsRelationship(source, sourcePath))) return false;
  const [generatedIndex, sourceIndex, sourceIds] = await Promise.all([
    loadCommentIndex(generated, COMMENTS_PATH, 'generated'),
    loadCommentIndex(source, COMMENTS_PATH, 'source'),
    loadCommentsIds(source, sourcePath),
  ]);
  if (!generatedIndex || !sourceIndex || !sourceIds) return false;
  const entries = directChildren(sourceIds.documentElement, 'commentId').filter(
    (item) => item.namespaceURI === COMMENTS_IDS_NAMESPACE,
  );
  if (entries.length > MAX_COMMENT_RECORDS) {
    throw new Error(
      'Registered source DOCX exceeds the durable-comment limit.',
    );
  }
  const candidates = entries.flatMap((entry) => {
    const paragraphId = normalizeDocumentParagraphId(
      namespacedAttribute(entry, 'paraId', COMMENTS_IDS_NAMESPACE),
    );
    const durableId = normalizeDurableId(
      namespacedAttribute(entry, 'durableId', COMMENTS_IDS_NAMESPACE),
    );
    const comments = paragraphId
      ? (sourceIndex.byParagraphId.get(paragraphId) ?? [])
      : [];
    return paragraphId && durableId && comments.length === 1
      ? [{ comment: comments[0], durableId }]
      : [];
  });
  const durableCounts = counts(candidates.map(({ durableId }) => durableId));
  const commentCounts = counts(candidates.map(({ comment }) => comment.id));
  const retained: Array<DurableRecord & { generatedParagraphId: string }> = [];
  for (const candidate of candidates) {
    const sourceComments = sourceIndex.byId.get(candidate.comment.id) ?? [];
    const generatedComments =
      generatedIndex.byId.get(candidate.comment.id) ?? [];
    if (
      durableCounts.get(candidate.durableId) !== 1 ||
      commentCounts.get(candidate.comment.id) !== 1 ||
      sourceComments.length !== 1 ||
      generatedComments.length !== 1 ||
      generatedIndex.byParagraphId.get(generatedComments[0].paragraphId)
        ?.length !== 1
    ) {
      continue;
    }
    retained.push({
      ...candidate,
      generatedParagraphId: generatedComments[0].paragraphId,
    });
  }
  if (!retained.length) return false;
  const output = parseXml(
    `<w16cid:commentsIds xmlns:w16cid="${COMMENTS_IDS_NAMESPACE}"/>`,
    'generated commentsIds',
  );
  for (const record of retained) {
    const element = output.createElementNS(
      COMMENTS_IDS_NAMESPACE,
      'w16cid:commentId',
    );
    element.setAttributeNS(
      COMMENTS_IDS_NAMESPACE,
      'w16cid:paraId',
      record.generatedParagraphId,
    );
    element.setAttributeNS(
      COMMENTS_IDS_NAMESPACE,
      'w16cid:durableId',
      record.durableId,
    );
    output.documentElement.append(element);
  }
  generated.file(sourcePath, serializeUtf8Xml(output));
  return true;
}

async function hasCommentsIdsRelationship(
  archive: JSZip,
  sourcePath: string,
): Promise<boolean> {
  const entry = archive.file(DOCUMENT_RELATIONSHIPS_PATH);
  if (!entry) return false;
  try {
    const document = parseXml(
      decodeXmlBytes(
        await entry.async('uint8array'),
        `source DOCX ${DOCUMENT_RELATIONSHIPS_PATH}`,
      ),
      `source DOCX ${DOCUMENT_RELATIONSHIPS_PATH}`,
    );
    const matches = directChildren(
      document.documentElement,
      'Relationship',
    ).filter((item) => {
      const mode = (attribute(item, 'TargetMode') ?? '').trim().toLowerCase();
      return (
        attribute(item, 'Type') === COMMENTS_IDS_RELATIONSHIP &&
        (!mode || mode === 'internal') &&
        resolvePartTarget(
          'word/document.xml',
          attribute(item, 'Target')?.trim() ?? '',
        ).toLowerCase() === sourcePath.toLowerCase()
      );
    });
    return matches.length === 1;
  } catch {
    return false;
  }
}

async function loadCommentIndex(
  archive: JSZip,
  path: string,
  role: string,
): Promise<CommentIndex | null> {
  const entry = archive.file(path);
  if (!entry) return null;
  try {
    const document = parseXml(
      decodeXmlBytes(await entry.async('uint8array'), `${role} DOCX ${path}`),
      `${role} DOCX ${path}`,
    );
    if (
      document.documentElement.localName !== 'comments' ||
      !isWordElement(document.documentElement)
    ) {
      return null;
    }
    const comments = directChildren(document.documentElement, 'comment').filter(
      isWordElement,
    );
    if (comments.length > MAX_COMMENT_RECORDS) {
      throw new Error(`${role} DOCX exceeds the durable-comment limit.`);
    }
    const byId = new Map<string, CommentRecord[]>();
    const byParagraphId = new Map<string, CommentRecord[]>();
    for (const comment of comments) {
      const nativeId = normalizeDocxCommentId(wordAttribute(comment, 'id'));
      const id = nativeId === null ? null : String(nativeId);
      const paragraphs = descendants(comment, 'p').filter(isWordElement);
      const paragraphId = normalizeDocumentParagraphId(
        namespacedAttribute(
          paragraphs.at(-1) ?? comment,
          'paraId',
          WORD_2010_NAMESPACE,
        ),
      );
      if (!id || !paragraphId) continue;
      const record = { id, paragraphId };
      addIndexValue(byId, id, record);
      addIndexValue(byParagraphId, paragraphId, record);
    }
    return { byId, byParagraphId };
  } catch (error) {
    if (error instanceof Error && error.message.includes('limit')) throw error;
    return null;
  }
}

async function loadCommentsIds(
  archive: JSZip,
  path: string,
): Promise<Document | null> {
  const entry = archive.file(path);
  if (!entry) return null;
  try {
    const document = parseXml(
      decodeXmlBytes(await entry.async('uint8array'), `source DOCX ${path}`),
      `source DOCX ${path}`,
    );
    return document.documentElement.localName === 'commentsIds' &&
      document.documentElement.namespaceURI === COMMENTS_IDS_NAMESPACE
      ? document
      : null;
  } catch {
    return null;
  }
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

function normalizeDurableId(value: string | null): string | null {
  const normalized = normalizeDocumentParagraphId(value);
  return normalized && Number.parseInt(normalized, 16) < 0x7fff_ffff
    ? normalized
    : null;
}

function counts(values: readonly string[]): Map<string, number> {
  const result = new Map<string, number>();
  for (const value of values) result.set(value, (result.get(value) ?? 0) + 1);
  return result;
}

function addIndexValue<T>(map: Map<string, T[]>, key: string, value: T): void {
  const values = map.get(key) ?? [];
  values.push(value);
  map.set(key, values);
}

function isWordElement(element: Element): boolean {
  return DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '');
}
