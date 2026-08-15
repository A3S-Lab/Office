import { MarkdownManager } from '@tiptap/markdown';
import { getSchema, type JSONContent } from '@tiptap/core';
import { createWorkDocumentExtensions } from './work-document-extensions';
import {
  createWorkDocumentModel,
  documentModelForContent,
} from './work-document-model';
import { materializeWorkDocumentContent } from './work-document-model-codec';
import type {
  WorkDocumentContent,
  WorkDocumentMark,
  WorkDocumentNode,
} from './work-types';

export const WORK_DOCUMENT_SOURCE_SCHEMA =
  'a3s.office.document.source' as const;
export const WORK_DOCUMENT_SOURCE_VERSION = 1 as const;
export const WORK_DOCUMENT_SOURCE_MEDIA_TYPE =
  'text/markdown; charset=utf-8' as const;

export interface WorkDocumentSource {
  schema: typeof WORK_DOCUMENT_SOURCE_SCHEMA;
  version: typeof WORK_DOCUMENT_SOURCE_VERSION;
  mediaType: typeof WORK_DOCUMENT_SOURCE_MEDIA_TYPE;
  content: string;
}

let markdownManager: MarkdownManager | null = null;
let documentSchema: ReturnType<typeof getSchema> | null = null;

/**
 * Project one structured Office document into an agent-readable source value.
 *
 * The source is deliberately a subordinate projection, not a replacement for
 * the lossless document snapshot. Version 1 supports one structured section so
 * it cannot silently flatten distinct section layouts.
 */
export function projectWorkDocumentSource(
  content: WorkDocumentContent,
): WorkDocumentSource {
  const section = singleDocumentSection(content);
  let source: string;
  try {
    source = documentMarkdownManager().serialize({
      type: 'doc',
      content: (section.content ?? []) as JSONContent[],
    });
  } catch (error) {
    throw new Error(
      'The structured document cannot be projected into the version 1 Markdown source contract.',
      { cause: error },
    );
  }
  return {
    schema: WORK_DOCUMENT_SOURCE_SCHEMA,
    version: WORK_DOCUMENT_SOURCE_VERSION,
    mediaType: WORK_DOCUMENT_SOURCE_MEDIA_TYPE,
    content: source,
  };
}

/**
 * Apply an agent-readable source revision to the document body while retaining
 * every Office-owned field and the exact section identity/layout attributes.
 */
export function applyWorkDocumentSource(
  content: WorkDocumentContent,
  source: WorkDocumentSource,
): WorkDocumentContent {
  assertDocumentSource(source);
  const model = documentModelForContent(content);
  if (!model) {
    throw new Error(
      'Applying a document source requires a synchronized structured snapshot.',
    );
  }
  const section = singleDocumentSection(content);
  let parsed: JSONContent;
  try {
    parsed = documentMarkdownManager().parse(source.content);
  } catch (error) {
    throw new Error(
      'The version 1 document source is not valid supported Markdown.',
      { cause: error },
    );
  }
  const parsedBody = parsed.content?.length
    ? (parsed.content as unknown as WorkDocumentNode[])
    : ([{ type: 'paragraph' }] satisfies WorkDocumentNode[]);
  const body = retainSurvivingDocumentCommentAnchors(
    section.content ?? [],
    parsedBody,
  );
  const root = normalizeDocumentRoot({
    ...model.root,
    content: [
      {
        ...section,
        ...(section.attrs ? { attrs: structuredClone(section.attrs) } : {}),
        content: body,
      },
    ],
  });
  const staged: WorkDocumentContent = {
    ...content,
    model: createWorkDocumentModel(content.html, root, model),
  };
  return materializeWorkDocumentContent(staged);
}

function normalizeDocumentRoot(root: WorkDocumentNode): WorkDocumentNode {
  try {
    documentSchema ??= getSchema(createWorkDocumentExtensions());
    return documentSchema.nodeFromJSON(root).toJSON() as WorkDocumentNode;
  } catch (error) {
    throw new Error(
      'The version 1 document source cannot be represented by the Office document schema.',
      { cause: error },
    );
  }
}

function singleDocumentSection(content: WorkDocumentContent): WorkDocumentNode {
  const model = documentModelForContent(content);
  if (!model) {
    throw new Error(
      'The document source contract requires a synchronized structured snapshot.',
    );
  }
  const sections = (model.root.content ?? []).filter(
    (node) => node.type === 'documentSection',
  );
  if (sections.length !== 1 || model.root.content?.length !== 1) {
    throw new Error(
      'The version 1 document source contract requires exactly one structured section.',
    );
  }
  return sections[0] as WorkDocumentNode;
}

function assertDocumentSource(
  source: WorkDocumentSource,
): asserts source is WorkDocumentSource {
  if (
    !source ||
    source.schema !== WORK_DOCUMENT_SOURCE_SCHEMA ||
    source.version !== WORK_DOCUMENT_SOURCE_VERSION ||
    source.mediaType !== WORK_DOCUMENT_SOURCE_MEDIA_TYPE ||
    typeof source.content !== 'string'
  ) {
    throw new Error('The A3S Office document source envelope is invalid.');
  }
}

function documentMarkdownManager(): MarkdownManager {
  markdownManager ??= new MarkdownManager({
    extensions: createWorkDocumentExtensions(),
    indentation: { style: 'space', size: 2 },
    markedOptions: { gfm: true, breaks: false, pedantic: false },
  });
  return markdownManager;
}

interface DocumentTextSpan {
  from: number;
  to: number;
  node: WorkDocumentNode;
}

interface DocumentTextIndex {
  content: string;
  spans: DocumentTextSpan[];
}

interface DocumentCommentAnchor {
  id: string;
  from: number;
  to: number;
  prefix: string;
  suffix: string;
  valid: boolean;
}

interface DocumentCommentRange {
  id: string;
  from: number;
  to: number;
}

const DOCUMENT_COMMENT_MARK = 'documentComment';
const DOCUMENT_COMMENT_CONTEXT_LENGTH = 64;

/**
 * Reapply an Office-owned comment mark only when its original anchor text has
 * one deterministic destination in the Agent-authored body. The Markdown
 * source cannot create comment identities, and ambiguous matches fail closed
 * instead of silently attaching a review thread to the wrong text.
 */
function retainSurvivingDocumentCommentAnchors(
  previousBody: readonly WorkDocumentNode[],
  nextBody: readonly WorkDocumentNode[],
): WorkDocumentNode[] {
  const previous = indexDocumentText(previousBody);
  const next = indexDocumentText(nextBody);
  const ranges = collectDocumentCommentAnchors(previous).flatMap((anchor) => {
    const anchorText = previous.content.slice(anchor.from, anchor.to);
    if (!anchor.valid || !anchorText) return [];
    const candidates = exactTextOccurrences(next.content, anchorText);
    const destination = uniqueCommentDestination(
      next.content,
      anchor,
      candidates,
    );
    return destination === null
      ? []
      : [
          {
            id: anchor.id,
            from: destination,
            to: destination + anchorText.length,
          },
        ];
  });
  return applyDocumentCommentRanges(nextBody, ranges);
}

function indexDocumentText(
  nodes: readonly WorkDocumentNode[],
): DocumentTextIndex {
  const spans: DocumentTextSpan[] = [];
  let content = '';
  const visit = (node: WorkDocumentNode): void => {
    if (node.type === 'text' && typeof node.text === 'string') {
      const from = content.length;
      content += node.text;
      spans.push({ from, to: content.length, node });
    }
    for (const child of node.content ?? []) visit(child);
  };
  for (const node of nodes) visit(node);
  return { content, spans };
}

function collectDocumentCommentAnchors(
  index: DocumentTextIndex,
): DocumentCommentAnchor[] {
  const anchors = new Map<string, DocumentCommentAnchor>();
  for (const span of index.spans) {
    for (const mark of span.node.marks ?? []) {
      const id = documentCommentId(mark);
      if (!id) continue;
      const current = anchors.get(id);
      if (current) {
        current.valid = current.valid && current.to === span.from;
        current.to = span.to;
        continue;
      }
      anchors.set(id, {
        id,
        from: span.from,
        to: span.to,
        prefix: '',
        suffix: '',
        valid: true,
      });
    }
  }
  for (const anchor of anchors.values()) {
    anchor.prefix = index.content.slice(
      Math.max(0, anchor.from - DOCUMENT_COMMENT_CONTEXT_LENGTH),
      anchor.from,
    );
    anchor.suffix = index.content.slice(
      anchor.to,
      anchor.to + DOCUMENT_COMMENT_CONTEXT_LENGTH,
    );
  }
  return Array.from(anchors.values()).sort(
    (left, right) => left.from - right.from || left.id.localeCompare(right.id),
  );
}

function documentCommentId(mark: WorkDocumentMark): string {
  if (mark.type !== DOCUMENT_COMMENT_MARK) return '';
  const id = mark.attrs?.id;
  return typeof id === 'string' ? id.trim() : '';
}

function exactTextOccurrences(content: string, value: string): number[] {
  const occurrences: number[] = [];
  let cursor = 0;
  while (cursor <= content.length - value.length) {
    const position = content.indexOf(value, cursor);
    if (position < 0) break;
    occurrences.push(position);
    cursor = position + 1;
  }
  return occurrences;
}

function uniqueCommentDestination(
  content: string,
  anchor: DocumentCommentAnchor,
  candidates: readonly number[],
): number | null {
  if (candidates.length === 1) return candidates[0] ?? null;
  const contextual = candidates.filter((candidate) => {
    const preceding = content.slice(
      Math.max(0, candidate - anchor.prefix.length),
      candidate,
    );
    const following = content.slice(
      candidate + (anchor.to - anchor.from),
      candidate + (anchor.to - anchor.from) + anchor.suffix.length,
    );
    return preceding === anchor.prefix && following === anchor.suffix;
  });
  return contextual.length === 1 ? (contextual[0] ?? null) : null;
}

function applyDocumentCommentRanges(
  nodes: readonly WorkDocumentNode[],
  ranges: readonly DocumentCommentRange[],
): WorkDocumentNode[] {
  let offset = 0;
  const rewrite = (node: WorkDocumentNode): WorkDocumentNode[] => {
    const marks = (node.marks ?? []).filter(
      (mark) => mark.type !== DOCUMENT_COMMENT_MARK,
    );
    const { content: _content, marks: _marks, ...properties } = node;
    if (node.type === 'text' && typeof node.text === 'string') {
      const from = offset;
      const to = from + node.text.length;
      offset = to;
      const boundaries = new Set([from, to]);
      for (const range of ranges) {
        if (range.from > from && range.from < to) boundaries.add(range.from);
        if (range.to > from && range.to < to) boundaries.add(range.to);
      }
      const ordered = Array.from(boundaries).sort(
        (left, right) => left - right,
      );
      const rewritten: WorkDocumentNode[] = [];
      for (let index = 0; index < ordered.length - 1; index += 1) {
        const chunkFrom = ordered[index];
        const chunkTo = ordered[index + 1];
        if (
          chunkFrom === undefined ||
          chunkTo === undefined ||
          chunkFrom === chunkTo
        ) {
          continue;
        }
        const commentMarks: WorkDocumentMark[] = ranges
          .filter((range) => range.from <= chunkFrom && range.to >= chunkTo)
          .map((range) => ({
            type: DOCUMENT_COMMENT_MARK,
            attrs: { id: range.id },
          }));
        const nextMarks = [...marks, ...commentMarks];
        rewritten.push({
          ...properties,
          text: node.text.slice(chunkFrom - from, chunkTo - from),
          ...(nextMarks.length ? { marks: nextMarks } : {}),
        });
      }
      return rewritten;
    }
    const content = (node.content ?? []).flatMap(rewrite);
    return [
      {
        ...properties,
        ...(content.length ? { content } : {}),
        ...(marks.length ? { marks } : {}),
      },
    ];
  };
  return nodes.flatMap(rewrite);
}
