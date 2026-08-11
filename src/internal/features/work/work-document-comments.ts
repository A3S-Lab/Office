import { type Editor, Mark, mergeAttributes } from '@tiptap/core';
import {
  Fragment,
  type Mark as ProseMirrorMark,
  type Node as ProseMirrorNode,
  Slice,
} from '@tiptap/pm/model';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import type {
  WorkDocumentComment,
  WorkDocumentCommentReply,
  WorkDocumentContent,
} from './work-types';

export interface WorkDocumentCommentAnchor {
  id: string;
  from: number;
  to: number;
  anchorText: string;
}

export interface WorkDocumentCommentView
  extends WorkDocumentComment,
    WorkDocumentCommentAnchor {}

export interface WorkDocumentCommentRange {
  from: number;
  to: number;
}

export interface InsertDocumentCommentOptions {
  id: string;
  range?: WorkDocumentCommentRange;
}

interface DocumentCommentOptions {
  getContent: () => WorkDocumentContent | null;
  onContentChange: (content: WorkDocumentContent) => void;
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentComment: {
      addDocumentCommentReply: (
        id: string,
        reply: WorkDocumentCommentReply,
      ) => ReturnType;
      deleteDocumentComment: (id: string) => ReturnType;
      showDocumentCommentDraft: (range: WorkDocumentCommentRange) => ReturnType;
      clearDocumentCommentDraft: () => ReturnType;
      insertDocumentComment: (
        options: InsertDocumentCommentOptions,
      ) => ReturnType;
      insertDocumentCommentThread: (
        comment: WorkDocumentComment,
        range?: WorkDocumentCommentRange,
      ) => ReturnType;
      removeDocumentComment: (id: string) => ReturnType;
      toggleDocumentCommentResolved: (id: string) => ReturnType;
    };
  }
}

const documentCommentPastePluginKey = new PluginKey('documentCommentPaste');
const documentCommentDraftPluginKey = new PluginKey<DecorationSet>(
  'documentCommentDraft',
);

export const DocumentComment = Mark.create<DocumentCommentOptions>({
  name: 'documentComment',
  priority: 1050,
  inclusive: false,
  keepOnSplit: true,
  excludes: '',

  addOptions() {
    return {
      getContent: () => null,
      onContentChange: () => undefined,
    };
  },

  addAttributes() {
    return {
      id: {
        default: '',
        parseHTML: (element) => element.getAttribute('data-comment-id') ?? '',
        renderHTML: (attributes) => ({ 'data-comment-id': attributes.id }),
      },
    };
  },

  parseHTML() {
    return [{ tag: 'span[data-document-comment]' }];
  },

  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      mergeAttributes(HTMLAttributes, { 'data-document-comment': 'true' }),
      0,
    ];
  },

  addCommands() {
    return {
      addDocumentCommentReply:
        (id, reply) =>
        ({ dispatch }) => {
          const content = this.options.getContent();
          if (
            !content ||
            !(content.comments ?? []).some((comment) => comment.id === id)
          ) {
            return false;
          }
          if (dispatch) {
            this.options.onContentChange({
              ...content,
              comments: appendDocumentCommentReply(
                content.comments ?? [],
                id,
                reply,
              ),
            });
          }
          return true;
        },
      deleteDocumentComment:
        (id) =>
        ({ state, dispatch }) => {
          const commentId = id.trim();
          const content = this.options.getContent();
          if (!commentId || !content) return false;
          const comments = content.comments ?? [];
          const hasRecord = comments.some(
            (comment) => comment.id === commentId,
          );
          const segments = documentCommentSegments(state.doc, this.type).filter(
            (segment) => segment.id === commentId,
          );
          if (!hasRecord && !segments.length) return false;
          if (!dispatch) return true;
          this.options.onContentChange({
            ...content,
            comments: removeDocumentCommentRecord(comments, commentId),
          });
          if (segments.length) {
            const transaction = state.tr;
            for (const segment of segments) {
              transaction.removeMark(segment.from, segment.to, this.type);
            }
            if (transaction.docChanged) dispatch(transaction);
          }
          return true;
        },
      showDocumentCommentDraft:
        (range) =>
        ({ state, dispatch }) => {
          const normalized = normalizeDocumentCommentRange(state.doc, range);
          if (
            !normalized ||
            !canInsertDocumentCommentInState(state.doc, this.type, normalized)
          )
            return false;
          dispatch?.(
            state.tr.setMeta(documentCommentDraftPluginKey, {
              range: normalized,
            }),
          );
          return true;
        },
      clearDocumentCommentDraft:
        () =>
        ({ state, dispatch }) => {
          if (!documentCommentDraftRangeFromState(state)) return false;
          dispatch?.(
            state.tr.setMeta(documentCommentDraftPluginKey, { range: null }),
          );
          return true;
        },
      insertDocumentComment:
        ({ id, range }) =>
        ({ state, dispatch }) => {
          const normalized = normalizeDocumentCommentRange(
            state.doc,
            range ?? state.selection,
          );
          const commentId = id.trim();
          if (
            !normalized ||
            !commentId ||
            !canInsertDocumentCommentInState(state.doc, this.type, normalized)
          )
            return false;
          const transaction = state.tr.addMark(
            normalized.from,
            normalized.to,
            this.type.create({ id: commentId }),
          );
          if (!transaction.docChanged) return false;
          transaction.setMeta(documentCommentDraftPluginKey, { range: null });
          dispatch?.(transaction);
          return true;
        },
      insertDocumentCommentThread:
        (comment, range) =>
        ({ state, dispatch }) => {
          const content = this.options.getContent();
          const normalized = normalizeDocumentCommentRange(
            state.doc,
            range ?? state.selection,
          );
          const commentId = comment.id.trim();
          if (
            !content ||
            !normalized ||
            !commentId ||
            (content.comments ?? []).some(
              (candidate) => candidate.id === commentId,
            ) ||
            !canInsertDocumentCommentInState(state.doc, this.type, normalized)
          ) {
            return false;
          }
          if (!dispatch) return true;
          this.options.onContentChange({
            ...content,
            comments: [...(content.comments ?? []), comment],
          });
          const transaction = state.tr.addMark(
            normalized.from,
            normalized.to,
            this.type.create({ id: commentId }),
          );
          transaction.setMeta(documentCommentDraftPluginKey, {
            range: null,
          });
          dispatch(transaction);
          return true;
        },
      removeDocumentComment:
        (id) =>
        ({ state, dispatch }) => {
          const commentId = id.trim();
          if (!commentId) return false;
          const segments = documentCommentSegments(state.doc, this.type).filter(
            (segment) => segment.id === commentId,
          );
          if (!segments.length) return false;
          const transaction = state.tr;
          for (const segment of segments) {
            transaction.removeMark(segment.from, segment.to, this.type);
          }
          if (!transaction.docChanged) return false;
          dispatch?.(transaction);
          return true;
        },
      toggleDocumentCommentResolved:
        (id) =>
        ({ dispatch }) => {
          const content = this.options.getContent();
          if (
            !content ||
            !(content.comments ?? []).some((comment) => comment.id === id)
          ) {
            return false;
          }
          if (dispatch) {
            this.options.onContentChange({
              ...content,
              comments: toggleDocumentCommentResolved(
                content.comments ?? [],
                id,
              ),
            });
          }
          return true;
        },
    };
  },

  addProseMirrorPlugins() {
    const type = this.type;
    return [
      new Plugin({
        key: documentCommentPastePluginKey,
        props: {
          transformPasted: (slice) =>
            stripDocumentCommentsFromSlice(slice, type),
        },
      }),
      new Plugin<DecorationSet>({
        key: documentCommentDraftPluginKey,
        state: {
          init: () => DecorationSet.empty,
          apply: (transaction, current) => {
            const mapped = current.map(transaction.mapping, transaction.doc);
            const meta = transaction.getMeta(documentCommentDraftPluginKey) as
              | {
                  range: WorkDocumentCommentRange | null;
                }
              | undefined;
            if (!meta) return mapped;
            const range = normalizeDocumentCommentRange(
              transaction.doc,
              meta.range,
            );
            if (!range) return DecorationSet.empty;
            return DecorationSet.create(transaction.doc, [
              Decoration.inline(
                range.from,
                range.to,
                {
                  class: 'work-document-comment-draft',
                  'data-document-comment-draft': 'true',
                },
                {
                  inclusiveStart: false,
                  inclusiveEnd: false,
                },
              ),
            ]);
          },
        },
        props: {
          decorations: (state) =>
            documentCommentDraftPluginKey.getState(state) ??
            DecorationSet.empty,
        },
      }),
    ];
  },
});

export function canInsertDocumentComment(
  editor: Editor,
  range: WorkDocumentCommentRange = editor.state.selection,
): boolean {
  const type = editor.schema.marks.documentComment;
  return canInsertDocumentCommentInState(editor.state.doc, type, range);
}

export function documentCommentDraftRange(
  editor: Editor,
): WorkDocumentCommentRange | null {
  return documentCommentDraftRangeFromState(editor.state);
}

function documentCommentDraftRangeFromState(
  state: Editor['state'],
): WorkDocumentCommentRange | null {
  const decorations = documentCommentDraftPluginKey.getState(state)?.find();
  if (!decorations?.length) return null;
  return {
    from: Math.min(...decorations.map((decoration) => decoration.from)),
    to: Math.max(...decorations.map((decoration) => decoration.to)),
  };
}

export function collectDocumentCommentAnchors(
  document: ProseMirrorNode,
): WorkDocumentCommentAnchor[] {
  const anchors = new Map<string, WorkDocumentCommentAnchor>();
  document.descendants((node, position) => {
    if (!node.isText || !node.text) return;
    const mark = documentCommentMark(node.marks);
    const id = mark ? stringAttribute(mark.attrs.id).trim() : '';
    if (!id) return;
    const current = anchors.get(id);
    if (current) {
      current.from = Math.min(current.from, position);
      current.to = Math.max(current.to, position + node.nodeSize);
      current.anchorText += node.text;
      return;
    }
    anchors.set(id, {
      id,
      from: position,
      to: position + node.nodeSize,
      anchorText: node.text,
    });
  });
  return Array.from(anchors.values()).sort(
    (left, right) => left.from - right.from,
  );
}

export function documentCommentViews(
  comments: readonly WorkDocumentComment[],
  anchors: readonly WorkDocumentCommentAnchor[],
): WorkDocumentCommentView[] {
  const byId = new Map(
    comments.map(
      (comment) => [comment.id, normalizeDocumentComment(comment)] as const,
    ),
  );
  return anchors.map((anchor) => ({
    ...(byId.get(anchor.id) ?? {
      id: anchor.id,
      author: '未知审阅者',
      date: '',
      text: '此批注的内容不可用。',
      resolved: false,
    }),
    ...anchor,
  }));
}

export function retainAnchoredDocumentComments(
  comments: readonly WorkDocumentComment[],
  anchors: readonly WorkDocumentCommentAnchor[],
  retainedIds: ReadonlySet<string> = new Set(),
): WorkDocumentComment[] {
  const ids = new Set(anchors.map((anchor) => anchor.id));
  return comments
    .filter((comment) => ids.has(comment.id) || retainedIds.has(comment.id))
    .map(normalizeDocumentComment);
}

export function appendDocumentCommentReply(
  comments: readonly WorkDocumentComment[],
  id: string,
  reply: WorkDocumentCommentReply,
): WorkDocumentComment[] {
  return comments.map((comment) =>
    comment.id === id
      ? {
          ...comment,
          replies: [...(comment.replies ?? []), reply],
        }
      : comment,
  );
}

export function toggleDocumentCommentResolved(
  comments: readonly WorkDocumentComment[],
  id: string,
): WorkDocumentComment[] {
  return comments.map((comment) =>
    comment.id === id ? { ...comment, resolved: !comment.resolved } : comment,
  );
}

export function removeDocumentCommentRecord(
  comments: readonly WorkDocumentComment[],
  id: string,
): WorkDocumentComment[] {
  return comments.filter((comment) => comment.id !== id);
}

export function stripDocumentCommentsFromSlice(
  slice: Slice,
  type: ProseMirrorMark['type'],
): Slice {
  return new Slice(
    stripDocumentCommentMarks(slice.content, type),
    slice.openStart,
    slice.openEnd,
  );
}

function selectionContainsDocumentComment(
  document: ProseMirrorNode,
  from: number,
  to: number,
): boolean {
  let found = false;
  document.nodesBetween(from, to, (node) => {
    if (node.isText && documentCommentMark(node.marks)) found = true;
    return !found;
  });
  return found;
}

function canInsertDocumentCommentInState(
  document: ProseMirrorNode,
  type: ProseMirrorMark['type'] | undefined,
  range: WorkDocumentCommentRange,
): boolean {
  const normalized = normalizeDocumentCommentRange(document, range);
  if (!type || !normalized) return false;
  const selectedText = document.textBetween(
    normalized.from,
    normalized.to,
    '\n',
  );
  return (
    Boolean(selectedText.trim()) &&
    !selectionContainsDocumentComment(document, normalized.from, normalized.to)
  );
}

function normalizeDocumentCommentRange(
  document: ProseMirrorNode,
  range: WorkDocumentCommentRange | null | undefined,
): WorkDocumentCommentRange | null {
  if (!range) return null;
  const from = Math.max(0, Math.min(document.content.size, range.from));
  const to = Math.max(0, Math.min(document.content.size, range.to));
  return from < to ? { from, to } : null;
}

function documentCommentSegments(
  document: ProseMirrorNode,
  type: ProseMirrorMark['type'],
): Array<{ id: string; from: number; to: number }> {
  const segments: Array<{ id: string; from: number; to: number }> = [];
  document.descendants((node, position) => {
    if (!node.isText) return;
    const mark = node.marks.find((candidate) => candidate.type === type);
    const id = mark ? stringAttribute(mark.attrs.id).trim() : '';
    if (id) segments.push({ id, from: position, to: position + node.nodeSize });
  });
  return segments;
}

function stripDocumentCommentMarks(
  fragment: Fragment,
  type: ProseMirrorMark['type'],
): Fragment {
  const nodes: ProseMirrorNode[] = [];
  fragment.forEach((node) => {
    const content = node.content.size
      ? stripDocumentCommentMarks(node.content, type)
      : node.content;
    const copy = node.content.size ? node.copy(content) : node;
    nodes.push(copy.mark(copy.marks.filter((mark) => mark.type !== type)));
  });
  return Fragment.fromArray(nodes);
}

function documentCommentMark(
  marks: readonly ProseMirrorMark[],
): ProseMirrorMark | undefined {
  return marks.find((mark) => mark.type.name === 'documentComment');
}

function normalizeDocumentComment(
  comment: WorkDocumentComment,
): WorkDocumentComment {
  return {
    id: comment.id,
    author: comment.author || '未知审阅者',
    date: comment.date || '',
    text: comment.text || '（空批注）',
    resolved: Boolean(comment.resolved),
    replies: comment.replies?.map((reply) => ({
      id: reply.id,
      author: reply.author || '未知审阅者',
      date: reply.date || '',
      text: reply.text || '（空回复）',
    })),
  };
}

function stringAttribute(value: unknown): string {
  return typeof value === 'string' ? value : '';
}
