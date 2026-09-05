import { type CommandProps, Mark, mergeAttributes } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import {
  Fragment,
  type Mark as ProseMirrorMark,
  type Node as ProseMirrorNode,
  Slice,
} from '@tiptap/pm/model';
import {
  EditorState,
  Plugin,
  PluginKey,
  TextSelection,
  type Transaction,
} from '@tiptap/pm/state';
import { trackDocumentFormattingTransaction } from './work-document-format-change-tracking';
import {
  parseDocumentCharacterFormatting,
  restoreDocumentCharacterFormatting,
} from './work-document-format-changes';
import {
  clearDocumentNumberingChangeAttributes,
  parseDocumentNumberingChange,
  restoredDocumentNumberingAttributes,
} from './work-document-numbering-changes';
import {
  clearDocumentParagraphChangeAttributes,
  parseDocumentParagraphFormatting,
  restoredDocumentParagraphAttributes,
} from './work-document-paragraph-format-changes';
import type {
  WorkDocumentChangeKind,
  WorkDocumentMoveRole,
} from './work-types';

export type {
  WorkDocumentChangeKind,
  WorkDocumentMoveRole,
} from './work-types';

export interface WorkDocumentChangeIdentity {
  id: string;
  actorId?: string;
  author: string;
  date: string;
}

export interface WorkDocumentChange extends WorkDocumentChangeIdentity {
  kind: WorkDocumentChangeKind;
  /** Present when a change is one side of a native Word move revision. */
  moveRole?: WorkDocumentMoveRole;
  from: number;
  to: number;
  text: string;
}

interface DocumentChangeOptions {
  isTracking: () => boolean;
  createChange: (kind: WorkDocumentChangeKind) => WorkDocumentChangeIdentity;
  onTrackingChange: (enabled: boolean) => void;
}

interface ChangeSegment {
  id: string;
  kind: WorkDocumentChangeKind;
  moveRole?: WorkDocumentMoveRole;
  from: number;
  to: number;
  before: string;
  text: string;
}

interface MoveSideSummary {
  text: string;
  from: number;
  to: number;
}

interface MoveChangeSummary {
  from?: MoveSideSummary;
  to?: MoveSideSummary;
}

interface ParagraphChangeSegment {
  id: string;
  kind: 'paragraph-formatting';
  position: number;
  from: number;
  to: number;
  before: string;
}

interface NumberingChangeSegment {
  id: string;
  kind: 'numbering';
  position: number;
  from: number;
  to: number;
  before: string;
}

interface BlockChangeSegment {
  id: string;
  kind: 'insertion' | 'deletion';
  position: number;
  from: number;
  to: number;
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentChange: {
      acceptAllDocumentChanges: () => ReturnType;
      acceptDocumentChange: (id: string) => ReturnType;
      acceptDocumentChanges: (ids: readonly string[]) => ReturnType;
      rejectAllDocumentChanges: () => ReturnType;
      rejectDocumentChange: (id: string) => ReturnType;
      rejectDocumentChanges: (ids: readonly string[]) => ReturnType;
      replaceDocumentTextWithTrackedChange: (
        from: number,
        to: number,
        text: string,
      ) => ReturnType;
      setDocumentTrackChanges: (enabled: boolean) => ReturnType;
      toggleDocumentTrackChanges: () => ReturnType;
    };
  }
}

const documentChangePluginKey = new PluginKey('documentChangeTracking');
const CONTINUOUS_INSERTION_WINDOW_MS = 30_000;

export const DocumentChange = Mark.create<DocumentChangeOptions>({
  name: 'documentChange',
  priority: 1100,
  inclusive: false,
  keepOnSplit: false,

  addOptions() {
    return {
      isTracking: () => false,
      createChange: () => ({
        id: createDocumentChangeId(),
        author: 'A3S Work',
        date: new Date().toISOString(),
      }),
      onTrackingChange: () => undefined,
    };
  },

  addAttributes() {
    return {
      kind: {
        default: 'insertion',
        parseHTML: (element) => {
          const declared = element.getAttribute('data-change-kind');
          if (declared === 'formatting') return 'formatting';
          if (declared === 'move') return 'move';
          return element.tagName.toLowerCase() === 'del'
            ? 'deletion'
            : 'insertion';
        },
        renderHTML: (attributes) => ({ 'data-change-kind': attributes.kind }),
      },
      moveRole: {
        default: '',
        parseHTML: (element) => {
          const role = element.getAttribute('data-change-move-role');
          return role === 'from' || role === 'to' ? role : '';
        },
        renderHTML: (attributes) =>
          attributes.moveRole === 'from' || attributes.moveRole === 'to'
            ? { 'data-change-move-role': attributes.moveRole }
            : {},
      },
      id: {
        default: '',
        parseHTML: (element) => element.getAttribute('data-change-id') ?? '',
        renderHTML: (attributes) => ({ 'data-change-id': attributes.id }),
      },
      actorId: {
        default: '',
        parseHTML: (element) =>
          element.getAttribute('data-change-actor-id') ?? '',
        renderHTML: (attributes) =>
          attributes.actorId
            ? { 'data-change-actor-id': attributes.actorId }
            : {},
      },
      author: {
        default: '',
        parseHTML: (element) =>
          element.getAttribute('data-change-author') ?? '',
        renderHTML: (attributes) => ({
          'data-change-author': attributes.author,
        }),
      },
      date: {
        default: '',
        parseHTML: (element) => element.getAttribute('data-change-date') ?? '',
        renderHTML: (attributes) => ({ 'data-change-date': attributes.date }),
      },
      before: {
        default: '',
        parseHTML: (element) =>
          element.getAttribute('data-change-before') ?? '',
        renderHTML: (attributes) =>
          attributes.before ? { 'data-change-before': attributes.before } : {},
      },
    };
  },

  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          paragraphChangeKind: paragraphChangeAttribute(
            'kind',
            'data-change-kind',
          ),
          paragraphChangeId: paragraphChangeAttribute('id', 'data-change-id'),
          paragraphChangeActorId: paragraphChangeAttribute(
            'actorId',
            'data-change-actor-id',
          ),
          paragraphChangeAuthor: paragraphChangeAttribute(
            'author',
            'data-change-author',
          ),
          paragraphChangeDate: paragraphChangeAttribute(
            'date',
            'data-change-date',
          ),
          paragraphChangeBefore: paragraphChangeAttribute(
            'before',
            'data-change-before',
          ),
          blockChangeKind: blockChangeAttribute('kind'),
          blockChangeId: blockChangeAttribute('id'),
          blockChangeActorId: blockChangeAttribute('actorId'),
          blockChangeAuthor: blockChangeAttribute('author'),
          blockChangeDate: blockChangeAttribute('date'),
        },
      },
      {
        types: ['orderedList'],
        attributes: {
          numberingChangeKind: numberingChangeAttribute(
            'kind',
            'data-change-kind',
          ),
          numberingChangeId: numberingChangeAttribute('id', 'data-change-id'),
          numberingChangeActorId: numberingChangeAttribute(
            'actorId',
            'data-change-actor-id',
          ),
          numberingChangeAuthor: numberingChangeAttribute(
            'author',
            'data-change-author',
          ),
          numberingChangeDate: numberingChangeAttribute(
            'date',
            'data-change-date',
          ),
          numberingChangeBefore: numberingChangeAttribute(
            'before',
            'data-change-before',
          ),
        },
      },
    ];
  },

  parseHTML() {
    return [
      { tag: 'ins[data-document-change]' },
      { tag: 'del[data-document-change]' },
      {
        tag: 'span[data-document-change][data-change-kind="formatting"]',
      },
    ];
  },

  renderHTML({ mark, HTMLAttributes }) {
    const tag =
      mark.attrs.kind === 'deletion'
        ? 'del'
        : mark.attrs.kind === 'move' && mark.attrs.moveRole === 'from'
          ? 'del'
          : mark.attrs.kind === 'formatting'
            ? 'span'
            : 'ins';
    return [
      tag,
      mergeAttributes(HTMLAttributes, { 'data-document-change': 'true' }),
      0,
    ];
  },

  addCommands() {
    return {
      acceptAllDocumentChanges: () => (props) =>
        resolveDocumentChangesCommand(props, this.type, 'accept') > 0,
      acceptDocumentChange: (id) => (props) =>
        resolveDocumentChangesCommand(
          props,
          this.type,
          'accept',
          new Set([id]),
        ) > 0,
      acceptDocumentChanges: (ids) => (props) =>
        resolveDocumentChangesCommand(
          props,
          this.type,
          'accept',
          new Set(ids),
        ) > 0,
      rejectAllDocumentChanges: () => (props) =>
        resolveDocumentChangesCommand(props, this.type, 'reject') > 0,
      rejectDocumentChange: (id) => (props) =>
        resolveDocumentChangesCommand(
          props,
          this.type,
          'reject',
          new Set([id]),
        ) > 0,
      rejectDocumentChanges: (ids) => (props) =>
        resolveDocumentChangesCommand(
          props,
          this.type,
          'reject',
          new Set(ids),
        ) > 0,
      replaceDocumentTextWithTrackedChange:
        (from, to, text) =>
        ({ state, tr }) => {
          if (from < 0 || to < from || to > state.doc.content.size)
            return false;
          trackedReplacement(
            tr,
            state.doc,
            this.type,
            from,
            to,
            text,
            this.options.createChange,
          );
          return tr.docChanged;
        },
      setDocumentTrackChanges: (enabled) => () => {
        if (enabled !== this.options.isTracking()) {
          this.options.onTrackingChange(enabled);
        }
        return true;
      },
      toggleDocumentTrackChanges: () => () => {
        this.options.onTrackingChange(!this.options.isTracking());
        return true;
      },
    };
  },

  addProseMirrorPlugins() {
    const changeType = this.type;
    const options = this.options;
    return [
      new Plugin({
        key: documentChangePluginKey,
        filterTransaction: (transaction, state) => {
          trackDocumentFormattingTransaction(
            transaction,
            state,
            changeType,
            {
              isTracking: options.isTracking,
              createChange: (kind) => {
                const identity = options.createChange(kind);
                return {
                  ...identity,
                  id: identity.id || createDocumentChangeId(),
                };
              },
            },
            documentChangePluginKey,
          );
          return true;
        },
        props: {
          handleTextInput: (view, from, to, text) => {
            if (!options.isTracking()) return false;
            const transaction = trackedReplacement(
              view.state.tr,
              view.state.doc,
              changeType,
              from,
              to,
              text,
              options.createChange,
            );
            view.dispatch(transaction);
            return true;
          },
          handleKeyDown: (view, event) => {
            if (
              !options.isTracking() ||
              event.isComposing ||
              event.metaKey ||
              event.ctrlKey ||
              (event.key !== 'Backspace' && event.key !== 'Delete')
            ) {
              return false;
            }
            const selection = view.state.selection;
            const range =
              selection.from !== selection.to
                ? { from: selection.from, to: selection.to }
                : adjacentTextRange(
                    view.state.doc,
                    selection.from,
                    event.key === 'Backspace' ? -1 : 1,
                  );
            if (!range) return false;
            const transaction = view.state.tr;
            const changed = trackDeletion(
              transaction,
              view.state.doc,
              changeType,
              range.from,
              range.to,
              changeMark(changeType, 'deletion', options.createChange),
            );
            if (!changed) return false;
            const cursor = Math.min(
              transaction.doc.content.size,
              transaction.mapping.map(range.from),
            );
            transaction.setSelection(
              TextSelection.near(transaction.doc.resolve(cursor)),
            );
            view.dispatch(transaction);
            return true;
          },
          handlePaste: (view, _event, slice) => {
            if (!options.isTracking()) return false;
            const insertion = changeMark(
              changeType,
              'insertion',
              options.createChange,
            );
            const markedSlice = new Slice(
              markFragment(slice.content, insertion, changeType),
              slice.openStart,
              slice.openEnd,
            );
            const { from, to } = view.state.selection;
            const transaction = view.state.tr;
            if (from !== to) {
              trackDeletion(
                transaction,
                view.state.doc,
                changeType,
                from,
                to,
                changeMark(changeType, 'deletion', options.createChange),
              );
            }
            const position = transaction.mapping.map(to, -1);
            transaction.replace(position, position, markedSlice);
            const cursor = Math.min(
              transaction.doc.content.size,
              position + markedSlice.size,
            );
            transaction.setSelection(
              TextSelection.near(transaction.doc.resolve(cursor)),
            );
            view.dispatch(transaction);
            return true;
          },
        },
      }),
    ];
  },
});

export function collectDocumentChanges(
  document: ProseMirrorNode,
): WorkDocumentChange[] {
  const changes = new Map<string, WorkDocumentChange>();
  const moveTextParts = new Map<string, MoveChangeSummary>();
  document.descendants((node, position) => {
    if (
      (node.type.name === 'paragraph' || node.type.name === 'heading') &&
      (node.attrs.blockChangeKind === 'insertion' ||
        node.attrs.blockChangeKind === 'deletion')
    ) {
      const kind = node.attrs.blockChangeKind as 'insertion' | 'deletion';
      const id =
        stringAttribute(node.attrs.blockChangeId) ||
        `block-change-at-${position}`;
      const key = `${kind}:${id}`;
      const from = position;
      const to = position + node.nodeSize;
      changes.set(key, {
        id,
        kind,
        ...(stringAttribute(node.attrs.blockChangeActorId)
          ? { actorId: stringAttribute(node.attrs.blockChangeActorId) }
          : {}),
        author: stringAttribute(node.attrs.blockChangeAuthor) || '未知审阅者',
        date: stringAttribute(node.attrs.blockChangeDate),
        from,
        to,
        text: node.textContent,
      });
    }
    if (
      (node.type.name === 'paragraph' || node.type.name === 'heading') &&
      node.attrs.paragraphChangeKind === 'paragraph-formatting'
    ) {
      const id =
        stringAttribute(node.attrs.paragraphChangeId) ||
        `paragraph-change-at-${position}`;
      const key = `paragraph-formatting:${id}`;
      const from = position + 1;
      const to = from + node.content.size;
      const current = changes.get(key);
      if (current) {
        current.from = Math.min(current.from, from);
        current.to = Math.max(current.to, to);
        current.text = `${current.text}\n${node.textContent}`;
      } else {
        changes.set(key, {
          id,
          kind: 'paragraph-formatting',
          ...(stringAttribute(node.attrs.paragraphChangeActorId)
            ? { actorId: stringAttribute(node.attrs.paragraphChangeActorId) }
            : {}),
          author:
            stringAttribute(node.attrs.paragraphChangeAuthor) || '未知审阅者',
          date: stringAttribute(node.attrs.paragraphChangeDate),
          from,
          to,
          text: node.textContent,
        });
      }
    }
    if (
      node.type.name === 'orderedList' &&
      node.attrs.numberingChangeKind === 'numbering'
    ) {
      const id =
        stringAttribute(node.attrs.numberingChangeId) ||
        `numbering-change-at-${position}`;
      const key = `numbering:${id}`;
      const from = position + 1;
      const to = from + node.content.size;
      const current = changes.get(key);
      if (current) {
        current.from = Math.min(current.from, from);
        current.to = Math.max(current.to, to);
        current.text = `${current.text}\n${node.textContent}`;
      } else {
        changes.set(key, {
          id,
          kind: 'numbering',
          ...(stringAttribute(node.attrs.numberingChangeActorId)
            ? { actorId: stringAttribute(node.attrs.numberingChangeActorId) }
            : {}),
          author:
            stringAttribute(node.attrs.numberingChangeAuthor) || '未知审阅者',
          date: stringAttribute(node.attrs.numberingChangeDate),
          from,
          to,
          text: node.textContent,
        });
      }
    }
    if (!node.isText || !node.text) return;
    const mark = documentChangeMark(node.marks);
    if (!mark) return;
    const kind = changeKind(mark.attrs.kind);
    const moveRole = documentMoveRole(mark.attrs.moveRole);
    const id = stringAttribute(mark.attrs.id) || `change-at-${position}`;
    const key = `${kind}:${id}`;
    // WPS/Word may assign a different revision id to the paragraph mark and
    // to the text wrapper inside that paragraph.  Once a text node is inside
    // any native block-change range, expose the block as the single review
    // item instead of leaking a duplicate inline change with the other id.
    if (positionIsInsideBlockChange(document, position)) {
      return;
    }
    const current = changes.get(key);
    if (current) {
      current.from = Math.min(current.from, position);
      current.to = Math.max(current.to, position + node.nodeSize);
      if (kind === 'move' && moveRole) {
        const parts = moveTextParts.get(key) ?? {};
        const previous = parts[moveRole];
        parts[moveRole] = previous
          ? {
              text: previous.text + node.text,
              from: Math.min(previous.from, position),
              to: Math.max(previous.to, position + node.nodeSize),
            }
          : { text: node.text, from: position, to: position + node.nodeSize };
        moveTextParts.set(key, parts);
        if (current.moveRole !== moveRole) {
          // A paired move has one review identity but two physical ranges. Do
          // not expose a misleading single-side role once both are collected.
          delete current.moveRole;
        }
      } else {
        current.text += node.text;
      }
      return;
    }
    if (kind === 'move' && moveRole) {
      const parts = moveTextParts.get(key) ?? {};
      parts[moveRole] = {
        text: node.text,
        from: position,
        to: position + node.nodeSize,
      };
      moveTextParts.set(key, parts);
    }
    changes.set(key, {
      id,
      kind,
      ...(moveRole ? { moveRole } : {}),
      ...(stringAttribute(mark.attrs.actorId)
        ? { actorId: stringAttribute(mark.attrs.actorId) }
        : {}),
      author: stringAttribute(mark.attrs.author) || '未知审阅者',
      date: stringAttribute(mark.attrs.date),
      from: position,
      to: position + node.nodeSize,
      text: node.text ?? '',
    });
  });
  for (const [key, parts] of moveTextParts) {
    const change = changes.get(key);
    if (!change) continue;
    const source = parts.from;
    const destination = parts.to;
    if (source && destination) {
      change.text =
        source.text === destination.text
          ? source.text
          : `${source.text} → ${destination.text}`;
      // A move has two disjoint ranges. Navigation targets the destination
      // when available so the review pane never selects unrelated content in
      // between the source and destination.
      change.from = destination.from;
      change.to = destination.to;
    } else {
      const only = source ?? destination;
      if (only) {
        change.text = only.text;
        change.from = only.from;
        change.to = only.to;
      }
    }
  }
  return Array.from(changes.values()).sort(
    (left, right) => left.from - right.from,
  );
}

function positionIsInsideBlockChange(
  document: ProseMirrorNode,
  position: number,
): boolean {
  const resolved = document.resolve(position);
  for (let depth = resolved.depth; depth > 0; depth -= 1) {
    const ancestor = resolved.node(depth);
    if (
      (ancestor.type.name === 'paragraph' ||
        ancestor.type.name === 'heading') &&
      (ancestor.attrs.blockChangeKind === 'insertion' ||
        ancestor.attrs.blockChangeKind === 'deletion')
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Resolves every revision on an immutable document snapshot. A malformed
 * formatting snapshot returns `null` instead of approximating the baseline.
 */
export function resolveAllDocumentChanges(
  document: ProseMirrorNode,
  decision: 'accept' | 'reject',
): ProseMirrorNode | null {
  if (!collectDocumentChanges(document).length) return document;
  const type = document.type.schema.marks.documentChange;
  if (!type) return null;
  const state = EditorState.create({ doc: document });
  const transaction = state.tr;
  return resolveDocumentChangesTransaction(state, transaction, type, decision) >
    0
    ? transaction.doc
    : null;
}

function resolveDocumentChangesCommand(
  { state, tr }: CommandProps,
  type: ProseMirrorMark['type'],
  decision: 'accept' | 'reject',
  ids?: Set<string>,
): number {
  return resolveDocumentChangesTransaction(state, tr, type, decision, ids);
}

function resolveDocumentChangesTransaction(
  state: CommandProps['state'],
  tr: Transaction,
  type: ProseMirrorMark['type'],
  decision: 'accept' | 'reject',
  ids?: Set<string>,
): number {
  const allBlockSegments = blockChangeSegments(state.doc);
  const blockSegments = allBlockSegments.filter(
    (segment) => !ids || ids.has(segment.id),
  );
  const segments = documentChangeSegments(state.doc, type).filter(
    (segment) =>
      !ids ||
      ids.has(segment.id) ||
      blockSegments.some((block) => segmentIsInsideBlock(segment, block)),
  );
  const paragraphSegments = paragraphChangeSegments(state.doc).filter(
    (segment) => !ids || ids.has(segment.id),
  );
  const numberingSegments = numberingChangeSegments(state.doc).filter(
    (segment) => !ids || ids.has(segment.id),
  );
  if (
    !segments.length &&
    !paragraphSegments.length &&
    !numberingSegments.length &&
    !blockSegments.length
  )
    return 0;
  if (
    paragraphSegments.some(
      (segment) => !parseDocumentParagraphFormatting(segment.before),
    ) ||
    numberingSegments.some(
      (segment) => !parseDocumentNumberingChange(segment.before),
    ) ||
    (decision === 'reject' &&
      segments.some(
        (segment) =>
          segment.kind === 'formatting' &&
          !parseDocumentCharacterFormatting(segment.before),
      )) ||
    segments.some(
      (segment) =>
        segment.kind === 'move' &&
        segment.moveRole !== 'from' &&
        segment.moveRole !== 'to',
    ) ||
    !validMoveSegments(segments)
  ) {
    return 0;
  }
  closeHistory(tr);
  tr.setMeta(documentChangePluginKey, { decision });
  const markRemovals: ChangeSegment[] = [];
  const contentDeletions: ChangeSegment[] = [];
  const formattingRejections: ChangeSegment[] = [];
  const removedBlockIds = new Set(
    blockSegments
      .filter(
        (segment) =>
          (decision === 'accept' && segment.kind === 'deletion') ||
          (decision === 'reject' && segment.kind === 'insertion'),
      )
      .map((segment) => segment.id),
  );
  const removedBlockSegments = blockSegments.filter((segment) =>
    removedBlockIds.has(segment.id),
  );
  for (const segment of segments) {
    if (
      removedBlockIds.has(segment.id) ||
      removedBlockSegments.some((block) => segmentIsInsideBlock(segment, block))
    ) {
      continue;
    }
    if (segment.kind === 'formatting') {
      markRemovals.push(segment);
      if (decision === 'reject') formattingRejections.push(segment);
      continue;
    }
    if (segment.kind === 'move') {
      const source = segment.moveRole === 'from';
      const removeContent = decision === 'accept' ? source : !source;
      (removeContent ? contentDeletions : markRemovals).push(segment);
      continue;
    }
    const removeMark =
      (decision === 'accept' && segment.kind === 'insertion') ||
      (decision === 'reject' && segment.kind === 'deletion');
    (removeMark ? markRemovals : contentDeletions).push(segment);
  }
  for (const segment of formattingRejections) {
    restoreDocumentCharacterFormatting(
      tr,
      state.schema,
      segment.from,
      segment.to,
      segment.before,
    );
  }
  for (const segment of paragraphSegments) {
    const node = tr.doc.nodeAt(segment.position);
    if (
      !node ||
      (node.type.name !== 'paragraph' && node.type.name !== 'heading')
    )
      return 0;
    const attributes =
      decision === 'reject'
        ? restoredDocumentParagraphAttributes(node.attrs, segment.before)
        : clearDocumentParagraphChangeAttributes(node.attrs);
    if (!attributes) return 0;
    tr.setNodeMarkup(segment.position, undefined, attributes);
  }
  for (const segment of numberingSegments) {
    const position = tr.mapping.map(segment.position);
    const node = tr.doc.nodeAt(position);
    if (!node || node.type.name !== 'orderedList') return 0;
    const attributes =
      decision === 'reject'
        ? restoredDocumentNumberingAttributes(node.attrs, segment.before)
        : clearDocumentNumberingChangeAttributes(node.attrs);
    if (!attributes) return 0;
    tr.setNodeMarkup(position, undefined, attributes);
  }
  for (const segment of blockSegments) {
    if (removedBlockIds.has(segment.id)) continue;
    const position = tr.mapping.map(segment.position);
    const node = tr.doc.nodeAt(position);
    if (
      !node ||
      (node.type.name !== 'paragraph' && node.type.name !== 'heading')
    ) {
      return 0;
    }
    tr.setNodeMarkup(
      position,
      undefined,
      clearDocumentBlockChangeAttributes(node.attrs),
    );
  }
  for (const segment of markRemovals) {
    tr.removeMark(segment.from, segment.to, type);
  }
  for (const segment of contentDeletions.sort(
    (left, right) => right.from - left.from,
  )) {
    tr.delete(segment.from, segment.to);
  }
  for (const segment of blockSegments
    .filter((candidate) => removedBlockIds.has(candidate.id))
    .sort((left, right) => right.position - left.position)) {
    const from = tr.mapping.map(segment.position);
    const node = tr.doc.nodeAt(from);
    if (
      !node ||
      (node.type.name !== 'paragraph' && node.type.name !== 'heading')
    ) {
      return 0;
    }
    tr.delete(from, from + node.nodeSize);
  }
  return tr.docChanged
    ? new Set(
        [
          ...segments,
          ...paragraphSegments,
          ...numberingSegments,
          ...blockSegments,
        ].map((segment) => segment.id),
      ).size
    : 0;
}

function segmentIsInsideBlock(
  segment: Pick<ChangeSegment, 'from' | 'to'>,
  block: Pick<BlockChangeSegment, 'from' | 'to'>,
): boolean {
  return segment.from >= block.from && segment.to <= block.to;
}

function trackedReplacement(
  transaction: Transaction,
  document: ProseMirrorNode,
  type: ProseMirrorMark['type'],
  from: number,
  to: number,
  text: string,
  createChange: (kind: WorkDocumentChangeKind) => WorkDocumentChangeIdentity,
): Transaction {
  if (from !== to) {
    trackDeletion(
      transaction,
      document,
      type,
      from,
      to,
      changeMark(type, 'deletion', createChange),
    );
  }
  const position = transaction.mapping.map(to, -1);
  if (text) {
    const insertion =
      from === to
        ? continuousInsertionMark(document, type, position, createChange)
        : changeMark(type, 'insertion', createChange);
    transaction.insertText(text, position);
    transaction.addMark(position, position + text.length, insertion);
  }
  const cursor = Math.min(transaction.doc.content.size, position + text.length);
  transaction.setSelection(TextSelection.near(transaction.doc.resolve(cursor)));
  return transaction;
}

function trackDeletion(
  transaction: Transaction,
  document: ProseMirrorNode,
  type: ProseMirrorMark['type'],
  from: number,
  to: number,
  deletion: ProseMirrorMark,
): boolean {
  const segments = textSegments(document, type, from, to);
  if (!segments.length) return false;
  for (const segment of segments) {
    if (!segment.kind) transaction.addMark(segment.from, segment.to, deletion);
  }
  for (const segment of segments
    .filter((segment) => segment.kind === 'insertion')
    .sort((left, right) => right.from - left.from)) {
    transaction.delete(segment.from, segment.to);
  }
  return true;
}

function textSegments(
  document: ProseMirrorNode,
  type: ProseMirrorMark['type'],
  from: number,
  to: number,
): Array<{ from: number; to: number; kind: WorkDocumentChangeKind | null }> {
  const segments: Array<{
    from: number;
    to: number;
    kind: WorkDocumentChangeKind | null;
  }> = [];
  document.nodesBetween(from, to, (node, position) => {
    if (!node.isText) return;
    const start = Math.max(from, position);
    const end = Math.min(to, position + node.nodeSize);
    if (start >= end) return;
    const mark = node.marks.find((candidate) => candidate.type === type);
    segments.push({
      from: start,
      to: end,
      kind: mark ? changeKind(mark.attrs.kind) : null,
    });
  });
  return segments;
}

function documentChangeSegments(
  document: ProseMirrorNode,
  type: ProseMirrorMark['type'],
): ChangeSegment[] {
  const segments: ChangeSegment[] = [];
  document.descendants((node, position) => {
    if (!node.isText) return;
    const mark = node.marks.find((candidate) => candidate.type === type);
    if (!mark) return;
    const moveRole = documentMoveRole(mark.attrs.moveRole);
    segments.push({
      id: stringAttribute(mark.attrs.id) || `change-at-${position}`,
      kind: changeKind(mark.attrs.kind),
      ...(moveRole ? { moveRole } : {}),
      from: position,
      to: position + node.nodeSize,
      before: stringAttribute(mark.attrs.before),
      text: node.text ?? '',
    });
  });
  return segments;
}

function validMoveSegments(segments: readonly ChangeSegment[]): boolean {
  const moves = new Map<string, { from: string; to: string }>();
  for (const segment of segments) {
    if (segment.kind !== 'move') continue;
    if (segment.moveRole !== 'from' && segment.moveRole !== 'to') return false;
    const current = moves.get(segment.id) ?? { from: '', to: '' };
    current[segment.moveRole] += segment.text;
    moves.set(segment.id, current);
  }
  return Array.from(moves.values()).every(({ from, to }) =>
    Boolean(from && to && from === to),
  );
}

function paragraphChangeSegments(
  document: ProseMirrorNode,
): ParagraphChangeSegment[] {
  const segments: ParagraphChangeSegment[] = [];
  document.descendants((node, position) => {
    if (
      (node.type.name !== 'paragraph' && node.type.name !== 'heading') ||
      node.attrs.paragraphChangeKind !== 'paragraph-formatting'
    ) {
      return;
    }
    segments.push({
      id:
        stringAttribute(node.attrs.paragraphChangeId) ||
        `paragraph-change-at-${position}`,
      kind: 'paragraph-formatting',
      position,
      from: position + 1,
      to: position + 1 + node.content.size,
      before: stringAttribute(node.attrs.paragraphChangeBefore),
    });
  });
  return segments;
}

function numberingChangeSegments(
  document: ProseMirrorNode,
): NumberingChangeSegment[] {
  const segments: NumberingChangeSegment[] = [];
  document.descendants((node, position) => {
    if (
      node.type.name !== 'orderedList' ||
      node.attrs.numberingChangeKind !== 'numbering'
    ) {
      return;
    }
    segments.push({
      id:
        stringAttribute(node.attrs.numberingChangeId) ||
        `numbering-change-at-${position}`,
      kind: 'numbering',
      position,
      from: position + 1,
      to: position + 1 + node.content.size,
      before: stringAttribute(node.attrs.numberingChangeBefore),
    });
  });
  return segments;
}

function blockChangeSegments(document: ProseMirrorNode): BlockChangeSegment[] {
  const segments: BlockChangeSegment[] = [];
  document.descendants((node, position) => {
    if (
      (node.type.name !== 'paragraph' && node.type.name !== 'heading') ||
      (node.attrs.blockChangeKind !== 'insertion' &&
        node.attrs.blockChangeKind !== 'deletion')
    ) {
      return;
    }
    segments.push({
      id:
        stringAttribute(node.attrs.blockChangeId) ||
        `block-change-at-${position}`,
      kind: node.attrs.blockChangeKind as 'insertion' | 'deletion',
      position,
      from: position + 1,
      to: position + 1 + node.content.size,
    });
  });
  return segments;
}

function markFragment(
  fragment: Fragment,
  insertion: ProseMirrorMark,
  type: ProseMirrorMark['type'],
): Fragment {
  const nodes: ProseMirrorNode[] = [];
  fragment.forEach((node) => {
    if (node.isText) {
      nodes.push(
        node.mark([
          ...node.marks.filter((mark) => mark.type !== type),
          insertion,
        ]),
      );
    } else if (node.content.size) {
      nodes.push(node.copy(markFragment(node.content, insertion, type)));
    } else {
      nodes.push(node);
    }
  });
  return Fragment.fromArray(nodes);
}

function adjacentTextRange(
  document: ProseMirrorNode,
  position: number,
  direction: -1 | 1,
): { from: number; to: number } | null {
  const resolved = document.resolve(position);
  const node = direction < 0 ? resolved.nodeBefore : resolved.nodeAfter;
  if (!node?.isText || !node.text) return null;
  const character =
    direction < 0 ? Array.from(node.text).at(-1) : Array.from(node.text).at(0);
  if (!character) return null;
  return direction < 0
    ? { from: position - character.length, to: position }
    : { from: position, to: position + character.length };
}

function changeMark(
  type: ProseMirrorMark['type'],
  kind: WorkDocumentChangeKind,
  createChange: (kind: WorkDocumentChangeKind) => WorkDocumentChangeIdentity,
  before = '',
): ProseMirrorMark {
  const identity = createChange(kind);
  return type.create({
    kind,
    id: identity.id || createDocumentChangeId(),
    actorId: identity.actorId ?? '',
    author: identity.author || 'A3S Work',
    date: identity.date || new Date().toISOString(),
    before,
  });
}

function continuousInsertionMark(
  document: ProseMirrorNode,
  type: ProseMirrorMark['type'],
  position: number,
  createChange: (kind: WorkDocumentChangeKind) => WorkDocumentChangeIdentity,
): ProseMirrorMark {
  const next = changeMark(type, 'insertion', createChange);
  const previousNode = document.resolve(position).nodeBefore;
  const previous = previousNode?.marks.find(
    (mark) => mark.type === type && changeKind(mark.attrs.kind) === 'insertion',
  );
  if (!previous) return next;
  if (
    stringAttribute(previous.attrs.author) !==
    stringAttribute(next.attrs.author)
  )
    return next;
  const previousTime = Date.parse(stringAttribute(previous.attrs.date));
  const nextTime = Date.parse(stringAttribute(next.attrs.date));
  if (
    !Number.isFinite(previousTime) ||
    !Number.isFinite(nextTime) ||
    Math.abs(nextTime - previousTime) > CONTINUOUS_INSERTION_WINDOW_MS
  ) {
    return next;
  }
  return previous;
}

function documentChangeMark(
  marks: readonly ProseMirrorMark[],
): ProseMirrorMark | undefined {
  return marks.find((mark) => mark.type.name === 'documentChange');
}

function changeKind(value: unknown): WorkDocumentChangeKind {
  if (value === 'deletion' || value === 'formatting' || value === 'move') {
    return value;
  }
  return 'insertion';
}

function documentMoveRole(value: unknown): WorkDocumentMoveRole | null {
  return value === 'from' || value === 'to' ? value : null;
}

function paragraphChangeAttribute(
  field: 'kind' | 'id' | 'actorId' | 'author' | 'date' | 'before',
  htmlName:
    | 'data-change-kind'
    | 'data-change-id'
    | 'data-change-actor-id'
    | 'data-change-author'
    | 'data-change-date'
    | 'data-change-before',
) {
  const modelName = `paragraphChange${field[0]?.toUpperCase() ?? ''}${field.slice(
    1,
  )}`;
  return {
    default: field === 'kind' ? null : '',
    parseHTML: (element: HTMLElement) => {
      if (
        element.getAttribute('data-document-change') !== 'true' ||
        element.getAttribute('data-change-kind') !== 'paragraph-formatting'
      ) {
        return field === 'kind' ? null : '';
      }
      return field === 'kind'
        ? 'paragraph-formatting'
        : (element.getAttribute(htmlName) ?? '');
    },
    renderHTML: (attributes: Record<string, unknown>) => {
      if (attributes.paragraphChangeKind !== 'paragraph-formatting') return {};
      if (field === 'kind') {
        return {
          'data-document-change': 'true',
          'data-change-kind': 'paragraph-formatting',
        };
      }
      const value = stringAttribute(attributes[modelName]);
      return value ? { [htmlName]: value } : {};
    },
  };
}

function numberingChangeAttribute(
  field: 'kind' | 'id' | 'actorId' | 'author' | 'date' | 'before',
  htmlName:
    | 'data-change-kind'
    | 'data-change-id'
    | 'data-change-actor-id'
    | 'data-change-author'
    | 'data-change-date'
    | 'data-change-before',
) {
  const modelName = `numberingChange${field[0]?.toUpperCase() ?? ''}${field.slice(
    1,
  )}`;
  return {
    default: field === 'kind' ? null : '',
    parseHTML: (element: HTMLElement) => {
      if (
        element.getAttribute('data-document-change') !== 'true' ||
        element.getAttribute('data-change-kind') !== 'numbering'
      ) {
        return field === 'kind' ? null : '';
      }
      return field === 'kind'
        ? 'numbering'
        : (element.getAttribute(htmlName) ?? '');
    },
    renderHTML: (attributes: Record<string, unknown>) => {
      if (attributes.numberingChangeKind !== 'numbering') return {};
      if (field === 'kind') {
        return {
          'data-document-change': 'true',
          'data-change-kind': 'numbering',
        };
      }
      const value = stringAttribute(attributes[modelName]);
      return value ? { [htmlName]: value } : {};
    },
  };
}

function blockChangeAttribute(
  field: 'kind' | 'id' | 'actorId' | 'author' | 'date',
) {
  const modelName = `blockChange${field[0]?.toUpperCase() ?? ''}${field.slice(
    1,
  )}`;
  const htmlName =
    field === 'kind'
      ? 'data-block-change-kind'
      : `data-block-change-${field.replace('actorId', 'actor-id')}`;
  return {
    default: field === 'kind' ? null : '',
    parseHTML: (element: HTMLElement) => {
      if (element.getAttribute('data-document-block-change') !== 'true') {
        return field === 'kind' ? null : '';
      }
      const value = element.getAttribute(htmlName) ?? '';
      if (field === 'kind') {
        return value === 'insertion' || value === 'deletion' ? value : null;
      }
      return value;
    },
    renderHTML: (attributes: Record<string, unknown>) => {
      const kind = attributes.blockChangeKind;
      if (kind !== 'insertion' && kind !== 'deletion') return {};
      if (field === 'kind') {
        return {
          'data-document-block-change': 'true',
          'data-block-change-kind': kind,
        };
      }
      const value = stringAttribute(attributes[modelName]);
      return value ? { [htmlName]: value } : {};
    },
  };
}

function clearDocumentBlockChangeAttributes(
  attributes: Record<string, unknown>,
): Record<string, unknown> {
  return {
    ...attributes,
    blockChangeKind: null,
    blockChangeId: '',
    blockChangeActorId: '',
    blockChangeAuthor: '',
    blockChangeDate: '',
  };
}

function stringAttribute(value: unknown): string {
  return typeof value === 'string' ? value : '';
}

function createDocumentChangeId(): string {
  const random =
    typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
      ? crypto.randomUUID()
      : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
  return `change-${random}`;
}
