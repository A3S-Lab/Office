import {
  type CommandProps,
  type Editor,
  mergeAttributes,
  Node,
} from '@tiptap/core';
import type {
  Mark as ProseMirrorMark,
  Node as ProseMirrorNode,
} from '@tiptap/pm/model';
import {
  type EditorState,
  Plugin,
  TextSelection,
  type Transaction,
} from '@tiptap/pm/state';
import { Mapping } from '@tiptap/pm/transform';
import { createWorkId } from './work-templates';

export type WorkDocumentBookmarkBoundaryKind = 'start' | 'end';

export interface WorkDocumentBookmark {
  id: string;
  name: string;
  nativeId: number;
  from: number;
  to: number;
}

interface WorkDocumentBookmarkIdentity {
  id: string;
  name: string;
  nativeId: number;
}

interface BookmarkBoundaryAtPosition {
  node: ProseMirrorNode;
  position: number;
  kind: WorkDocumentBookmarkBoundaryKind;
}

interface BookmarkPairAtPosition extends WorkDocumentBookmark {
  start: BookmarkBoundaryAtPosition;
  end: BookmarkBoundaryAtPosition;
}

interface BookmarkPairCollection {
  pairs: BookmarkPairAtPosition[];
  orphans: BookmarkBoundaryAtPosition[];
}

interface BookmarkRegistry {
  ids: Set<string>;
  names: Set<string>;
  nativeIds: Set<number>;
}

interface BookmarkRename {
  from: number;
  to: number;
  previousName: string;
  nextName: string;
}

const BOOKMARK_NAME_PATTERN = /^[\p{L}_][\p{L}\p{N}_]*$/u;
const BOOKMARK_UI_NAME_PATTERN = /^[\p{L}][\p{L}\p{N}_]*$/u;
const MAX_BOOKMARK_NAME_LENGTH = 40;
const MAX_BOOKMARK_NATIVE_ID = 0x7fff_ffff;
const MISSING_LINK_CLASS = 'work-document-link-missing';

export const DOCUMENT_BOOKMARK_VALIDATION_MESSAGE =
  '书签名称需以字母开头，只能包含字母、数字和下划线，且不超过 40 个字符。';
export const DOCUMENT_BOOKMARK_DUPLICATE_MESSAGE = '文档中已存在同名书签。';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentBookmark: {
      insertDocumentBookmark: (name: string) => ReturnType;
      deleteDocumentBookmark: (id: string) => ReturnType;
    };
  }
}

export const DocumentBookmarkBoundary = Node.create({
  name: 'documentBookmarkBoundary',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: false,

  addCommands() {
    return {
      insertDocumentBookmark: (name) => (props) =>
        insertDocumentBookmarkCommand(props, name),
      deleteDocumentBookmark: (id) => (props) =>
        deleteDocumentBookmarkCommand(props, id),
    };
  },

  addProseMirrorPlugins() {
    return [createDocumentBookmarkPlugin(this.name)];
  },

  addAttributes() {
    return {
      id: {
        default: '',
        parseHTML: (element) => element.dataset.bookmarkId ?? '',
        renderHTML: () => ({}),
      },
      name: {
        default: '',
        parseHTML: (element) => element.dataset.bookmarkName ?? '',
        renderHTML: () => ({}),
      },
      nativeId: {
        default: null,
        parseHTML: (element) =>
          normalizeDocumentBookmarkNativeId(element.dataset.officeBookmarkId),
        renderHTML: () => ({}),
      },
      kind: {
        default: 'start',
        parseHTML: (element) =>
          documentBookmarkBoundaryKind(element.dataset.bookmarkKind),
        renderHTML: () => ({}),
      },
    };
  },

  parseHTML() {
    return [
      {
        tag: 'span[data-document-bookmark-boundary]',
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          return {
            id: node.dataset.bookmarkId ?? '',
            name: node.dataset.bookmarkName ?? '',
            nativeId: normalizeDocumentBookmarkNativeId(
              node.dataset.officeBookmarkId,
            ),
            kind: documentBookmarkBoundaryKind(node.dataset.bookmarkKind),
          };
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const name = normalizeDocumentBookmarkName(node.attrs.name) ?? '';
    const nativeId = normalizeDocumentBookmarkNativeId(node.attrs.nativeId);
    const kind = documentBookmarkBoundaryKind(node.attrs.kind);
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        ...(kind === 'start' && name ? { id: name } : {}),
        'data-document-bookmark-boundary': 'true',
        'data-bookmark-kind': kind,
        'data-bookmark-id':
          typeof node.attrs.id === 'string' ? node.attrs.id : '',
        'data-bookmark-name': name,
        'data-office-bookmark-id':
          nativeId === null ? undefined : String(nativeId),
        class: `work-document-bookmark-boundary ${kind}`,
        contenteditable: 'false',
        'aria-hidden': 'true',
      }),
    ];
  },

  renderText() {
    return '';
  },
});

export function validateDocumentBookmarkName(value: string): string | null {
  const name = value.trim();
  return BOOKMARK_UI_NAME_PATTERN.test(name) &&
    Array.from(name).length <= MAX_BOOKMARK_NAME_LENGTH
    ? null
    : DOCUMENT_BOOKMARK_VALIDATION_MESSAGE;
}

export function normalizeDocumentBookmarkName(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const name = value.trim();
  return BOOKMARK_NAME_PATTERN.test(name) &&
    Array.from(name).length <= MAX_BOOKMARK_NAME_LENGTH
    ? name
    : null;
}

export function normalizeDocumentBookmarkNativeId(
  value: unknown,
): number | null {
  if (
    value === null ||
    value === undefined ||
    value === '' ||
    (typeof value === 'string' && !value.trim())
  ) {
    return null;
  }
  const number = typeof value === 'number' ? value : Number(value);
  return Number.isSafeInteger(number) &&
    number >= 0 &&
    number <= MAX_BOOKMARK_NATIVE_ID
    ? number
    : null;
}

export function editorDocumentBookmarks(
  editor: Editor,
): WorkDocumentBookmark[] {
  return collectDocumentBookmarkPairs(
    editor.state.doc,
    'documentBookmarkBoundary',
  ).pairs.map(({ id, name, nativeId, from, to }) => ({
    id,
    name,
    nativeId,
    from,
    to,
  }));
}

export function activeDocumentBookmark(
  editor: Editor,
): WorkDocumentBookmark | null {
  const { from, to } = editor.state.selection;
  const matches = editorDocumentBookmarks(editor).filter(
    (bookmark) => from >= bookmark.from + 1 && to <= bookmark.to,
  );
  return (
    matches.sort(
      (left, right) => left.to - left.from - (right.to - right.from),
    )[0] ?? null
  );
}

export function documentBookmarkNameExists(
  editor: Editor,
  value: string,
  exceptId?: string,
): boolean {
  const name = value.trim().toLowerCase();
  return editorDocumentBookmarks(editor).some(
    (bookmark) =>
      bookmark.id !== exceptId && bookmark.name.toLowerCase() === name,
  );
}

export function normalizeDocumentBookmarksHtml(source: string): string {
  const document = new DOMParser().parseFromString(source, 'text/html');
  const boundaries = Array.from(
    document.body.querySelectorAll<HTMLElement>(
      'span[data-document-bookmark-boundary]',
    ),
  );
  const collection = collectDomBookmarkPairs(boundaries);
  for (const orphan of collection.orphans) orphan.remove();
  const registry = createBookmarkRegistry();
  for (const pair of collection.pairs) {
    const identity = uniqueDocumentBookmarkIdentity(pair, registry);
    applyBookmarkIdentityToElement(pair.start, identity, 'start');
    applyBookmarkIdentityToElement(pair.end, identity, 'end');
  }
  synchronizeDomInternalLinks(document.body, registry.names);
  return document.body.innerHTML;
}

function insertDocumentBookmarkCommand(
  { dispatch, editor, state, tr }: CommandProps,
  value: string,
): boolean {
  const name = value.trim();
  const boundaryType = editor.schema.nodes.documentBookmarkBoundary;
  const selection = state.selection;
  if (
    !boundaryType ||
    validateDocumentBookmarkName(name) ||
    documentBookmarkNameExists(editor, name) ||
    !selection.$from.parent.inlineContent ||
    !selection.$to.parent.inlineContent
  ) {
    return false;
  }
  if (!dispatch) return true;
  const registry = bookmarkRegistryForDocument(state.doc, boundaryType.name);
  const identity = uniqueDocumentBookmarkIdentity(
    {
      id: createUniqueBookmarkId(registry.ids),
      name,
      nativeId: nextBookmarkNativeId(registry.nativeIds),
    },
    registry,
  );
  const start = boundaryType.create({ ...identity, kind: 'start' });
  const end = boundaryType.create({ ...identity, kind: 'end' });
  const from = selection.from;
  const to = selection.to;
  tr.insert(to, end);
  tr.insert(from, start);
  tr.setSelection(TextSelection.create(tr.doc, from + 1, to + 1));
  dispatch(tr.scrollIntoView());
  return true;
}

function deleteDocumentBookmarkCommand(
  { dispatch, editor, state, tr }: CommandProps,
  id: string,
): boolean {
  const bookmark = collectDocumentBookmarkPairs(
    state.doc,
    editor.schema.nodes.documentBookmarkBoundary?.name ?? '',
  ).pairs.find((candidate) => candidate.id === id);
  if (!bookmark) return false;
  if (!dispatch) return true;
  tr.delete(bookmark.to, bookmark.to + bookmark.end.node.nodeSize);
  tr.delete(bookmark.from, bookmark.from + bookmark.start.node.nodeSize);
  tr.setSelection(TextSelection.near(tr.doc.resolve(bookmark.from)));
  dispatch(tr.scrollIntoView());
  return true;
}

function createDocumentBookmarkPlugin(boundaryNodeName: string): Plugin {
  return new Plugin({
    view(view) {
      const transaction = normalizeDocumentBookmarks(
        view.state,
        boundaryNodeName,
      );
      if (transaction) {
        transaction.setMeta('addToHistory', false);
        view.dispatch(transaction);
      }
      return {};
    },
    appendTransaction(transactions, oldState, newState) {
      if (!transactions.some((transaction) => transaction.docChanged)) {
        return null;
      }
      return normalizeDocumentBookmarks(
        newState,
        boundaryNodeName,
        oldState,
        transactions,
      );
    },
  });
}

function normalizeDocumentBookmarks(
  state: EditorState,
  boundaryNodeName: string,
  oldState?: EditorState,
  transactions: readonly Transaction[] = [],
): Transaction | null {
  const collection = collectDocumentBookmarkPairs(state.doc, boundaryNodeName);
  const retained = oldState
    ? retainedDocumentBookmarkPairs(
        oldState,
        state,
        transactions,
        boundaryNodeName,
      )
    : new Set<string>();
  const ordered = [
    ...collection.pairs.filter((pair) => retained.has(pairPositionKey(pair))),
    ...collection.pairs.filter((pair) => !retained.has(pairPositionKey(pair))),
  ];
  const registry = createBookmarkRegistry();
  const updates = new Map<number, WorkDocumentBookmarkIdentity>();
  const renames: BookmarkRename[] = [];
  const effectivePairs: BookmarkPairAtPosition[] = [];
  for (const pair of ordered) {
    const identity = uniqueDocumentBookmarkIdentity(pair, registry);
    if (!sameBookmarkIdentity(pair.start.node.attrs, identity)) {
      updates.set(pair.from, identity);
    }
    if (!sameBookmarkIdentity(pair.end.node.attrs, identity)) {
      updates.set(pair.to, identity);
    }
    if (!sameBookmarkIdentity(pair, identity)) {
      if (pair.name && pair.name !== identity.name) {
        renames.push({
          from: pair.from,
          to: pair.to,
          previousName: pair.name,
          nextName: identity.name,
        });
      }
    }
    effectivePairs.push({ ...pair, ...identity });
  }

  const tr = state.tr;
  for (const boundary of documentBookmarkBoundaries(
    state.doc,
    boundaryNodeName,
  )) {
    const identity = updates.get(boundary.position);
    if (!identity) continue;
    tr.setNodeMarkup(boundary.position, undefined, {
      ...boundary.node.attrs,
      ...identity,
    });
  }
  synchronizeInternalLinkMarks(state, tr, effectivePairs, renames);
  for (const orphan of [...collection.orphans].sort(
    (left, right) => right.position - left.position,
  )) {
    tr.delete(orphan.position, orphan.position + orphan.node.nodeSize);
  }
  return tr.docChanged ? tr : null;
}

function retainedDocumentBookmarkPairs(
  oldState: EditorState,
  newState: EditorState,
  transactions: readonly Transaction[],
  boundaryNodeName: string,
): Set<string> {
  const mapping = transactionMapping(transactions);
  const retained = new Set<string>();
  const current = collectDocumentBookmarkPairs(
    newState.doc,
    boundaryNodeName,
  ).pairs;
  const currentById = new Map<string, BookmarkPairAtPosition[]>();
  for (const pair of current) {
    const matches = currentById.get(pair.id) ?? [];
    matches.push(pair);
    currentById.set(pair.id, matches);
  }
  for (const previous of collectDocumentBookmarkPairs(
    oldState.doc,
    boundaryNodeName,
  ).pairs) {
    const mappedFrom = mapping.mapResult(previous.from, 1);
    const mappedTo = mapping.mapResult(previous.to, 1);
    const exact = current.find(
      (pair) =>
        pair.from === mappedFrom.pos &&
        pair.to === mappedTo.pos &&
        pair.id === previous.id,
    );
    if (exact) {
      retained.add(pairPositionKey(exact));
      continue;
    }
    const sameIdentity = currentById.get(previous.id);
    if (sameIdentity?.length === 1) {
      retained.add(pairPositionKey(sameIdentity[0]));
    }
  }
  return retained;
}

function synchronizeInternalLinkMarks(
  state: EditorState,
  tr: Transaction,
  bookmarks: readonly BookmarkPairAtPosition[],
  renames: readonly BookmarkRename[],
): void {
  const linkType = state.schema.marks.link;
  if (!linkType) return;
  const names = new Set(
    bookmarks.map((bookmark) => bookmark.name.toLowerCase()),
  );
  state.doc.descendants((node, position) => {
    if (!node.isText) return;
    const link = node.marks.find((mark) => mark.type === linkType);
    if (!link) return;
    const href = typeof link.attrs.href === 'string' ? link.attrs.href : '';
    if (!href.startsWith('#')) return;
    const target = href.slice(1);
    const rename = renames
      .filter(
        (candidate) =>
          position > candidate.from &&
          position < candidate.to &&
          target.toLowerCase() === candidate.previousName.toLowerCase(),
      )
      .sort((left, right) => left.to - left.from - (right.to - right.from))[0];
    const nextHref = rename ? `#${rename.nextName}` : href;
    const nextTarget = nextHref.slice(1).toLowerCase();
    const nextClass = toggleClassToken(
      link.attrs.class,
      MISSING_LINK_CLASS,
      !names.has(nextTarget),
    );
    if (nextHref === href && nextClass === normalizedClass(link.attrs.class)) {
      return;
    }
    replaceLinkMark(state, tr, node, position, link, {
      ...link.attrs,
      href: nextHref,
      class: nextClass || null,
    });
  });
}

function replaceLinkMark(
  _state: EditorState,
  tr: Transaction,
  node: ProseMirrorNode,
  position: number,
  link: ProseMirrorMark,
  attributes: Record<string, unknown>,
): void {
  tr.removeMark(position, position + node.nodeSize, link.type);
  tr.addMark(position, position + node.nodeSize, link.type.create(attributes));
}

function collectDocumentBookmarkPairs(
  document: ProseMirrorNode,
  boundaryNodeName: string,
): BookmarkPairCollection {
  const boundaries = documentBookmarkBoundaries(document, boundaryNodeName);
  const open = new Map<string, BookmarkBoundaryAtPosition[]>();
  const pairs: BookmarkPairAtPosition[] = [];
  const orphans: BookmarkBoundaryAtPosition[] = [];
  for (const boundary of boundaries) {
    const key = bookmarkBoundaryPairKey(boundary.node.attrs);
    if (boundary.kind === 'start') {
      const stack = open.get(key) ?? [];
      stack.push(boundary);
      open.set(key, stack);
      continue;
    }
    const stack = open.get(key);
    const start = stack?.pop();
    if (!start) {
      orphans.push(boundary);
      continue;
    }
    pairs.push(bookmarkPair(start, boundary));
  }
  for (const stack of open.values()) orphans.push(...stack);
  pairs.sort((left, right) => left.from - right.from || left.to - right.to);
  return { pairs, orphans };
}

function documentBookmarkBoundaries(
  document: ProseMirrorNode,
  boundaryNodeName: string,
): BookmarkBoundaryAtPosition[] {
  if (!boundaryNodeName) return [];
  const boundaries: BookmarkBoundaryAtPosition[] = [];
  document.descendants((node, position) => {
    if (node.type.name !== boundaryNodeName) return;
    boundaries.push({
      node,
      position,
      kind: documentBookmarkBoundaryKind(node.attrs.kind),
    });
  });
  return boundaries;
}

function bookmarkPair(
  start: BookmarkBoundaryAtPosition,
  end: BookmarkBoundaryAtPosition,
): BookmarkPairAtPosition {
  return {
    id: bookmarkInternalId(start.node.attrs.id),
    name: normalizeDocumentBookmarkName(start.node.attrs.name) ?? '',
    nativeId:
      normalizeDocumentBookmarkNativeId(start.node.attrs.nativeId) ?? -1,
    from: start.position,
    to: end.position,
    start,
    end,
  };
}

function bookmarkRegistryForDocument(
  document: ProseMirrorNode,
  boundaryNodeName: string,
): BookmarkRegistry {
  const registry = createBookmarkRegistry();
  for (const pair of collectDocumentBookmarkPairs(document, boundaryNodeName)
    .pairs) {
    const identity = bookmarkIdentity(pair);
    if (identity) reserveBookmarkIdentity(identity, registry);
  }
  return registry;
}

function uniqueDocumentBookmarkIdentity(
  source: Partial<WorkDocumentBookmarkIdentity>,
  registry: BookmarkRegistry,
): WorkDocumentBookmarkIdentity {
  const preferred = bookmarkIdentity(source);
  if (preferred && !bookmarkIdentityConflicts(preferred, registry)) {
    reserveBookmarkIdentity(preferred, registry);
    return preferred;
  }
  const baseName = normalizeDocumentBookmarkName(source.name) ?? 'Bookmark';
  const identity = {
    id: createUniqueBookmarkId(registry.ids),
    name: uniqueBookmarkName(baseName, registry.names),
    nativeId: nextBookmarkNativeId(registry.nativeIds),
  };
  reserveBookmarkIdentity(identity, registry);
  return identity;
}

function bookmarkIdentity(
  source: Partial<WorkDocumentBookmarkIdentity>,
): WorkDocumentBookmarkIdentity | null {
  const id = bookmarkInternalId(source.id);
  const name = normalizeDocumentBookmarkName(source.name);
  const nativeId = normalizeDocumentBookmarkNativeId(source.nativeId);
  return id && name && nativeId !== null ? { id, name, nativeId } : null;
}

function bookmarkIdentityConflicts(
  identity: WorkDocumentBookmarkIdentity,
  registry: BookmarkRegistry,
): boolean {
  return (
    registry.ids.has(identity.id) ||
    registry.names.has(identity.name.toLowerCase()) ||
    registry.nativeIds.has(identity.nativeId)
  );
}

function reserveBookmarkIdentity(
  identity: WorkDocumentBookmarkIdentity,
  registry: BookmarkRegistry,
): void {
  registry.ids.add(identity.id);
  registry.names.add(identity.name.toLowerCase());
  registry.nativeIds.add(identity.nativeId);
}

function createBookmarkRegistry(): BookmarkRegistry {
  return { ids: new Set(), names: new Set(), nativeIds: new Set() };
}

function createUniqueBookmarkId(ids: ReadonlySet<string>): string {
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const id = createWorkId('bookmark');
    if (!ids.has(id)) return id;
  }
  let suffix = 1;
  while (ids.has(`bookmark-${suffix}`)) suffix += 1;
  return `bookmark-${suffix}`;
}

function uniqueBookmarkName(base: string, names: ReadonlySet<string>): string {
  if (!names.has(base.toLowerCase())) return base;
  let suffix = 2;
  while (suffix <= MAX_BOOKMARK_NATIVE_ID) {
    const ending = `_${suffix}`;
    const prefix = Array.from(base)
      .slice(0, MAX_BOOKMARK_NAME_LENGTH - ending.length)
      .join('');
    const candidate = `${prefix}${ending}`;
    if (!names.has(candidate.toLowerCase())) return candidate;
    suffix += 1;
  }
  throw new Error('No unique Word bookmark name is available.');
}

function nextBookmarkNativeId(ids: ReadonlySet<number>): number {
  for (let id = 0; id <= MAX_BOOKMARK_NATIVE_ID; id += 1) {
    if (!ids.has(id)) return id;
  }
  throw new Error('No unique Word bookmark identifier is available.');
}

function sameBookmarkIdentity(
  source: Partial<WorkDocumentBookmarkIdentity>,
  identity: WorkDocumentBookmarkIdentity,
): boolean {
  return (
    bookmarkInternalId(source.id) === identity.id &&
    normalizeDocumentBookmarkName(source.name) === identity.name &&
    normalizeDocumentBookmarkNativeId(source.nativeId) === identity.nativeId
  );
}

function bookmarkInternalId(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}

function bookmarkBoundaryPairKey(attributes: Record<string, unknown>): string {
  const id = bookmarkInternalId(attributes.id);
  if (id) return `id:${id}`;
  return `legacy:${String(attributes.name ?? '')}:${String(
    attributes.nativeId ?? '',
  )}`;
}

function documentBookmarkBoundaryKind(
  value: unknown,
): WorkDocumentBookmarkBoundaryKind {
  return value === 'end' ? 'end' : 'start';
}

function transactionMapping(transactions: readonly Transaction[]): Mapping {
  const mapping = new Mapping();
  for (const transaction of transactions) {
    mapping.appendMapping(transaction.mapping);
  }
  return mapping;
}

function pairPositionKey(pair: Pick<WorkDocumentBookmark, 'from' | 'to'>) {
  return `${pair.from}:${pair.to}`;
}

interface DomBookmarkPair extends Partial<WorkDocumentBookmarkIdentity> {
  start: HTMLElement;
  end: HTMLElement;
}

function collectDomBookmarkPairs(boundaries: HTMLElement[]): {
  pairs: DomBookmarkPair[];
  orphans: HTMLElement[];
} {
  const open = new Map<string, HTMLElement[]>();
  const pairs: DomBookmarkPair[] = [];
  const orphans: HTMLElement[] = [];
  for (const boundary of boundaries) {
    const key = domBookmarkPairKey(boundary);
    if (
      documentBookmarkBoundaryKind(boundary.dataset.bookmarkKind) === 'start'
    ) {
      const stack = open.get(key) ?? [];
      stack.push(boundary);
      open.set(key, stack);
      continue;
    }
    const start = open.get(key)?.pop();
    if (!start) {
      orphans.push(boundary);
      continue;
    }
    pairs.push({
      id: start.dataset.bookmarkId,
      name: start.dataset.bookmarkName,
      nativeId:
        normalizeDocumentBookmarkNativeId(start.dataset.officeBookmarkId) ??
        undefined,
      start,
      end: boundary,
    });
  }
  for (const stack of open.values()) orphans.push(...stack);
  return { pairs, orphans };
}

function domBookmarkPairKey(element: HTMLElement): string {
  const id = element.dataset.bookmarkId?.trim();
  return id
    ? `id:${id}`
    : `legacy:${element.dataset.bookmarkName ?? ''}:${
        element.dataset.officeBookmarkId ?? ''
      }`;
}

function applyBookmarkIdentityToElement(
  element: HTMLElement,
  identity: WorkDocumentBookmarkIdentity,
  kind: WorkDocumentBookmarkBoundaryKind,
): void {
  element.dataset.documentBookmarkBoundary = 'true';
  element.dataset.bookmarkKind = kind;
  element.dataset.bookmarkId = identity.id;
  element.dataset.bookmarkName = identity.name;
  element.dataset.officeBookmarkId = String(identity.nativeId);
  element.classList.add('work-document-bookmark-boundary', kind);
  element.contentEditable = 'false';
  element.setAttribute('aria-hidden', 'true');
  if (kind === 'start') element.id = identity.name;
  else element.removeAttribute('id');
}

function synchronizeDomInternalLinks(
  root: HTMLElement,
  names: ReadonlySet<string>,
): void {
  for (const link of root.querySelectorAll<HTMLAnchorElement>('a[href^="#"]')) {
    const target = (link.getAttribute('href') ?? '').slice(1).toLowerCase();
    link.classList.toggle(MISSING_LINK_CLASS, !names.has(target));
    if (!link.className) link.removeAttribute('class');
  }
}

function normalizedClass(value: unknown): string {
  return typeof value === 'string'
    ? value.trim().split(/\s+/).filter(Boolean).join(' ')
    : '';
}

function toggleClassToken(
  value: unknown,
  token: string,
  enabled: boolean,
): string {
  const tokens = new Set(normalizedClass(value).split(' ').filter(Boolean));
  if (enabled) tokens.add(token);
  else tokens.delete(token);
  return Array.from(tokens).join(' ');
}
