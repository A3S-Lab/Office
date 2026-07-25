import { type CommandProps, Extension, type Editor } from '@tiptap/core';
import { BulletList, OrderedList } from '@tiptap/extension-list';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';

export type DocumentBulletListStyle = 'disc' | 'circle' | 'square';

export type DocumentOrderedListStyle =
  | 'decimal'
  | 'lower-alpha'
  | 'upper-alpha'
  | 'lower-roman'
  | 'upper-roman';

export interface DocumentOrderedListState {
  start: number;
  style: DocumentOrderedListStyle;
}

export const MAX_DOCUMENT_NUMBERING_START = 2_147_483_647;

type OrderedListType = 'a' | 'A' | 'i' | 'I' | null;

interface ActiveList {
  depth: number;
  node: ProseMirrorNode;
  position: number;
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentLists: {
      applyDocumentBulletList: (style: DocumentBulletListStyle) => ReturnType;
      applyDocumentOrderedList: (style: DocumentOrderedListStyle) => ReturnType;
      clearDocumentList: () => ReturnType;
      continueDocumentNumbering: () => ReturnType;
      restartDocumentNumbering: () => ReturnType;
      setDocumentNumberingStart: (start: number) => ReturnType;
    };
  }
}

export const DocumentBulletList = BulletList.extend({
  addAttributes() {
    return {
      bulletStyle: {
        default: 'disc',
        parseHTML: (element: HTMLElement) =>
          parsedBulletStyle(element) ?? 'disc',
        renderHTML: (attributes: Record<string, unknown>) => {
          const style = normalizeBulletStyle(attributes.bulletStyle) ?? 'disc';
          return style === 'disc'
            ? {}
            : {
                'data-office-bullet-style': style,
                style: `list-style-type: ${style}`,
              };
        },
      },
    };
  },
});

export const DocumentOrderedList = OrderedList;

export const DocumentListCommands = Extension.create({
  name: 'documentListCommands',

  addCommands() {
    return {
      applyDocumentBulletList: (style) => (props) =>
        applyDocumentBulletList(props, style),
      applyDocumentOrderedList: (style) => (props) =>
        applyDocumentOrderedList(props, style),
      clearDocumentList: () => (props) => clearDocumentList(props),
      continueDocumentNumbering: () => (props) =>
        continueDocumentNumbering(props),
      restartDocumentNumbering: () => (props) =>
        setDocumentNumberingStart(props, 1),
      setDocumentNumberingStart: (start) => (props) =>
        setDocumentNumberingStart(props, start),
    };
  },
});

export function documentBulletListStyle(
  editor: Editor,
): DocumentBulletListStyle | null {
  const active = activeList(editor.state, 'bulletList');
  if (!active) return null;
  return normalizeBulletStyle(active.node.attrs.bulletStyle) ?? 'disc';
}

export function documentOrderedListState(
  editor: Editor,
): DocumentOrderedListState | null {
  const active = activeList(editor.state, 'orderedList');
  if (!active) return null;
  return {
    start: normalizedNumberingStart(active.node.attrs.start),
    style: orderedListStyle(active.node.attrs.type),
  };
}

function applyDocumentBulletList(
  props: CommandProps,
  style: DocumentBulletListStyle,
): boolean {
  const normalized = normalizeBulletStyle(style);
  if (!normalized) return false;
  const current = activeList(props.state, 'bulletList');
  if (current) {
    return updateListAttributes(props, current, { bulletStyle: normalized });
  }
  return props
    .chain()
    .toggleBulletList()
    .updateAttributes('bulletList', { bulletStyle: normalized })
    .run();
}

function applyDocumentOrderedList(
  props: CommandProps,
  style: DocumentOrderedListStyle,
): boolean {
  const type = orderedListType(style);
  if (type === undefined) return false;
  const current = activeList(props.state, 'orderedList');
  if (current) return updateListAttributes(props, current, { type });
  return props
    .chain()
    .toggleOrderedList()
    .updateAttributes('orderedList', { start: 1, type })
    .run();
}

function clearDocumentList(props: CommandProps): boolean {
  if (activeList(props.state, 'bulletList')) {
    return props.commands.toggleBulletList();
  }
  if (activeList(props.state, 'orderedList')) {
    return props.commands.toggleOrderedList();
  }
  return false;
}

function setDocumentNumberingStart(
  props: CommandProps,
  start: number,
): boolean {
  if (!validNumberingStart(start)) return false;
  const current = activeList(props.state, 'orderedList');
  if (!current) return false;
  return updateListAttributes(props, current, { start });
}

function continueDocumentNumbering(props: CommandProps): boolean {
  const current = activeList(props.state, 'orderedList');
  if (!current) return false;
  const previous = precedingOrderedList(props.state.doc, current);
  if (!previous) return false;
  const start = Math.min(
    MAX_DOCUMENT_NUMBERING_START,
    normalizedNumberingStart(previous.node.attrs.start) +
      previous.node.childCount,
  );
  return updateListAttributes(props, current, {
    start,
    type: normalizedOrderedListType(previous.node.attrs.type),
  });
}

function updateListAttributes(
  { dispatch, state }: CommandProps,
  list: ActiveList,
  attributes: Record<string, unknown>,
): boolean {
  const next = { ...list.node.attrs, ...attributes };
  if (
    Object.entries(attributes).every(
      ([key, value]) => list.node.attrs[key] === value,
    )
  ) {
    return true;
  }
  dispatch?.(
    state.tr
      .setNodeMarkup(list.position, undefined, next, list.node.marks)
      .scrollIntoView(),
  );
  return true;
}

function activeList(
  state: CommandProps['state'],
  nodeName: 'bulletList' | 'orderedList',
): ActiveList | null {
  const position = state.selection.$from;
  for (let depth = position.depth; depth > 0; depth -= 1) {
    const node = position.node(depth);
    if (node.type.name !== nodeName) continue;
    return { depth, node, position: position.before(depth) };
  }
  return null;
}

function precedingOrderedList(
  document: ProseMirrorNode,
  current: ActiveList,
): ActiveList | null {
  let previous: ActiveList | null = null;
  document.descendants((node, position) => {
    if (position >= current.position) return false;
    if (node.type.name !== 'orderedList') return true;
    const parentDepth = document.resolve(position).depth;
    if (parentDepth === current.depth - 1) {
      previous = { depth: parentDepth + 1, node, position };
    }
    return true;
  });
  return previous;
}

function parsedBulletStyle(
  element: HTMLElement,
): DocumentBulletListStyle | null {
  const direct =
    element.dataset.officeBulletStyle || element.style.listStyleType;
  const fromList = normalizeBulletStyle(direct);
  if (fromList) return fromList;
  return normalizeBulletStyle(
    element.querySelector<HTMLElement>(':scope > li')?.style.listStyleType,
  );
}

function normalizeBulletStyle(value: unknown): DocumentBulletListStyle | null {
  return value === 'disc' || value === 'circle' || value === 'square'
    ? value
    : null;
}

function orderedListType(
  style: DocumentOrderedListStyle,
): OrderedListType | undefined {
  if (style === 'decimal') return null;
  if (style === 'lower-alpha') return 'a';
  if (style === 'upper-alpha') return 'A';
  if (style === 'lower-roman') return 'i';
  if (style === 'upper-roman') return 'I';
  return undefined;
}

function orderedListStyle(value: unknown): DocumentOrderedListStyle {
  const type = normalizedOrderedListType(value);
  if (type === 'a') return 'lower-alpha';
  if (type === 'A') return 'upper-alpha';
  if (type === 'i') return 'lower-roman';
  if (type === 'I') return 'upper-roman';
  return 'decimal';
}

function normalizedOrderedListType(value: unknown): OrderedListType {
  return value === 'a' || value === 'A' || value === 'i' || value === 'I'
    ? value
    : null;
}

function normalizedNumberingStart(value: unknown): number {
  const start = Number(value);
  return validNumberingStart(start) ? start : 1;
}

function validNumberingStart(value: number): boolean {
  return (
    Number.isSafeInteger(value) &&
    value >= 1 &&
    value <= MAX_DOCUMENT_NUMBERING_START
  );
}
