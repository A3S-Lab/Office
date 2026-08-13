import type { Editor } from '@tiptap/core';
import type {
  Mark as ProseMirrorMark,
  Node as ProseMirrorNode,
} from '@tiptap/pm/model';

type DocumentFormatMarkName = (typeof documentFormatMarkNames)[number];
type DocumentFormatBlockName = 'heading' | 'paragraph';

interface DocumentFormatMarkSnapshot {
  attrs: Record<string, unknown>;
  type: DocumentFormatMarkName;
}

interface DocumentFormatBlockSnapshot {
  attrs: Record<string, unknown>;
  type: DocumentFormatBlockName;
}

interface DocumentFormatSnapshot {
  block: DocumentFormatBlockSnapshot;
  marks: DocumentFormatMarkSnapshot[];
}

let documentFormatClipboard: DocumentFormatSnapshot | null = null;
const documentFormatClipboardListeners = new Set<() => void>();

export function copyDocumentFormatting(editor: Editor): boolean {
  if (editor.isDestroyed) return false;
  const block = documentFormatBlockAtSelection(editor);
  if (!block) return false;
  const marks = documentFormatMarksAtSelection(editor)
    .filter((mark) => isDocumentFormatMarkName(mark.type.name))
    .map((mark) => ({
      type: mark.type.name as DocumentFormatMarkName,
      attrs: cloneFormattingRecord(mark.attrs),
    }));
  documentFormatClipboard = { block, marks };
  notifyDocumentFormatClipboardListeners();
  return true;
}

export function pasteDocumentFormatting(editor: Editor): boolean {
  if (editor.isDestroyed || !documentFormatClipboard) return false;
  const { state } = editor;
  const { from, to, empty } = state.selection;
  const transaction = state.tr;
  const blockPositions = documentFormatBlockPositions(
    state.doc,
    from,
    to,
    empty,
  );
  const blockType = state.schema.nodes[documentFormatClipboard.block.type];
  if (!blockPositions.length) return false;

  for (const position of blockPositions) {
    const node = transaction.doc.nodeAt(position);
    if (!node || !isDocumentFormatBlockName(node.type.name)) continue;
    const resolved = transaction.doc.resolve(position);
    const index = resolved.index();
    const compatibleBlockType =
      blockType && resolved.parent.canReplaceWith(index, index + 1, blockType)
        ? blockType
        : node.type;
    transaction.setNodeMarkup(
      position,
      compatibleBlockType,
      {
        ...node.attrs,
        ...cloneFormattingRecord(documentFormatClipboard.block.attrs),
      },
      node.marks,
    );
  }

  const clipboardMarks = documentFormatClipboard.marks.flatMap((snapshot) => {
    const type = state.schema.marks[snapshot.type];
    return type ? [type.create(cloneFormattingRecord(snapshot.attrs))] : [];
  });
  if (empty) {
    editor.view.dispatch(transaction.scrollIntoView());
    let commandChain = editor.chain().focus();
    for (const markName of documentFormatMarkNames) {
      if (state.schema.marks[markName]) {
        commandChain = commandChain.unsetMark(markName);
      }
    }
    for (const mark of documentFormatClipboard.marks) {
      if (state.schema.marks[mark.type]) {
        commandChain = commandChain.setMark(
          mark.type,
          cloneFormattingRecord(mark.attrs),
        );
      }
    }
    return commandChain.run();
  }

  for (const markName of documentFormatMarkNames) {
    const markType = state.schema.marks[markName];
    if (markType) transaction.removeMark(from, to, markType);
  }
  for (const mark of clipboardMarks) transaction.addMark(from, to, mark);

  editor.view.focus();
  editor.view.dispatch(transaction.scrollIntoView());
  return true;
}

export function hasDocumentFormatClipboard(): boolean {
  return documentFormatClipboard !== null;
}

export function subscribeDocumentFormatClipboard(
  listener: () => void,
): () => void {
  documentFormatClipboardListeners.add(listener);
  return () => documentFormatClipboardListeners.delete(listener);
}

export function clearDocumentFormatClipboard(): void {
  if (!documentFormatClipboard) return;
  documentFormatClipboard = null;
  notifyDocumentFormatClipboardListeners();
}

const documentFormatMarkNames = [
  'bold',
  'italic',
  'underline',
  'strike',
  'subscript',
  'superscript',
  'textStyle',
  'highlight',
] as const;

const documentFormatBlockAttributeNames = [
  'textAlign',
  'lineHeight',
  'lineRule',
  'autoLineHeight',
  'spaceBefore',
  'spaceAfter',
  'indentLevel',
  'rightIndent',
  'firstLineIndent',
  'keepLines',
  'keepWithNext',
  'pageBreakBefore',
  'widowControl',
  'contextualSpacing',
  'outlineLevel',
  'paragraphDirection',
  'paragraphShading',
  'tabStops',
] as const;

function documentFormatBlockAtSelection(
  editor: Editor,
): DocumentFormatBlockSnapshot | null {
  const { $from } = editor.state.selection;
  for (let depth = $from.depth; depth > 0; depth -= 1) {
    const node = $from.node(depth);
    if (!isDocumentFormatBlockName(node.type.name)) continue;
    const attrs: Record<string, unknown> = {};
    for (const name of documentFormatBlockAttributeNames) {
      if (Object.hasOwn(node.attrs, name)) {
        attrs[name] = cloneFormattingValue(node.attrs[name]);
      }
    }
    if (node.type.name === 'heading') attrs.level = node.attrs.level;
    return { type: node.type.name, attrs };
  }
  return null;
}

function documentFormatMarksAtSelection(
  editor: Editor,
): readonly ProseMirrorMark[] {
  const { doc, selection, storedMarks } = editor.state;
  if (selection.empty) return storedMarks ?? selection.$from.marks();
  let marks: readonly ProseMirrorMark[] | null = null;
  doc.nodesBetween(selection.from, selection.to, (node) => {
    if (!marks && node.isText) marks = node.marks;
    return marks === null;
  });
  return marks ?? selection.$from.marks();
}

function documentFormatBlockPositions(
  doc: ProseMirrorNode,
  from: number,
  to: number,
  empty: boolean,
): number[] {
  if (empty) {
    const resolved = doc.resolve(from);
    for (let depth = resolved.depth; depth > 0; depth -= 1) {
      if (isDocumentFormatBlockName(resolved.node(depth).type.name)) {
        return [resolved.before(depth)];
      }
    }
    return [];
  }
  const positions: number[] = [];
  doc.nodesBetween(from, to, (node, position) => {
    if (isDocumentFormatBlockName(node.type.name)) positions.push(position);
  });
  return positions;
}

function isDocumentFormatMarkName(
  value: string,
): value is DocumentFormatMarkName {
  return (documentFormatMarkNames as readonly string[]).includes(value);
}

function isDocumentFormatBlockName(
  value: string,
): value is DocumentFormatBlockName {
  return value === 'heading' || value === 'paragraph';
}

function cloneFormattingRecord(
  value: Record<string, unknown>,
): Record<string, unknown> {
  return Object.fromEntries(
    Object.entries(value).map(([key, entry]) => [
      key,
      cloneFormattingValue(entry),
    ]),
  );
}

function cloneFormattingValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(cloneFormattingValue);
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value).map(([key, entry]) => [
        key,
        cloneFormattingValue(entry),
      ]),
    );
  }
  return value;
}

function notifyDocumentFormatClipboardListeners(): void {
  for (const listener of documentFormatClipboardListeners) listener();
}
