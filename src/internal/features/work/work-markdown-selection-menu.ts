import type { Editor } from '@tiptap/core';
import type { Transaction } from '@tiptap/pm/state';
import type {
  MarkdownSourceEdit,
  MarkdownSourceSelection,
} from './editors/markdown-source-commands';
import { replaceMarkdownSourceSelection } from './editors/markdown-source-commands';
import { documentPlainTextAsHtml } from './work-document-selection-menu';
import type { WorkDocumentSelectionMenuIcon } from './work-document-selection-menu';
import type { WorkMarkdownContent } from './work-types';

const SURROUNDING_TEXT_LIMIT = 2_000;

export type WorkMarkdownSelectionMenuIcon = WorkDocumentSelectionMenuIcon;

export type WorkMarkdownSelectionCommandFailure =
  | 'command-rejected'
  | 'editor-unavailable'
  | 'invalid-text'
  | 'read-only'
  | 'stale-selection';

export type WorkMarkdownSelectionCommandResult =
  | { applied: true }
  | {
      applied: false;
      reason: WorkMarkdownSelectionCommandFailure;
    };

export interface WorkMarkdownSelectionCommands {
  copyText(): Promise<boolean>;
  insertAfter(text: string): WorkMarkdownSelectionCommandResult;
  insertBefore(text: string): WorkMarkdownSelectionCommandResult;
  replaceText(text: string): WorkMarkdownSelectionCommandResult;
}

export interface WorkMarkdownSelectionSnapshot {
  selection: {
    surface: 'source' | 'visual';
    from: number;
    to: number;
    text: string;
    rawText: string;
    beforeText: string;
    afterText: string;
  };
  document: {
    content: WorkMarkdownContent;
    markdown: string;
    text: string;
  };
}

export interface WorkMarkdownSelectionContext
  extends WorkMarkdownSelectionSnapshot {
  commands: WorkMarkdownSelectionCommands;
}

export interface WorkMarkdownSelectionMenuItem {
  id: string;
  label: string;
  icon?: WorkMarkdownSelectionMenuIcon;
  shortcut?: string;
  ariaKeyShortcut?: string;
  danger?: boolean;
  disabled?: boolean;
  separatorBefore?: boolean;
  onSelect(context: WorkMarkdownSelectionContext): void | Promise<void>;
}

export type WorkGetMarkdownSelectionMenuItems = (
  context: WorkMarkdownSelectionSnapshot,
) => readonly WorkMarkdownSelectionMenuItem[];

export interface WorkMarkdownSelectionAction {
  context: WorkMarkdownSelectionContext;
  dispose(): void;
}

export function createWorkMarkdownSourceSelectionSnapshot(
  markdown: string,
  selection: MarkdownSourceSelection,
  content: WorkMarkdownContent,
): WorkMarkdownSelectionSnapshot | null {
  const from = Math.max(0, Math.min(markdown.length, selection.start));
  const to = Math.max(from, Math.min(markdown.length, selection.end));
  const rawText = markdown.slice(from, to);
  const text = rawText.trim();
  if (!text) return null;
  return {
    selection: {
      surface: 'source',
      from,
      to,
      text,
      rawText,
      beforeText: markdown.slice(0, from).slice(-SURROUNDING_TEXT_LIMIT),
      afterText: markdown.slice(to).slice(0, SURROUNDING_TEXT_LIMIT),
    },
    document: {
      content:
        content.markdown === markdown ? content : { ...content, markdown },
      markdown,
      text: markdown,
    },
  };
}

export function createWorkMarkdownVisualSelectionSnapshot(
  editor: Editor,
  content: WorkMarkdownContent,
): WorkMarkdownSelectionSnapshot | null {
  const { from, to, empty } = editor.state.selection;
  if (empty) return null;
  const rawText = editor.state.doc.textBetween(from, to, '\n');
  const text = rawText.trim();
  if (!text) return null;
  const markdown = editor.getMarkdown();
  return {
    selection: {
      surface: 'visual',
      from,
      to,
      text,
      rawText,
      beforeText: editor.state.doc
        .textBetween(0, from, '\n')
        .slice(-SURROUNDING_TEXT_LIMIT),
      afterText: editor.state.doc
        .textBetween(to, editor.state.doc.content.size, '\n')
        .slice(0, SURROUNDING_TEXT_LIMIT),
    },
    document: {
      content:
        content.markdown === markdown ? content : { ...content, markdown },
      markdown,
      text: editor.state.doc.textBetween(
        0,
        editor.state.doc.content.size,
        '\n',
      ),
    },
  };
}

export function createWorkMarkdownSourceSelectionAction(
  snapshot: WorkMarkdownSelectionSnapshot,
  getMarkdown: () => string,
  applyEdit: (edit: MarkdownSourceEdit) => boolean,
): WorkMarkdownSelectionAction {
  let from = snapshot.selection.from;
  let to = snapshot.selection.to;
  let expectedText = snapshot.selection.rawText;
  let disposed = false;
  const currentSelection = ():
    | MarkdownSourceSelection
    | WorkMarkdownSelectionCommandFailure => {
    if (disposed) return 'editor-unavailable';
    const markdown = getMarkdown();
    if (markdown.slice(from, to) !== expectedText) return 'stale-selection';
    return { start: from, end: to, direction: 'none' };
  };
  const editText = (
    text: string,
    placement: 'after' | 'before' | 'replace',
  ): WorkMarkdownSelectionCommandResult => {
    if (typeof text !== 'string' || (placement !== 'replace' && !text)) {
      return { applied: false, reason: 'invalid-text' };
    }
    const selection = currentSelection();
    if (typeof selection === 'string') {
      return { applied: false, reason: selection };
    }
    const editSelection =
      placement === 'before'
        ? { ...selection, end: selection.start }
        : placement === 'after'
          ? { ...selection, start: selection.end }
          : selection;
    const edit = replaceMarkdownSourceSelection(
      getMarkdown(),
      editSelection,
      text,
    );
    if (!applyEdit(edit)) {
      return { applied: false, reason: 'command-rejected' };
    }
    if (placement === 'before') {
      from += text.length;
      to += text.length;
    } else if (placement === 'replace') {
      to = from + text.length;
      expectedText = text;
    }
    return { applied: true };
  };
  return {
    context: {
      ...snapshot,
      commands: {
        copyText: () => copyMarkdownSelection(snapshot.selection.rawText),
        insertAfter: (text) => editText(text, 'after'),
        insertBefore: (text) => editText(text, 'before'),
        replaceText: (text) => editText(text, 'replace'),
      },
    },
    dispose: () => {
      disposed = true;
    },
  };
}

export function createWorkMarkdownVisualSelectionAction(
  editor: Editor,
  snapshot: WorkMarkdownSelectionSnapshot,
): WorkMarkdownSelectionAction {
  let from = snapshot.selection.from;
  let to = snapshot.selection.to;
  let stale = false;
  let disposed = false;
  const handleTransaction = ({ transaction }: { transaction: Transaction }) => {
    if (!transaction.docChanged || stale || disposed) return;
    const mappedFrom = transaction.mapping.mapResult(from, 1);
    const mappedTo = transaction.mapping.mapResult(to, -1);
    from = mappedFrom.pos;
    to = mappedTo.pos;
    stale =
      mappedFrom.deletedAcross ||
      mappedTo.deletedAcross ||
      from > to ||
      transaction.doc.textBetween(from, to, '\n') !==
        snapshot.selection.rawText;
  };
  editor.on('transaction', handleTransaction);
  const currentRange = ():
    | { from: number; to: number }
    | WorkMarkdownSelectionCommandFailure => {
    if (disposed || editor.isDestroyed) return 'editor-unavailable';
    if (!editor.isEditable) return 'read-only';
    if (
      stale ||
      from < 0 ||
      to < from ||
      to > editor.state.doc.content.size ||
      editor.state.doc.textBetween(from, to, '\n') !==
        snapshot.selection.rawText
    ) {
      return 'stale-selection';
    }
    return { from, to };
  };
  const editText = (
    text: string,
    placement: 'after' | 'before' | 'replace',
  ): WorkMarkdownSelectionCommandResult => {
    if (typeof text !== 'string' || (placement !== 'replace' && !text)) {
      return { applied: false, reason: 'invalid-text' };
    }
    const range = currentRange();
    if (typeof range === 'string') {
      return { applied: false, reason: range };
    }
    const editRange =
      placement === 'before'
        ? { from: range.from, to: range.from }
        : placement === 'after'
          ? { from: range.to, to: range.to }
          : range;
    const applied = editor
      .chain()
      .focus()
      .setTextSelection(editRange)
      .insertContent(documentPlainTextAsHtml(text))
      .run();
    return applied
      ? { applied: true }
      : { applied: false, reason: 'command-rejected' };
  };
  return {
    context: {
      ...snapshot,
      commands: {
        copyText: () => copyMarkdownSelection(snapshot.selection.rawText),
        insertAfter: (text) => editText(text, 'after'),
        insertBefore: (text) => editText(text, 'before'),
        replaceText: (text) => editText(text, 'replace'),
      },
    },
    dispose: () => {
      if (disposed) return;
      disposed = true;
      editor.off('transaction', handleTransaction);
    },
  };
}

async function copyMarkdownSelection(text: string): Promise<boolean> {
  try {
    if (typeof navigator === 'undefined' || !navigator.clipboard?.writeText) {
      return false;
    }
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    return false;
  }
}
