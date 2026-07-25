import type { Editor } from '@tiptap/core';
import type { Transaction } from '@tiptap/pm/state';
import type { WorkDocumentContent, WorkDocumentNode } from './work-types';

const SURROUNDING_TEXT_LIMIT = 2_000;

export type WorkDocumentSelectionMenuIcon =
  | 'copy'
  | 'language'
  | 'message'
  | 'quote'
  | 'sparkles'
  | 'wand';

export interface WorkDocumentSelectionSnapshot {
  selection: {
    from: number;
    to: number;
    text: string;
    rawText: string;
    beforeText: string;
    afterText: string;
    model: WorkDocumentNode;
  };
  document: {
    content: WorkDocumentContent;
    html: string;
    text: string;
  };
}

export type WorkDocumentSelectionCommandFailure =
  | 'command-rejected'
  | 'editor-unavailable'
  | 'invalid-text'
  | 'read-only'
  | 'stale-selection';

export type WorkDocumentSelectionCommandResult =
  | { applied: true }
  | {
      applied: false;
      reason: WorkDocumentSelectionCommandFailure;
    };

export interface WorkDocumentSelectionCommands {
  copyText(): Promise<boolean>;
  insertAfter(text: string): WorkDocumentSelectionCommandResult;
  insertBefore(text: string): WorkDocumentSelectionCommandResult;
  replaceText(text: string): WorkDocumentSelectionCommandResult;
}

export interface WorkDocumentSelectionContext
  extends WorkDocumentSelectionSnapshot {
  commands: WorkDocumentSelectionCommands;
}

export interface WorkDocumentSelectionMenuItem {
  id: string;
  label: string;
  icon?: WorkDocumentSelectionMenuIcon;
  shortcut?: string;
  ariaKeyShortcut?: string;
  danger?: boolean;
  disabled?: boolean;
  separatorBefore?: boolean;
  onSelect(context: WorkDocumentSelectionContext): void | Promise<void>;
}

export type WorkGetDocumentSelectionMenuItems = (
  context: WorkDocumentSelectionSnapshot,
) => readonly WorkDocumentSelectionMenuItem[];

export interface WorkDocumentSelectionAction {
  context: WorkDocumentSelectionContext;
  dispose(): void;
}

export function createWorkDocumentSelectionSnapshot(
  editor: Editor,
  content: WorkDocumentContent,
): WorkDocumentSelectionSnapshot | null {
  const { from, to, empty } = editor.state.selection;
  if (empty) return null;

  const rawText = editor.state.doc.textBetween(from, to, '\n');
  const text = rawText.trim();
  if (!text) return null;

  const beforeText = editor.state.doc
    .textBetween(0, from, '\n')
    .slice(-SURROUNDING_TEXT_LIMIT);
  const afterText = editor.state.doc
    .textBetween(to, editor.state.doc.content.size, '\n')
    .slice(0, SURROUNDING_TEXT_LIMIT);
  const selectedContent = editor.state.selection.content().content.toJSON();

  return {
    selection: {
      from,
      to,
      text,
      rawText,
      beforeText,
      afterText,
      model: {
        type: 'doc',
        ...(selectedContent?.length
          ? { content: selectedContent as WorkDocumentNode[] }
          : {}),
      },
    },
    document: {
      content,
      html: editor.getHTML(),
      text: editor.state.doc.textBetween(
        0,
        editor.state.doc.content.size,
        '\n',
      ),
    },
  };
}

export function createWorkDocumentSelectionAction(
  editor: Editor,
  snapshot: WorkDocumentSelectionSnapshot,
  getTrackChanges: () => boolean,
): WorkDocumentSelectionAction {
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
    | WorkDocumentSelectionCommandFailure => {
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
  ): WorkDocumentSelectionCommandResult => {
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
    const applied = getTrackChanges()
      ? editor.commands.replaceDocumentTextWithTrackedChange(
          editRange.from,
          editRange.to,
          text,
        )
      : editor
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
        copyText: async () => {
          try {
            if (
              typeof navigator === 'undefined' ||
              !navigator.clipboard?.writeText
            ) {
              return false;
            }
            await navigator.clipboard.writeText(snapshot.selection.rawText);
            return true;
          } catch {
            return false;
          }
        },
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

export function documentPlainTextAsHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
    .replace(/\r?\n/g, '<br>');
}
