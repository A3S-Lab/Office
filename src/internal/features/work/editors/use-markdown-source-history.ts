import { useCallback, useRef, useState } from 'react';
import type {
  MarkdownSourceEdit,
  MarkdownSourceSelection,
} from './markdown-source-commands';

const MARKDOWN_SOURCE_HISTORY_LIMIT = 100;
const MARKDOWN_TYPING_COALESCE_MS = 1_000;

interface MarkdownSourceHistorySnapshot extends MarkdownSourceEdit {}

interface MarkdownSourceTypingGroup {
  at: number;
  inputType?: string;
}

interface MarkdownSourceHistoryState {
  past: MarkdownSourceHistorySnapshot[];
  present: MarkdownSourceHistorySnapshot;
  future: MarkdownSourceHistorySnapshot[];
  typing: MarkdownSourceTypingGroup | null;
}

export function useMarkdownSourceHistory(initialMarkdown: string) {
  const historyRef = useRef<MarkdownSourceHistoryState>({
    past: [],
    present: {
      markdown: initialMarkdown,
      selection: collapsedSelection(0),
    },
    future: [],
    typing: null,
  });
  const [, setVersion] = useState(0);
  const refresh = useCallback(() => setVersion((value) => value + 1), []);

  const record = useCallback(
    (
      edit: MarkdownSourceEdit,
      options: { typing?: boolean; inputType?: string } = {},
    ): boolean => {
      const history = historyRef.current;
      const next = normalizeSnapshot(edit);
      if (next.markdown === history.present.markdown) {
        history.present = next;
        return false;
      }

      const now = Date.now();
      const coalesce =
        options.typing === true &&
        canCoalesceTyping(history, next, options.inputType, now);
      if (!coalesce) {
        history.past = [
          ...history.past.slice(-(MARKDOWN_SOURCE_HISTORY_LIMIT - 1)),
          history.present,
        ];
      }
      history.present = next;
      history.future = [];
      history.typing = options.typing
        ? { at: now, inputType: options.inputType }
        : null;
      refresh();
      return true;
    },
    [refresh],
  );

  const reset = useCallback(
    (markdown: string, selection: MarkdownSourceSelection): void => {
      const next = normalizeSnapshot({ markdown, selection });
      const history = historyRef.current;
      const changed =
        history.past.length > 0 ||
        history.future.length > 0 ||
        history.present.markdown !== next.markdown;
      historyRef.current = {
        past: [],
        present: next,
        future: [],
        typing: null,
      };
      if (changed) refresh();
    },
    [refresh],
  );

  const updateSelection = useCallback(
    (selection: MarkdownSourceSelection): void => {
      const history = historyRef.current;
      const normalized = normalizeSelection(
        selection,
        history.present.markdown.length,
      );
      if (sameSelection(normalized, history.present.selection)) return;
      history.present = { ...history.present, selection: normalized };
      history.typing = null;
    },
    [],
  );

  const undo = useCallback((): MarkdownSourceHistorySnapshot | null => {
    const history = historyRef.current;
    const previous = history.past.at(-1);
    if (!previous) return null;
    history.past = history.past.slice(0, -1);
    history.future = [
      ...history.future.slice(-(MARKDOWN_SOURCE_HISTORY_LIMIT - 1)),
      history.present,
    ];
    history.present = previous;
    history.typing = null;
    refresh();
    return previous;
  }, [refresh]);

  const redo = useCallback((): MarkdownSourceHistorySnapshot | null => {
    const history = historyRef.current;
    const next = history.future.at(-1);
    if (!next) return null;
    history.future = history.future.slice(0, -1);
    history.past = [
      ...history.past.slice(-(MARKDOWN_SOURCE_HISTORY_LIMIT - 1)),
      history.present,
    ];
    history.present = next;
    history.typing = null;
    refresh();
    return next;
  }, [refresh]);

  return {
    canUndo: historyRef.current.past.length > 0,
    canRedo: historyRef.current.future.length > 0,
    record,
    redo,
    reset,
    undo,
    updateSelection,
  };
}

function canCoalesceTyping(
  history: MarkdownSourceHistoryState,
  next: MarkdownSourceHistorySnapshot,
  inputType: string | undefined,
  now: number,
): boolean {
  const typing = history.typing;
  if (!typing || now - typing.at > MARKDOWN_TYPING_COALESCE_MS) return false;
  if (
    !isContinuousTypingInput(typing.inputType) ||
    !isContinuousTypingInput(inputType)
  ) {
    return false;
  }
  return (
    history.present.selection.start === history.present.selection.end &&
    next.selection.start === next.selection.end
  );
}

function isContinuousTypingInput(inputType: string | undefined): boolean {
  return (
    !inputType ||
    inputType === 'insertText' ||
    inputType === 'insertCompositionText' ||
    inputType === 'insertLineBreak' ||
    inputType === 'deleteContentBackward' ||
    inputType === 'deleteContentForward'
  );
}

function normalizeSnapshot(
  snapshot: MarkdownSourceHistorySnapshot,
): MarkdownSourceHistorySnapshot {
  return {
    markdown: snapshot.markdown,
    selection: normalizeSelection(snapshot.selection, snapshot.markdown.length),
  };
}

function normalizeSelection(
  selection: MarkdownSourceSelection,
  length: number,
): MarkdownSourceSelection {
  const start = Math.max(0, Math.min(length, selection.start));
  const end = Math.max(start, Math.min(length, selection.end));
  return { start, end, direction: selection.direction };
}

function collapsedSelection(position: number): MarkdownSourceSelection {
  return { start: position, end: position, direction: 'none' };
}

function sameSelection(
  left: MarkdownSourceSelection,
  right: MarkdownSourceSelection,
): boolean {
  return (
    left.start === right.start &&
    left.end === right.end &&
    left.direction === right.direction
  );
}
