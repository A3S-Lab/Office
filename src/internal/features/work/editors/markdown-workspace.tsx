import type { Editor } from '@tiptap/core';
import { EditorContent } from '@tiptap/react';
import {
  type CSSProperties,
  type FormEvent,
  type KeyboardEvent,
  type MouseEvent,
  type PointerEvent,
  type RefObject,
  type UIEvent,
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import type {
  MarkdownSourceCommand,
  MarkdownSourceSelection,
} from './markdown-source-commands';

export type MarkdownViewMode = 'visual' | 'source' | 'split';

const DEFAULT_MARKDOWN_SPLIT_PERCENT = 50;
const MIN_MARKDOWN_SPLIT_PERCENT = 30;
const MAX_MARKDOWN_SPLIT_PERCENT = 70;
const MARKDOWN_SPLIT_KEYBOARD_STEP = 5;

export function proportionalMarkdownScrollTop(
  sourceScrollTop: number,
  sourceScrollHeight: number,
  sourceClientHeight: number,
  targetScrollHeight: number,
  targetClientHeight: number,
): number {
  const sourceRange = Math.max(0, sourceScrollHeight - sourceClientHeight);
  const targetRange = Math.max(0, targetScrollHeight - targetClientHeight);
  if (!sourceRange || !targetRange) return 0;

  const progress = Math.min(1, Math.max(0, sourceScrollTop / sourceRange));
  return progress * targetRange;
}

export function MarkdownWorkspace({
  editor,
  markdown,
  mode,
  readOnly = false,
  sourceRef,
  sourceSelectionRequest,
  onSourceChange,
  onSourceCommand,
  onSourceContextMenu,
  onSourceIntent,
  onSourceRedo,
  onSourceSelectionChange,
  onSourceUndo,
  onVisualContextMenu,
  onVisualIntent,
}: {
  editor: Editor;
  markdown: string;
  mode: MarkdownViewMode;
  readOnly?: boolean;
  sourceRef: RefObject<HTMLTextAreaElement | null>;
  sourceSelectionRequest?: MarkdownSourceSelection & { revision: number };
  onSourceChange: (
    markdown: string,
    selection: MarkdownSourceSelection,
    inputType?: string,
  ) => void;
  onSourceCommand?: (command: MarkdownSourceCommand) => boolean;
  onSourceContextMenu?: (event: MouseEvent<HTMLTextAreaElement>) => void;
  onSourceIntent?: () => void;
  onSourceRedo?: () => boolean;
  onSourceSelectionChange?: (selection: MarkdownSourceSelection) => void;
  onSourceUndo?: () => boolean;
  onVisualContextMenu?: (event: MouseEvent<HTMLElement>) => void;
  onVisualIntent?: () => void;
}) {
  const workspaceRef = useRef<HTMLDivElement>(null);
  const visualRef = useRef<HTMLElement>(null);
  const synchronizedTargetRef = useRef<'source' | 'visual' | null>(null);
  const releaseFrameRef = useRef<number | null>(null);
  const sourceInputTypeRef = useRef<string | undefined>(undefined);
  const [sourcePanePercent, setSourcePanePercent] = useState(
    DEFAULT_MARKDOWN_SPLIT_PERCENT,
  );
  const [resizing, setResizing] = useState(false);
  const showSource = !readOnly && mode !== 'visual';
  const showVisual = readOnly || mode !== 'source';

  const releaseScrollLock = useCallback(() => {
    if (releaseFrameRef.current !== null) {
      cancelAnimationFrame(releaseFrameRef.current);
    }
    releaseFrameRef.current = requestAnimationFrame(() => {
      synchronizedTargetRef.current = null;
      releaseFrameRef.current = null;
    });
  }, []);

  useEffect(
    () => () => {
      if (releaseFrameRef.current !== null) {
        cancelAnimationFrame(releaseFrameRef.current);
      }
    },
    [],
  );

  useLayoutEffect(() => {
    if (!sourceSelectionRequest) return;
    const source = sourceRef.current;
    if (!source) return;
    source.focus({ preventScroll: true });
    source.setSelectionRange(
      sourceSelectionRequest.start,
      sourceSelectionRequest.end,
      sourceSelectionRequest.direction,
    );
  }, [sourceRef, sourceSelectionRequest]);

  const handleSourceScroll = useCallback(
    (event: UIEvent<HTMLTextAreaElement>) => {
      if (mode !== 'split') return;
      if (synchronizedTargetRef.current === 'source') {
        synchronizedTargetRef.current = null;
        return;
      }

      const source = event.currentTarget;
      const target = visualRef.current;
      if (!target) return;
      synchronizedTargetRef.current = 'visual';
      target.scrollTop = proportionalMarkdownScrollTop(
        source.scrollTop,
        source.scrollHeight,
        source.clientHeight,
        target.scrollHeight,
        target.clientHeight,
      );
      releaseScrollLock();
    },
    [mode, releaseScrollLock],
  );

  const handleVisualScroll = useCallback(
    (event: UIEvent<HTMLElement>) => {
      if (mode !== 'split') return;
      if (synchronizedTargetRef.current === 'visual') {
        synchronizedTargetRef.current = null;
        return;
      }

      const source = event.currentTarget;
      const target = sourceRef.current;
      if (!target) return;
      synchronizedTargetRef.current = 'source';
      target.scrollTop = proportionalMarkdownScrollTop(
        source.scrollTop,
        source.scrollHeight,
        source.clientHeight,
        target.scrollHeight,
        target.clientHeight,
      );
      releaseScrollLock();
    },
    [mode, releaseScrollLock],
  );

  const resizeFromPointer = useCallback((clientX: number) => {
    const workspace = workspaceRef.current;
    if (!workspace) return;
    const bounds = workspace.getBoundingClientRect();
    if (bounds.width <= 0) return;
    setSourcePanePercent(
      clampMarkdownSplitPercent(
        Math.round(((clientX - bounds.left) / bounds.width) * 1000) / 10,
      ),
    );
  }, []);

  const handleResizeStart = (event: PointerEvent<HTMLHRElement>) => {
    event.preventDefault();
    event.currentTarget.setPointerCapture(event.pointerId);
    setResizing(true);
    resizeFromPointer(event.clientX);
  };

  const handleResizeMove = (event: PointerEvent<HTMLHRElement>) => {
    if (!event.currentTarget.hasPointerCapture(event.pointerId)) return;
    resizeFromPointer(event.clientX);
  };

  const handleResizeEnd = (event: PointerEvent<HTMLHRElement>) => {
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }
    setResizing(false);
  };

  const handleResizeKeyDown = (event: KeyboardEvent<HTMLHRElement>) => {
    let next: number | null = null;
    if (event.key === 'ArrowLeft') {
      next = sourcePanePercent - MARKDOWN_SPLIT_KEYBOARD_STEP;
    } else if (event.key === 'ArrowRight') {
      next = sourcePanePercent + MARKDOWN_SPLIT_KEYBOARD_STEP;
    } else if (event.key === 'Home') {
      next = MIN_MARKDOWN_SPLIT_PERCENT;
    } else if (event.key === 'End') {
      next = MAX_MARKDOWN_SPLIT_PERCENT;
    } else if (event.key === 'Enter') {
      next = DEFAULT_MARKDOWN_SPLIT_PERCENT;
    }
    if (next === null) return;
    event.preventDefault();
    setSourcePanePercent(clampMarkdownSplitPercent(next));
  };

  const workspaceStyle = {
    '--work-markdown-source-pane': `${sourcePanePercent}%`,
  } as CSSProperties;

  return (
    <div
      ref={workspaceRef}
      className={`work-markdown-workspace ${mode}`}
      data-split-resizing={resizing || undefined}
      style={workspaceStyle}
    >
      {showSource && (
        <section
          aria-label="Markdown 源码窗格"
          className="work-markdown-pane source"
        >
          {mode === 'split' && (
            <header className="work-markdown-pane-label">源码</header>
          )}
          <textarea
            ref={sourceRef}
            aria-label="Markdown 源码"
            value={markdown}
            spellCheck
            onBeforeInput={(event: FormEvent<HTMLTextAreaElement>) => {
              const inputType = (event.nativeEvent as InputEvent).inputType;
              if (inputType === 'historyUndo') {
                event.preventDefault();
                onSourceUndo?.();
                return;
              }
              if (inputType === 'historyRedo') {
                event.preventDefault();
                onSourceRedo?.();
                return;
              }
              sourceInputTypeRef.current = inputType || undefined;
            }}
            onChange={(event) => {
              onSourceIntent?.();
              const source = event.currentTarget;
              onSourceChange(
                source.value,
                sourceSelection(source),
                sourceInputTypeRef.current,
              );
              sourceInputTypeRef.current = undefined;
            }}
            onFocus={onSourceIntent}
            onContextMenu={onSourceContextMenu}
            onKeyDown={(event) => {
              if (event.altKey || !(event.metaKey || event.ctrlKey)) return;
              const key = event.key.toLocaleLowerCase();
              if (key === 'z') {
                event.preventDefault();
                if (event.shiftKey) onSourceRedo?.();
                else onSourceUndo?.();
                return;
              }
              if (key === 'y' && !event.shiftKey) {
                event.preventDefault();
                onSourceRedo?.();
                return;
              }
              if (event.shiftKey) return;
              const command =
                key === 'b' ? 'bold' : key === 'i' ? 'italic' : null;
              if (!command || !onSourceCommand?.(command)) return;
              event.preventDefault();
            }}
            onPointerDown={onSourceIntent}
            onSelect={(event) => {
              onSourceIntent?.();
              onSourceSelectionChange?.(sourceSelection(event.currentTarget));
            }}
            onScroll={handleSourceScroll}
          />
        </section>
      )}
      {mode === 'split' && (
        <hr
          className="work-markdown-splitter"
          aria-label="调整源码与编辑结果宽度"
          aria-orientation="vertical"
          aria-valuemin={MIN_MARKDOWN_SPLIT_PERCENT}
          aria-valuemax={MAX_MARKDOWN_SPLIT_PERCENT}
          aria-valuenow={sourcePanePercent}
          aria-valuetext={`源码窗格 ${sourcePanePercent}%`}
          tabIndex={0}
          title="拖动调整分栏宽度，双击恢复均分"
          onDoubleClick={() =>
            setSourcePanePercent(DEFAULT_MARKDOWN_SPLIT_PERCENT)
          }
          onKeyDown={handleResizeKeyDown}
          onPointerCancel={handleResizeEnd}
          onPointerDown={handleResizeStart}
          onPointerMove={handleResizeMove}
          onPointerUp={handleResizeEnd}
        />
      )}
      {showVisual && (
        <section
          ref={visualRef}
          aria-label={readOnly ? 'Markdown 预览窗格' : 'Markdown 编辑结果窗格'}
          className="work-markdown-pane visual"
          onContextMenu={onVisualContextMenu}
          onFocusCapture={onVisualIntent}
          onPointerDownCapture={onVisualIntent}
          onScroll={handleVisualScroll}
        >
          {mode === 'split' && (
            <header className="work-markdown-pane-label">编辑结果</header>
          )}
          <div className="work-markdown-canvas">
            <EditorContent editor={editor} />
          </div>
        </section>
      )}
    </div>
  );
}

function sourceSelection(source: HTMLTextAreaElement): MarkdownSourceSelection {
  return {
    start: source.selectionStart,
    end: source.selectionEnd,
    direction: source.selectionDirection,
  };
}

function clampMarkdownSplitPercent(value: number): number {
  return Math.min(
    MAX_MARKDOWN_SPLIT_PERCENT,
    Math.max(MIN_MARKDOWN_SPLIT_PERCENT, value),
  );
}
