import type { Extensions } from '@tiptap/core';
import { useEditor } from '@tiptap/react';
import {
  type CSSProperties,
  type MouseEvent,
  useCallback,
  useDeferredValue,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { WorkEditorLoadingState } from '../components/work-editor-loading-state';
import { createWorkMarkdownExtensions } from '../work-markdown-extensions';
import type { WorkMarkdownContent } from '../work-types';
import {
  createWorkMarkdownSourceSelectionAction,
  createWorkMarkdownSourceSelectionSnapshot,
  createWorkMarkdownVisualSelectionSnapshot,
  type WorkGetMarkdownSelectionMenuItems,
  type WorkMarkdownSelectionSnapshot,
} from '../work-markdown-selection-menu';
import {
  MarkdownSelectionContextMenu,
  type MarkdownSelectionMenuState,
} from './markdown-selection-context-menu';
import { MarkdownStatus } from './markdown-status';
import {
  applyMarkdownSourceCommand,
  type MarkdownSourceCommand,
  type MarkdownSourceEdit,
  type MarkdownSourceSelection,
  type MarkdownSourceSelectionState,
  replaceMarkdownSourceSelection,
} from './markdown-source-commands';
import { MarkdownToolbar } from './markdown-toolbar';
import { type MarkdownViewMode, MarkdownWorkspace } from './markdown-workspace';
import { mergeOfficeTiptapExtensions } from './office-tiptap-extensions';
import { useMarkdownSourceHistory } from './use-markdown-source-history';
import {
  type WorkOfficeFileAction,
  WorkOfficePreviewBar,
} from './work-office-chrome';

export { markdownTaskCheckboxLabel } from '../work-markdown-extensions';

export interface MarkdownEditorProps {
  content: WorkMarkdownContent;
  extensions?: Extensions;
  preview: boolean;
  saveStatus?: string;
  fileActions?: readonly WorkOfficeFileAction[];
  getSelectionMenuItems?: WorkGetMarkdownSelectionMenuItems;
  onChange: (content: WorkMarkdownContent) => void;
}

const MARKDOWN_PREVIEW_SYNC_DELAY = 160;
const EMPTY_MARKDOWN_EXTENSIONS: Extensions = [];

export function MarkdownEditor({
  content,
  extensions: additionalExtensions = EMPTY_MARKDOWN_EXTENSIONS,
  preview,
  saveStatus,
  fileActions,
  getSelectionMenuItems,
  onChange,
}: MarkdownEditorProps) {
  const contentRef = useRef(content);
  const onChangeRef = useRef(onChange);
  const receivedContentRef = useRef(content);
  const appliedMarkdownRef = useRef(content.markdown);
  const emittedMarkdownRef = useRef<string | null>(null);
  const sourceMarkdownRef = useRef(content.markdown);
  const initialMarkdownRef = useRef(content.markdown);
  const previewTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const sourceTextareaRef = useRef<HTMLTextAreaElement>(null);
  const [sourceMarkdown, setSourceMarkdown] = useState(content.markdown);
  const [editingSurface, setEditingSurface] = useState<'source' | 'visual'>(
    'visual',
  );
  const [selectionMenu, setSelectionMenu] =
    useState<MarkdownSelectionMenuState | null>(null);
  const [sourceSelectionRequest, setSourceSelectionRequest] = useState<
    | (MarkdownSourceEdit['selection'] & {
        revision: number;
      })
    | undefined
  >();
  const [viewMode, setViewMode] = useState<MarkdownViewMode>('split');
  const [zoom, setZoom] = useState(100);
  const [, setSelectionVersion] = useState(0);
  const {
    canRedo: canRedoSource,
    canUndo: canUndoSource,
    record: recordSourceHistory,
    redo: redoSourceHistory,
    reset: resetSourceHistory,
    undo: undoSourceHistory,
    updateSelection: updateSourceHistorySelection,
  } = useMarkdownSourceHistory(content.markdown);
  contentRef.current = content;
  onChangeRef.current = onChange;

  const cancelPreviewSync = useCallback(() => {
    if (previewTimerRef.current === null) return;
    clearTimeout(previewTimerRef.current);
    previewTimerRef.current = null;
  }, []);

  const extensions = useMemo(
    () =>
      mergeOfficeTiptapExtensions(
        'MarkdownEditor',
        createWorkMarkdownExtensions(),
        additionalExtensions,
      ),
    [additionalExtensions],
  );
  const editorProps = useMemo(
    () => ({
      attributes: {
        'aria-label': 'Markdown 编辑区',
        'aria-multiline': 'true',
        role: 'textbox',
        spellcheck: 'true',
      },
    }),
    [],
  );
  const editor = useEditor({
    extensions,
    content: initialMarkdownRef.current,
    contentType: 'markdown',
    editable: !preview,
    editorProps,
    onUpdate: ({ editor: current }) => {
      const markdown = current.getMarkdown();
      if (markdown === appliedMarkdownRef.current) return;
      cancelPreviewSync();
      const next = { ...contentRef.current, markdown };
      appliedMarkdownRef.current = markdown;
      emittedMarkdownRef.current = markdown;
      sourceMarkdownRef.current = markdown;
      contentRef.current = next;
      setSourceMarkdown(markdown);
      resetSourceHistory(
        markdown,
        textareaSelection(sourceTextareaRef.current, markdown.length),
      );
      onChangeRef.current(next);
    },
    onSelectionUpdate: () => setSelectionVersion((value) => value + 1),
  });

  const applyMarkdownToEditor = useCallback(
    (markdown: string) => {
      cancelPreviewSync();
      if (
        !editor ||
        editor.isDestroyed ||
        appliedMarkdownRef.current === markdown
      ) {
        return;
      }
      appliedMarkdownRef.current = markdown;
      editor.commands.setWorkMarkdown(markdown, { emitUpdate: false });
    },
    [cancelPreviewSync, editor],
  );

  const queueMarkdownPreview = useCallback(
    (markdown: string, immediate = false) => {
      cancelPreviewSync();
      if (
        !editor ||
        editor.isDestroyed ||
        appliedMarkdownRef.current === markdown
      ) {
        return;
      }
      if (immediate) {
        applyMarkdownToEditor(markdown);
        return;
      }
      previewTimerRef.current = setTimeout(() => {
        previewTimerRef.current = null;
        applyMarkdownToEditor(markdown);
      }, MARKDOWN_PREVIEW_SYNC_DELAY);
    },
    [applyMarkdownToEditor, cancelPreviewSync, editor],
  );

  useEffect(() => cancelPreviewSync, [cancelPreviewSync]);

  useEffect(() => {
    if (!editor) return;
    const applyViewState = () => {
      if (editor.isDestroyed) return;
      editor.setEditable(!preview);
      editor.view.dom.setAttribute(
        'aria-label',
        preview ? 'Markdown 预览' : 'Markdown 编辑区',
      );
      editor.view.dom.setAttribute('aria-readonly', String(preview));
      editor.view.dom.setAttribute('role', preview ? 'document' : 'textbox');
      if (preview) editor.view.dom.removeAttribute('aria-multiline');
      else editor.view.dom.setAttribute('aria-multiline', 'true');
    };
    applyViewState();
    editor.on('mount', applyViewState);
    return () => {
      editor.off('mount', applyViewState);
    };
  }, [editor, preview]);

  useEffect(() => {
    if (!editor || receivedContentRef.current === content) return;
    receivedContentRef.current = content;
    const markdown = content.markdown;
    if (sourceMarkdownRef.current !== markdown) {
      sourceMarkdownRef.current = markdown;
      setSourceMarkdown(markdown);
    }
    if (emittedMarkdownRef.current === markdown) {
      emittedMarkdownRef.current = null;
      return;
    }
    emittedMarkdownRef.current = null;
    resetSourceHistory(
      markdown,
      textareaSelection(sourceTextareaRef.current, markdown.length),
    );
    if (preview || viewMode !== 'source') {
      queueMarkdownPreview(markdown, true);
    }
  }, [
    content,
    editor,
    preview,
    queueMarkdownPreview,
    resetSourceHistory,
    viewMode,
  ]);

  const updateSource = useCallback(
    (markdown: string) => {
      if (markdown === sourceMarkdownRef.current) return;
      sourceMarkdownRef.current = markdown;
      setSourceMarkdown(markdown);
      const next = { ...contentRef.current, markdown };
      contentRef.current = next;
      emittedMarkdownRef.current = markdown;
      onChangeRef.current(next);
      if (viewMode !== 'source') {
        queueMarkdownPreview(markdown);
      }
    },
    [queueMarkdownPreview, viewMode],
  );

  const getSourceSelection =
    useCallback((): MarkdownSourceSelectionState | null => {
      const source = sourceTextareaRef.current;
      if (!source) return null;
      const selection = textareaSelection(
        source,
        sourceMarkdownRef.current.length,
      );
      return {
        markdown: sourceMarkdownRef.current,
        selection,
        text: sourceMarkdownRef.current.slice(selection.start, selection.end),
      };
    }, []);

  const requestSourceSelection = useCallback(
    (selection: MarkdownSourceSelection) => {
      setSourceSelectionRequest((current) => ({
        ...selection,
        revision: (current?.revision ?? 0) + 1,
      }));
    },
    [],
  );

  const applySourceEdit = useCallback(
    (edit: MarkdownSourceEdit): boolean => {
      recordSourceHistory(edit);
      updateSource(edit.markdown);
      requestSourceSelection(edit.selection);
      return true;
    },
    [recordSourceHistory, requestSourceSelection, updateSource],
  );

  const changeSource = useCallback(
    (
      markdown: string,
      selection: MarkdownSourceSelection,
      inputType?: string,
    ) => {
      const edit = { markdown, selection };
      if (!recordSourceHistory(edit, { typing: true, inputType })) return;
      updateSource(markdown);
    },
    [recordSourceHistory, updateSource],
  );

  const applySourceHistory = useCallback(
    (edit: MarkdownSourceEdit | null): boolean => {
      if (!edit) return false;
      updateSource(edit.markdown);
      requestSourceSelection(edit.selection);
      return true;
    },
    [requestSourceSelection, updateSource],
  );

  const undoSource = useCallback(
    () => applySourceHistory(undoSourceHistory()),
    [applySourceHistory, undoSourceHistory],
  );
  const redoSource = useCallback(
    () => applySourceHistory(redoSourceHistory()),
    [applySourceHistory, redoSourceHistory],
  );

  const runSourceCommand = useCallback(
    (command: MarkdownSourceCommand): boolean => {
      const current = getSourceSelection();
      if (!current) return false;
      return applySourceEdit(
        applyMarkdownSourceCommand(
          current.markdown,
          current.selection,
          command,
        ),
      );
    },
    [applySourceEdit, getSourceSelection],
  );

  const replaceSourceSelection = useCallback(
    (
      replacement: string,
      selectedRange?: { start: number; end: number },
    ): boolean => {
      const current = getSourceSelection();
      if (!current) return false;
      return applySourceEdit(
        replaceMarkdownSourceSelection(
          current.markdown,
          current.selection,
          replacement,
          selectedRange,
        ),
      );
    },
    [applySourceEdit, getSourceSelection],
  );

  const handleSourceIntent = useCallback(() => setEditingSurface('source'), []);
  const handleVisualIntent = useCallback(() => {
    setEditingSurface('visual');
    applyMarkdownToEditor(sourceMarkdownRef.current);
  }, [applyMarkdownToEditor]);

  const openSourceSelectionMenu = useCallback(
    (event: MouseEvent<HTMLTextAreaElement>) => {
      if (!getSelectionMenuItems) return;
      const current = getSourceSelection();
      if (!current) return;
      const snapshot = createWorkMarkdownSourceSelectionSnapshot(
        current.markdown,
        current.selection,
        contentRef.current,
      );
      if (!snapshot) return;
      event.preventDefault();
      const items = getSelectionMenuItems(snapshot);
      setSelectionMenu(
        items.length
          ? {
              x: event.clientX,
              y: event.clientY,
              snapshot,
              items,
            }
          : null,
      );
    },
    [getSelectionMenuItems, getSourceSelection],
  );

  const openVisualSelectionMenu = useCallback(
    (event: MouseEvent<HTMLElement>) => {
      if (!editor || !getSelectionMenuItems) return;
      const snapshot = createWorkMarkdownVisualSelectionSnapshot(
        editor,
        contentRef.current,
      );
      if (!snapshot) return;
      event.preventDefault();
      const items = getSelectionMenuItems(snapshot);
      setSelectionMenu(
        items.length
          ? {
              x: event.clientX,
              y: event.clientY,
              snapshot,
              items,
            }
          : null,
      );
    },
    [editor, getSelectionMenuItems],
  );

  const createSourceSelectionAction = useCallback(
    (snapshot: WorkMarkdownSelectionSnapshot) =>
      createWorkMarkdownSourceSelectionAction(
        snapshot,
        () => sourceMarkdownRef.current,
        applySourceEdit,
      ),
    [applySourceEdit],
  );

  const changeViewMode = useCallback(
    (mode: MarkdownViewMode) => {
      if (mode !== 'source') {
        applyMarkdownToEditor(sourceMarkdownRef.current);
      }
      if (mode === 'visual') setEditingSurface('visual');
      if (mode === 'source') setEditingSurface('source');
      setViewMode(mode);
    },
    [applyMarkdownToEditor],
  );

  const deferredMarkdown = useDeferredValue(sourceMarkdown);
  const metrics = useMemo(
    () => markdownMetrics(deferredMarkdown),
    [deferredMarkdown],
  );
  const editorStyle = {
    '--work-markdown-zoom': zoom,
  } as CSSProperties;

  if (!editor) {
    return <WorkEditorLoadingState title="正在准备 Markdown 编辑器" />;
  }

  if (preview) {
    return (
      <section className="work-markdown-editor preview" style={editorStyle}>
        <WorkOfficePreviewBar
          ariaLabel="Markdown 预览工具"
          label="只读预览"
          detail={`${metrics.lineCount} 行`}
          fileActions={fileActions}
          className="work-markdown-ribbon"
        />
        <MarkdownWorkspace
          editor={editor}
          markdown={sourceMarkdown}
          mode="visual"
          readOnly
          sourceRef={sourceTextareaRef}
          onSourceChange={updateSource}
          onVisualIntent={handleVisualIntent}
        />
        <MarkdownStatus
          characterCount={metrics.characterCount}
          lineCount={metrics.lineCount}
          saveStatus={saveStatus}
          zoom={zoom}
          onZoomChange={setZoom}
        />
      </section>
    );
  }

  return (
    <section className="work-markdown-editor" style={editorStyle}>
      <MarkdownToolbar
        editor={editor}
        fileActions={fileActions}
        sourceEditing={editingSurface === 'source'}
        canSourceRedo={canRedoSource}
        canSourceUndo={canUndoSource}
        viewMode={viewMode}
        getSourceSelection={getSourceSelection}
        onSourceCommand={runSourceCommand}
        onSourceRedo={redoSource}
        onSourceReplace={replaceSourceSelection}
        onSourceUndo={undoSource}
        onViewModeChange={changeViewMode}
      />
      <MarkdownWorkspace
        editor={editor}
        markdown={sourceMarkdown}
        mode={viewMode}
        sourceRef={sourceTextareaRef}
        sourceSelectionRequest={sourceSelectionRequest}
        onSourceChange={changeSource}
        onSourceCommand={runSourceCommand}
        onSourceContextMenu={
          getSelectionMenuItems ? openSourceSelectionMenu : undefined
        }
        onSourceIntent={handleSourceIntent}
        onSourceRedo={redoSource}
        onSourceSelectionChange={updateSourceHistorySelection}
        onSourceUndo={undoSource}
        onVisualContextMenu={
          getSelectionMenuItems ? openVisualSelectionMenu : undefined
        }
        onVisualIntent={handleVisualIntent}
      />
      <MarkdownStatus
        characterCount={metrics.characterCount}
        lineCount={metrics.lineCount}
        saveStatus={saveStatus}
        zoom={zoom}
        onZoomChange={setZoom}
      />
      {selectionMenu && (
        <MarkdownSelectionContextMenu
          editor={editor}
          menu={selectionMenu}
          createSourceAction={createSourceSelectionAction}
          onClose={() => setSelectionMenu(null)}
        />
      )}
    </section>
  );
}

function textareaSelection(
  textarea: HTMLTextAreaElement | null,
  fallbackPosition: number,
): MarkdownSourceSelection {
  if (!textarea) {
    return {
      start: fallbackPosition,
      end: fallbackPosition,
      direction: 'none',
    };
  }
  return {
    start: textarea.selectionStart,
    end: textarea.selectionEnd,
    direction: textarea.selectionDirection,
  };
}

function markdownMetrics(markdown: string): {
  lineCount: number;
  characterCount: number;
} {
  return {
    lineCount: markdown
      ? markdown.replace(/\r\n?/g, '\n').split('\n').length
      : 1,
    characterCount: Array.from(markdown).length,
  };
}
