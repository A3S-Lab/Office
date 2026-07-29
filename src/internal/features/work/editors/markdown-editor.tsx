import type { Extensions } from '@tiptap/core';
import { useEditor } from '@tiptap/react';
import {
  type CSSProperties,
  useCallback,
  useDeferredValue,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { WorkEditorLoadingState } from '../components/work-editor-loading-state';
import {
  type WorkspaceContextMenuEvent,
  workspaceContextMenuPosition,
} from '../../workspace/components/workspace-context-menu';
import {
  createWorkMarkdownExtensions,
  markdownTaskCheckboxLabel,
} from '../work-markdown-extensions';
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
import {
  type MarkdownEditingSurface,
  restoreMarkdownEditingSurfaceFocus,
} from './markdown-editor-focus';
import { MarkdownToolbar } from './markdown-toolbar';
import { type MarkdownViewMode, MarkdownWorkspace } from './markdown-workspace';
import { mergeOfficeTiptapExtensions } from './office-tiptap-extensions';
import { useMarkdownSourceHistory } from './use-markdown-source-history';
import {
  type WorkOfficeFileAction,
  WorkOfficePreviewBar,
} from './work-office-chrome';

export { markdownTaskCheckboxLabel };

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
  const editor = useEditor(
    {
      extensions,
      content: initialMarkdownRef.current,
      contentType: 'markdown',
      editable: !preview && viewMode === 'visual',
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
    },
    [extensions],
  );

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
    const visualEditorReadOnly = preview || viewMode !== 'visual';
    let checkboxFrame: number | null = null;
    const applyTaskCheckboxState = () => {
      for (const checkbox of editor.view.dom.querySelectorAll<HTMLInputElement>(
        'li[data-type="taskItem"] > label input[type="checkbox"]',
      )) {
        checkbox.disabled = visualEditorReadOnly;
        checkbox.setAttribute('aria-disabled', String(visualEditorReadOnly));
        checkbox.setAttribute(
          'aria-label',
          markdownTaskCheckboxLabel({
            attrs: { checked: checkbox.checked },
            textContent:
              checkbox.closest('li[data-type="taskItem"]')?.textContent ?? '',
          }),
        );
      }
    };
    const scheduleTaskCheckboxState = () => {
      if (checkboxFrame !== null) cancelAnimationFrame(checkboxFrame);
      checkboxFrame = requestAnimationFrame(() => {
        checkboxFrame = null;
        applyTaskCheckboxState();
      });
    };
    const applyViewState = () => {
      if (editor.isDestroyed) return;
      editor.setEditable(!visualEditorReadOnly, false);
      editor.view.dom.setAttribute(
        'aria-label',
        visualEditorReadOnly ? 'Markdown 预览' : 'Markdown 编辑区',
      );
      editor.view.dom.setAttribute(
        'aria-readonly',
        String(visualEditorReadOnly),
      );
      editor.view.dom.setAttribute(
        'role',
        visualEditorReadOnly ? 'document' : 'textbox',
      );
      if (visualEditorReadOnly) {
        editor.view.dom.removeAttribute('aria-multiline');
        editor.view.dom.tabIndex = 0;
      } else {
        editor.view.dom.setAttribute('aria-multiline', 'true');
        editor.view.dom.removeAttribute('tabindex');
      }
      applyTaskCheckboxState();
      scheduleTaskCheckboxState();
    };
    applyViewState();
    editor.on('mount', applyViewState);
    editor.on('update', applyTaskCheckboxState);
    editor.on('transaction', scheduleTaskCheckboxState);
    return () => {
      if (checkboxFrame !== null) cancelAnimationFrame(checkboxFrame);
      editor.off('mount', applyViewState);
      editor.off('update', applyTaskCheckboxState);
      editor.off('transaction', scheduleTaskCheckboxState);
    };
  }, [editor, preview, viewMode]);

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
      target?: MarkdownSourceSelectionState,
    ): boolean => {
      const current = target ?? getSourceSelection();
      if (!current) return false;
      if (current.markdown !== sourceMarkdownRef.current) return false;
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

  const handleSourceSelectionChange = useCallback(
    (selection: MarkdownSourceSelection) => {
      updateSourceHistorySelection(selection);
      setSelectionVersion((value) => value + 1);
    },
    [updateSourceHistorySelection],
  );
  const handleVisualIntent = useCallback(() => {
    applyMarkdownToEditor(sourceMarkdownRef.current);
  }, [applyMarkdownToEditor]);

  const openSourceSelectionMenu = useCallback(
    (event: WorkspaceContextMenuEvent<HTMLTextAreaElement>): boolean => {
      if (!getSelectionMenuItems) return false;
      const current = getSourceSelection();
      if (!current) return false;
      const snapshot = createWorkMarkdownSourceSelectionSnapshot(
        current.markdown,
        current.selection,
        contentRef.current,
      );
      if (!snapshot) return false;
      event.preventDefault();
      event.stopPropagation();
      const position = workspaceContextMenuPosition(event);
      const items = getSelectionMenuItems(snapshot);
      setSelectionMenu(
        items.length
          ? {
              x: position.x,
              y: position.y,
              snapshot,
              items,
            }
          : null,
      );
      return items.length > 0;
    },
    [getSelectionMenuItems, getSourceSelection],
  );

  const openVisualSelectionMenu = useCallback(
    (event: WorkspaceContextMenuEvent<HTMLElement>): boolean => {
      if (!editor || !getSelectionMenuItems) return false;
      const snapshot = createWorkMarkdownVisualSelectionSnapshot(
        editor,
        contentRef.current,
      );
      if (!snapshot) return false;
      event.preventDefault();
      event.stopPropagation();
      const position = workspaceContextMenuPosition(event);
      const items = getSelectionMenuItems(snapshot);
      setSelectionMenu(
        items.length
          ? {
              x: position.x,
              y: position.y,
              snapshot,
              items,
            }
          : null,
      );
      return items.length > 0;
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
      const focusSurface: MarkdownEditingSurface =
        mode === 'visual' ? 'visual' : 'source';
      if (mode !== 'source') {
        applyMarkdownToEditor(sourceMarkdownRef.current);
      }
      setViewMode(mode);
      restoreMarkdownEditingSurfaceFocus(() => {
        if (focusSurface === 'source') return sourceTextareaRef.current;
        return editor && !editor.isDestroyed ? editor.view.dom : null;
      });
    },
    [applyMarkdownToEditor, editor],
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
        sourceEditing={viewMode !== 'visual'}
        canSourceRedo={canRedoSource}
        canSourceUndo={canUndoSource}
        viewMode={viewMode}
        getSourceFocusTarget={() => sourceTextareaRef.current}
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
        visualReadOnly={viewMode === 'split'}
        sourceRef={sourceTextareaRef}
        sourceSelectionRequest={sourceSelectionRequest}
        onSourceChange={changeSource}
        onSourceCommand={runSourceCommand}
        onSourceContextMenu={
          getSelectionMenuItems ? openSourceSelectionMenu : undefined
        }
        onSourceRedo={redoSource}
        onSourceSelectionChange={handleSourceSelectionChange}
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
