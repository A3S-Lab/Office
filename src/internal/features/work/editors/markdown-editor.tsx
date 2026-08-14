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
import {
  createWorkOfficeMarkdownCollaborationBinding,
  readWorkOfficeMarkdownCollaboration,
  type WorkOfficeMarkdownCollaborationBinding,
} from '../../../collaboration/office-markdown-collaboration';
import type { WorkOfficeCollaborationSession } from '../../../collaboration/office-collaboration';
import {
  type WorkspaceContextMenuEvent,
  workspaceContextMenuPosition,
  workspaceTextControlSelectionBounds,
} from '../../workspace/components/workspace-context-menu';
import { WorkEditorLoadingState } from '../components/work-editor-loading-state';
import {
  createWorkMarkdownExtensions,
  markdownTaskCheckboxLabel,
} from '../work-markdown-extensions';
import {
  createWorkMarkdownSourceSelectionAction,
  createWorkMarkdownSourceSelectionSnapshot,
  createWorkMarkdownVisualSelectionSnapshot,
  type WorkGetMarkdownSelectionMenuItems,
  type WorkMarkdownSelectionSnapshot,
} from '../work-markdown-selection-menu';
import type { WorkMarkdownContent } from '../work-types';
import {
  type MarkdownEditingSurface,
  restoreMarkdownEditingSurfaceFocus,
} from './markdown-editor-focus';
import {
  MarkdownSelectionContextMenu,
  type MarkdownSelectionMenuState,
} from './markdown-selection-context-menu';
import {
  applyMarkdownSourceCommand,
  type MarkdownSourceCommand,
  type MarkdownSourceEdit,
  type MarkdownSourceSelection,
  type MarkdownSourceSelectionState,
  replaceMarkdownSourceSelection,
} from './markdown-source-commands';
import { MarkdownStatus } from './markdown-status';
import { MarkdownToolbar } from './markdown-toolbar';
import { type MarkdownViewMode, MarkdownWorkspace } from './markdown-workspace';
import { useOfficeEditorFocusOrigin } from './office-editor-focus-handoff';
import { mergeOfficeTiptapExtensions } from './office-tiptap-extensions';
import { useMarkdownSourceHistory } from './use-markdown-source-history';
import {
  stepOfficeZoom,
  useOfficeEditorWheelZoom,
} from './use-office-editor-wheel-zoom';
import {
  type WorkOfficeFileAction,
  WorkOfficePreviewBar,
} from './work-office-chrome';

export { markdownTaskCheckboxLabel };

export interface MarkdownEditorProps {
  /** The host must initialize and synchronize this session before mounting. */
  collaboration?: WorkOfficeCollaborationSession;
  autoFocus?: boolean;
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
  autoFocus = true,
  collaboration,
  content,
  extensions: additionalExtensions = EMPTY_MARKDOWN_EXTENSIONS,
  preview,
  saveStatus,
  fileActions,
  getSelectionMenuItems,
  onChange,
}: MarkdownEditorProps) {
  const collaborationRef = useRef(collaboration);
  if (collaborationRef.current !== collaboration) {
    throw new Error(
      'MarkdownEditor collaboration sessions cannot be replaced while mounted. Remount the editor for another shared document.',
    );
  }
  const collaborationBindingRef = useRef<
    WorkOfficeMarkdownCollaborationBinding | undefined
  >(undefined);
  const collaborative = collaboration !== undefined;
  const collaborationEditable = !collaboration || collaboration.mode === 'edit';
  const readOnly = preview || !collaborationEditable;
  const initialContent = collaboration
    ? readWorkOfficeMarkdownCollaboration(collaboration)
    : content;
  const contentRef = useRef(initialContent);
  const onChangeRef = useRef(onChange);
  const receivedContentRef = useRef(content);
  const appliedMarkdownRef = useRef(initialContent.markdown);
  const emittedMarkdownRef = useRef<string | null>(null);
  const sourceMarkdownRef = useRef(initialContent.markdown);
  const initialMarkdownRef = useRef(initialContent.markdown);
  const previewTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const sourceTextareaRef = useRef<HTMLTextAreaElement>(null);
  const markdownRootRef = useRef<HTMLElement>(null);
  const initialEditorFocusRequestedRef = useRef(false);
  const editorFocusOrigin = useOfficeEditorFocusOrigin();
  const [sourceMarkdown, setSourceMarkdown] = useState(initialContent.markdown);
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
  } = useMarkdownSourceHistory(initialContent.markdown);
  if (!collaborative) contentRef.current = content;
  onChangeRef.current = onChange;

  const [, refreshCollaborationHistory] = useState(0);
  const [collaborationReady, setCollaborationReady] = useState(!collaborative);

  const cancelPreviewSync = useCallback(() => {
    if (previewTimerRef.current === null) return;
    clearTimeout(previewTimerRef.current);
    previewTimerRef.current = null;
  }, []);

  const extensions = useMemo(
    () =>
      mergeOfficeTiptapExtensions(
        'MarkdownEditor',
        createWorkMarkdownExtensions({ collaborative }),
        additionalExtensions,
      ),
    [additionalExtensions, collaborative],
  );
  const editorProps = useMemo(
    () => ({
      attributes: {
        'aria-label': 'Markdown 编辑区',
        'aria-multiline': 'true',
        role: 'textbox',
        spellcheck: 'true',
      },
      handleKeyDown: (_view: unknown, event: KeyboardEvent) => {
        if (
          !collaborative ||
          readOnly ||
          event.altKey ||
          !(event.metaKey || event.ctrlKey)
        ) {
          return false;
        }
        const key = event.key.toLocaleLowerCase();
        if (key === 'z') {
          if (event.shiftKey) collaborationBindingRef.current?.redo();
          else collaborationBindingRef.current?.undo();
          return true;
        }
        if (key === 'y' && !event.shiftKey) {
          collaborationBindingRef.current?.redo();
          return true;
        }
        return false;
      },
    }),
    [collaborative, readOnly],
  );
  const editor = useEditor(
    {
      extensions,
      content: initialMarkdownRef.current,
      contentType: 'markdown',
      editable: !readOnly && viewMode === 'visual',
      editorProps,
      onUpdate: ({ editor: current }) => {
        const markdown = current.getMarkdown();
        if (markdown === appliedMarkdownRef.current) return;
        cancelPreviewSync();
        if (collaborative) {
          appliedMarkdownRef.current = markdown;
          collaborationBindingRef.current?.replace(markdown);
          return;
        }
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

  useEffect(() => {
    if (!collaboration) return;
    const binding = createWorkOfficeMarkdownCollaborationBinding(collaboration);
    collaborationBindingRef.current = binding;
    setCollaborationReady(true);
    return () => {
      if (collaborationBindingRef.current === binding) {
        collaborationBindingRef.current = undefined;
      }
      binding.destroy();
    };
  }, [collaboration]);

  useEffect(() => {
    if (!collaboration) return;
    const binding = collaborationBindingRef.current;
    if (!binding) return;
    const unsubscribeContent = binding.subscribe(({ content: next }) => {
      contentRef.current = next;
      sourceMarkdownRef.current = next.markdown;
      setSourceMarkdown(next.markdown);
      if (preview || viewMode !== 'source') {
        queueMarkdownPreview(next.markdown, viewMode === 'visual');
      }
      onChangeRef.current(next);
    });
    const unsubscribeHistory = binding.subscribeHistory(() =>
      refreshCollaborationHistory((value) => value + 1),
    );
    const current = binding.content();
    if (current.markdown !== sourceMarkdownRef.current) {
      contentRef.current = current;
      sourceMarkdownRef.current = current.markdown;
      setSourceMarkdown(current.markdown);
      if (preview || viewMode !== 'source') {
        queueMarkdownPreview(current.markdown, viewMode === 'visual');
      }
      onChangeRef.current(current);
    }
    return () => {
      unsubscribeContent();
      unsubscribeHistory();
    };
  }, [collaboration, preview, queueMarkdownPreview, viewMode]);

  useEffect(() => cancelPreviewSync, [cancelPreviewSync]);

  useEffect(() => {
    if (!editor) return;
    const visualEditorReadOnly = readOnly || viewMode !== 'visual';
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
  }, [editor, readOnly, viewMode]);

  useEffect(() => {
    if (collaborative) return;
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
    if (readOnly || viewMode !== 'source') {
      queueMarkdownPreview(markdown, true);
    }
  }, [
    content,
    collaborative,
    editor,
    readOnly,
    queueMarkdownPreview,
    resetSourceHistory,
    viewMode,
  ]);

  const updateSource = useCallback(
    (markdown: string) => {
      if (markdown === sourceMarkdownRef.current) return;
      if (collaborative) {
        collaborationBindingRef.current?.replace(markdown);
        return;
      }
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
    [collaborative, queueMarkdownPreview, viewMode],
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
      if (!collaborative) recordSourceHistory(edit);
      updateSource(edit.markdown);
      requestSourceSelection(edit.selection);
      return true;
    },
    [collaborative, recordSourceHistory, requestSourceSelection, updateSource],
  );

  const changeSource = useCallback(
    (
      markdown: string,
      selection: MarkdownSourceSelection,
      inputType?: string,
    ) => {
      const edit = { markdown, selection };
      if (!collaborative) {
        if (!recordSourceHistory(edit, { typing: true, inputType })) return;
      }
      updateSource(markdown);
    },
    [collaborative, recordSourceHistory, updateSource],
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

  const undoSource = useCallback(() => {
    if (collaborative) return collaborationBindingRef.current?.undo() ?? false;
    return applySourceHistory(undoSourceHistory());
  }, [applySourceHistory, collaborative, undoSourceHistory]);
  const redoSource = useCallback(() => {
    if (collaborative) return collaborationBindingRef.current?.redo() ?? false;
    return applySourceHistory(redoSourceHistory());
  }, [applySourceHistory, collaborative, redoSourceHistory]);

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
      if (!collaborative) updateSourceHistorySelection(selection);
      setSelectionVersion((value) => value + 1);
    },
    [collaborative, updateSourceHistorySelection],
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
      const position = workspaceContextMenuPosition(
        event,
        workspaceTextControlSelectionBounds(event.currentTarget),
      );
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
  const canUndoMarkdown = collaborative
    ? (collaborationBindingRef.current?.canUndo() ?? false)
    : canUndoSource;
  const canRedoMarkdown = collaborative
    ? (collaborationBindingRef.current?.canRedo() ?? false)
    : canRedoSource;
  const metrics = useMemo(
    () => markdownMetrics(deferredMarkdown),
    [deferredMarkdown],
  );
  const editorStyle = {
    '--work-markdown-zoom': zoom,
  } as CSSProperties;
  useOfficeEditorWheelZoom({
    enabled: Boolean(editor),
    scopeRef: markdownRootRef,
    onZoomIn: () =>
      setZoom((current) =>
        stepOfficeZoom(current, 'in', { minimum: 60, maximum: 180 }),
      ),
    onZoomOut: () =>
      setZoom((current) =>
        stepOfficeZoom(current, 'out', { minimum: 60, maximum: 180 }),
      ),
  });

  useEffect(() => {
    if (
      !autoFocus ||
      readOnly ||
      !editor ||
      !collaborationReady ||
      initialEditorFocusRequestedRef.current
    ) {
      return;
    }
    initialEditorFocusRequestedRef.current = true;
    restoreMarkdownEditingSurfaceFocus(() => {
      if (viewMode !== 'visual') return sourceTextareaRef.current;
      return editor.isDestroyed ? null : editor.view.dom;
    }, editorFocusOrigin);
  }, [
    autoFocus,
    collaborationReady,
    editor,
    editorFocusOrigin,
    readOnly,
    viewMode,
  ]);

  if (!editor || !collaborationReady) {
    return <WorkEditorLoadingState title="正在准备 Markdown 编辑器" />;
  }

  if (readOnly) {
    return (
      <section
        ref={markdownRootRef}
        className="work-markdown-editor preview"
        style={editorStyle}
      >
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
    <section
      ref={markdownRootRef}
      className="work-markdown-editor"
      style={editorStyle}
    >
      <MarkdownToolbar
        editor={editor}
        fileActions={fileActions}
        collaborative={collaborative}
        sourceEditing={viewMode !== 'visual'}
        canSourceRedo={canRedoMarkdown}
        canSourceUndo={canUndoMarkdown}
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
