import type { Extensions } from '@tiptap/core';
import Placeholder from '@tiptap/extension-placeholder';
import { EditorContent, useEditor } from '@tiptap/react';
import {
  type CSSProperties,
  type FocusEvent,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { useDialogFocusScope } from '../../../design-system/primitives/overlay/dialog-focus-scope';
import {
  isWorkspaceContextMenuKeyboardEvent,
  type WorkspaceContextMenuEvent,
  workspaceContextMenuPosition,
} from '../../workspace/components/workspace-context-menu';
import { WorkEditorLoadingState } from '../components/work-editor-loading-state';
import type { WorkEditorAgentRequest } from '../work-agent-request';
import {
  collectDocumentChanges,
  type WorkDocumentChangeKind,
} from '../work-document-changes';
import { documentCitationCount } from '../work-document-citation-nodes';
import {
  collectDocumentCommentAnchors,
  retainAnchoredDocumentComments,
} from '../work-document-comments';
import { createWorkDocumentExtensions } from '../work-document-extensions';
import type { WorkDocumentLayoutFont } from '../work-document-fonts';
import { documentMargins, millimetersToPixels } from '../work-document-layout';
import {
  createWorkDocumentModel,
  resolveWorkDocumentEditorInput,
} from '../work-document-model';
import { normalizeDocumentPageChrome } from '../work-document-page-chrome';
import {
  documentPageColor,
  normalizeDocumentPageColor,
} from '../work-document-page-color';
import {
  DocumentPagination,
  documentPageMetrics,
  documentPaginationSurfaceHeight,
} from '../work-document-pagination';
import { documentParagraphIndent } from '../work-document-paragraph-formatting';
import {
  documentInitialSectionLayout,
  normalizeDocumentHtml,
  syncDocumentContentFromHtml,
} from '../work-document-section';
import { activeDocumentSection } from '../work-document-section-editor';
import {
  createWorkDocumentSelectionSnapshot,
  documentPlainTextAsHtml,
  type WorkGetDocumentSelectionMenuItems,
} from '../work-document-selection-menu';
import { documentParagraphTabStops } from '../work-document-tab-stops';
import { createWorkId } from '../work-templates';
import type { WorkDocumentContent, WorkDocumentNode } from '../work-types';
import { DocumentChangesPanel } from './document-changes-panel';
import { DocumentCitationsPanel } from './document-citations-panel';
import { DocumentCommentsPanel } from './document-comments-panel';
import { restoreDocumentEditorFocus } from './document-editor-focus';
import { fallbackPaginationPageDescriptor } from './document-editor-pagination';
import {
  documentCurrentPage,
  documentPageCount,
  documentTextStatistics,
} from './document-editor-support';
import {
  type DocumentFindReplaceMode,
  DocumentFindReplacePanel,
} from './document-find-replace-panel';
import {
  DocumentLayoutPanel,
  type DocumentLayoutPanelTab,
} from './document-layout-panel';
import { DocumentNavigationPanel } from './document-navigation-panel';
import { DocumentPageChromeRichTextEditor } from './document-page-chrome-editor';
import { DocumentPageStack } from './document-page-stack';
import { DocumentRuler } from './document-ruler';
import {
  DocumentSelectionContextMenu,
  type DocumentSelectionMenuState,
} from './document-selection-context-menu';
import { DocumentSelectionToolbar } from './document-selection-toolbar';
import { DocumentStatisticsDialog } from './document-statistics-dialog';
import { DocumentStatusBar } from './document-status-bar';
import { DocumentToolbar, type DocumentViewMode } from './document-toolbar';
import { DocumentVerticalRuler } from './document-vertical-ruler';
import {
  clampDocumentZoom,
  type DocumentZoomFit,
  documentZoomForFit,
} from './document-zoom';
import { OfficeFileInput, useOfficeDialog } from './office-controls';
import {
  useOfficeTaskPaneEscape,
  useOfficeTaskPaneModal,
} from './office-task-pane';
import { mergeOfficeTiptapExtensions } from './office-tiptap-extensions';
import { useDocumentComments } from './use-document-comments';
import { useDocumentInsertCommands } from './use-document-insert-commands';
import { useDocumentLayoutFonts } from './use-document-layout-fonts';
import { useDocumentPageChrome } from './use-document-page-chrome';
import { useDocumentPagination } from './use-document-pagination';
import { useOfficeEditorWheelZoom } from './use-office-editor-wheel-zoom';
import {
  type WorkOfficeFileAction,
  WorkOfficePreviewBar,
  WorkOfficeStatusBar,
} from './work-office-chrome';

export interface DocumentEditorProps {
  artifactId?: string;
  content: WorkDocumentContent;
  extensions?: Extensions;
  preview: boolean;
  saveStatus?: string;
  kernelWasmUrl?: string;
  layoutFonts?: readonly WorkDocumentLayoutFont[];
  fileActions?: readonly WorkOfficeFileAction[];
  getSelectionMenuItems?: WorkGetDocumentSelectionMenuItems;
  onChange: (content: WorkDocumentContent) => void;
  onAgentRequest?: (request: WorkEditorAgentRequest) => void | Promise<void>;
}

const EMPTY_DOCUMENT_LAYOUT_FONTS: readonly WorkDocumentLayoutFont[] = [];
const EMPTY_DOCUMENT_EXTENSIONS: Extensions = [];

type DocumentTaskPane =
  | DocumentFindReplaceMode
  | 'citations'
  | 'changes'
  | 'layout'
  | 'navigation';

function createTrackedDocumentChange(_kind: WorkDocumentChangeKind) {
  return {
    id: createWorkId('change'),
    author: 'A3S Work 用户',
    date: new Date().toISOString(),
  };
}

export function DocumentEditor({
  artifactId,
  content,
  extensions = EMPTY_DOCUMENT_EXTENSIONS,
  preview,
  saveStatus = '已自动保存',
  kernelWasmUrl,
  layoutFonts = EMPTY_DOCUMENT_LAYOUT_FONTS,
  fileActions,
  getSelectionMenuItems,
  onChange,
  onAgentRequest,
}: DocumentEditorProps) {
  const imageInputRef = useRef<HTMLInputElement>(null);
  const pageHeaderRef = useRef<HTMLElement>(null);
  const pageFooterRef = useRef<HTMLElement>(null);
  const workspaceRef = useRef<HTMLDivElement>(null);
  const reviewSurfaceRef = useRef<HTMLDivElement>(null);
  const pageSurfaceRef = useRef<HTMLElement>(null);
  const taskPaneInvokerRef = useRef<HTMLElement | null>(null);
  const commentsDraftFocusRef = useRef<HTMLElement | null>(null);
  const citationsDraftFocusRef = useRef<HTMLElement | null>(null);
  const statisticsInvokerRef = useRef<HTMLElement | null>(null);
  const contentRef = useRef(content);
  const onChangeRef = useRef(onChange);
  const trackChangesRef = useRef(Boolean(content.trackChanges));
  const normalizedContent = useMemo(
    () => normalizeDocumentHtml(content),
    [content],
  );
  const editorInput = useMemo(
    () => resolveWorkDocumentEditorInput(content, normalizedContent),
    [content, normalizedContent],
  );
  const initialContentRef = useRef(editorInput.source);
  const appliedSourceKeyRef = useRef(editorInput.sourceKey);
  const [taskPane, setTaskPane] = useState<DocumentTaskPane | null>(null);
  const [layoutPanelTab, setLayoutPanelTab] =
    useState<DocumentLayoutPanelTab>('page');
  const [findReplaceFocusRequest, setFindReplaceFocusRequest] = useState(0);
  const [citationsDirty, setCitationsDirty] = useState(false);
  const [commentsDirty, setCommentsDirty] = useState(false);
  const taskPaneDialog = useOfficeDialog();
  const taskPaneModal = useOfficeTaskPaneModal();
  const layoutOpen = taskPane === 'layout';
  const navigationOpen = taskPane === 'navigation';
  const changesOpen = taskPane === 'changes';
  const citationsOpen = taskPane === 'citations';
  const findReplaceMode =
    taskPane === 'find' || taskPane === 'replace' ? taskPane : null;
  const [spellcheckEnabled, setSpellcheckEnabled] = useState(true);
  const [viewMode, setViewMode] = useState<DocumentViewMode>('page');
  const [showRulers, setShowRulers] = useState(false);
  const [zoom, setZoom] = useState(90);
  const [selectionMenu, setSelectionMenu] =
    useState<DocumentSelectionMenuState | null>(null);
  const [selectionVersion, setSelectionVersion] = useState(0);
  const [statisticsOpen, setStatisticsOpen] = useState(false);
  const loadedLayoutFontIds = useDocumentLayoutFonts(layoutFonts);
  contentRef.current = content;
  onChangeRef.current = onChange;
  trackChangesRef.current = Boolean(content.trackChanges);
  const editorExtensions = useMemo(
    () =>
      mergeOfficeTiptapExtensions(
        'DocumentEditor',
        [
          ...createWorkDocumentExtensions({
            getContent: () => contentRef.current,
            isTracking: () => trackChangesRef.current,
            createChange: createTrackedDocumentChange,
            onContentChange: (next) => {
              contentRef.current = next;
              onChangeRef.current(next);
            },
            onTrackingChange: (trackChanges) => {
              trackChangesRef.current = trackChanges;
              const next = { ...contentRef.current, trackChanges };
              contentRef.current = next;
              onChangeRef.current(next);
            },
          }),
          Placeholder.configure({ placeholder: '在这里开始输入…' }),
          DocumentPagination,
        ],
        extensions,
      ),
    [extensions],
  );
  const editorProps = useMemo(
    () => ({
      attributes: {
        'aria-label': '文档正文',
        'aria-multiline': 'true',
        role: 'textbox',
        spellcheck: 'true',
      },
    }),
    [],
  );
  const editor = useEditor({
    extensions: editorExtensions,
    content: initialContentRef.current,
    editable: !preview,
    editorProps,
    onUpdate: ({ editor: current }) => {
      const anchors = collectDocumentCommentAnchors(current.state.doc);
      const synchronized = syncDocumentContentFromHtml(
        contentRef.current,
        current.getHTML(),
      );
      const model = createWorkDocumentModel(
        synchronized.html,
        current.getJSON() as unknown as WorkDocumentNode,
        contentRef.current.model,
      );
      const next: WorkDocumentContent = {
        ...synchronized,
        model,
        comments: retainAnchoredDocumentComments(
          contentRef.current.comments ?? [],
          anchors,
        ),
      };
      appliedSourceKeyRef.current = `model:${model.revision}:${model.htmlFingerprint}`;
      contentRef.current = next;
      onChange(next);
    },
    onSelectionUpdate: () => setSelectionVersion((value) => value + 1),
  });
  const documentComments = useDocumentComments({
    contentRef,
    editor,
    onBeforeDraft: () => setTaskPane(null),
  });
  const rememberTaskPaneInvoker = useCallback(() => {
    const active = document.activeElement;
    if (active instanceof HTMLElement && active.isConnected) {
      taskPaneInvokerRef.current = active;
    }
  }, []);
  const taskPaneRestoreTarget = useCallback(() => {
    const invoker = taskPaneInvokerRef.current;
    if (
      invoker?.isConnected &&
      !invoker.matches(':disabled') &&
      !invoker.closest('[hidden], [aria-hidden="true"]')
    ) {
      return invoker;
    }
    return editor && !editor.isDestroyed ? editor.view.dom : null;
  }, [editor]);
  const rememberTaskPaneDraftFocus = useCallback(
    (event: FocusEvent<HTMLDivElement>) => {
      const target = event.target;
      if (
        !(target instanceof HTMLElement) ||
        !target.matches(
          'input:not([type="hidden"]), textarea, [contenteditable="true"], [role="combobox"]',
        )
      ) {
        return;
      }
      if (target.closest('.work-document-comments-panel')) {
        commentsDraftFocusRef.current = target;
      } else if (target.closest('.work-document-citations-panel')) {
        citationsDraftFocusRef.current = target;
      }
    },
    [],
  );
  const restoreTaskPaneDraftFocus = useCallback(
    (pane: 'comments' | 'citations') => {
      requestAnimationFrame(() => {
        const remembered =
          pane === 'comments'
            ? commentsDraftFocusRef.current
            : citationsDraftFocusRef.current;
        const fallbackSelector =
          pane === 'comments'
            ? '.work-document-comments-panel textarea[aria-label="批注内容"], .work-document-comments-panel textarea[aria-label^="回复批注 "]'
            : '.work-document-citations-panel input:not([type="hidden"]), .work-document-citations-panel textarea, .work-document-citations-panel [role="combobox"]';
        const target =
          (remembered?.isConnected ? remembered : null) ??
          workspaceRef.current?.querySelector<HTMLElement>(fallbackSelector);
        if (target) {
          target.focus({ preventScroll: true });
        } else if (editor && !editor.isDestroyed) {
          editor.view.dom.focus({ preventScroll: true });
        }
      });
    },
    [editor],
  );
  const requestEditorViewChange = useCallback(
    async (
      nextPane: DocumentTaskPane | null,
      closeComments: boolean,
    ): Promise<boolean> => {
      if (
        taskPane === 'citations' &&
        nextPane !== 'citations' &&
        citationsDirty
      ) {
        const discard = await taskPaneDialog.confirm({
          title: '放弃未保存的文献更改？',
          description: '文献库中尚未保存的修改不会保留。',
          confirmLabel: '放弃更改',
          confirmTone: 'danger',
        });
        if (!discard) {
          restoreTaskPaneDraftFocus('citations');
          return false;
        }
      }
      if (closeComments && commentsDirty) {
        const discard = await taskPaneDialog.confirm({
          title: '放弃未完成的批注？',
          description: documentComments.draft
            ? '尚未添加的批注内容不会保留。'
            : '尚未发送的回复不会保留。',
          confirmLabel: '放弃内容',
          confirmTone: 'danger',
        });
        if (!discard) {
          restoreTaskPaneDraftFocus('comments');
          return false;
        }
      }
      if (taskPane === 'citations' && nextPane !== 'citations')
        setCitationsDirty(false);
      if (closeComments) {
        setCommentsDirty(false);
        documentComments.closeDraft(false);
        documentComments.setOpen(false);
      }
      setTaskPane(nextPane);
      return true;
    },
    [
      citationsDirty,
      commentsDirty,
      documentComments.closeDraft,
      documentComments.draft,
      documentComments.setOpen,
      restoreTaskPaneDraftFocus,
      taskPane,
      taskPaneDialog.confirm,
    ],
  );
  const restoreClosedTaskPaneFocus = useCallback(() => {
    requestAnimationFrame(() => {
      const target = taskPaneModal
        ? taskPaneRestoreTarget()
        : editor && !editor.isDestroyed
          ? editor.view.dom
          : null;
      target?.focus({ preventScroll: true });
    });
  }, [editor, taskPaneModal, taskPaneRestoreTarget]);
  const closeTaskPane = useCallback(async () => {
    if (!(await requestEditorViewChange(null, false))) return;
    restoreClosedTaskPaneFocus();
  }, [requestEditorViewChange, restoreClosedTaskPaneFocus]);
  const closeCommentsPanel = useCallback(async () => {
    if (!(await requestEditorViewChange(null, true))) return;
    restoreClosedTaskPaneFocus();
  }, [requestEditorViewChange, restoreClosedTaskPaneFocus]);
  const toggleTaskPane = useCallback(
    async (pane: Exclude<DocumentTaskPane, DocumentFindReplaceMode>) => {
      if (taskPane !== pane) rememberTaskPaneInvoker();
      return requestEditorViewChange(taskPane === pane ? null : pane, true);
    },
    [rememberTaskPaneInvoker, requestEditorViewChange, taskPane],
  );
  const openFindReplace = useCallback(
    async (mode: DocumentFindReplaceMode) => {
      if (taskPane !== mode) rememberTaskPaneInvoker();
      const accepted = await requestEditorViewChange(mode, true);
      if (accepted) setFindReplaceFocusRequest((current) => current + 1);
      return accepted;
    },
    [rememberTaskPaneInvoker, requestEditorViewChange, taskPane],
  );
  const startCommentDraft = useCallback(async () => {
    rememberTaskPaneInvoker();
    if (!(await requestEditorViewChange(null, false))) return;
    documentComments.startDraft();
  }, [
    documentComments.startDraft,
    rememberTaskPaneInvoker,
    requestEditorViewChange,
  ]);
  const toggleCommentsPanel = useCallback(async () => {
    if (documentComments.open) {
      await closeCommentsPanel();
      return;
    }
    rememberTaskPaneInvoker();
    if (!(await requestEditorViewChange(null, false))) return;
    documentComments.setOpen(true);
  }, [
    closeCommentsPanel,
    documentComments.open,
    documentComments.setOpen,
    rememberTaskPaneInvoker,
    requestEditorViewChange,
  ]);

  const activeTaskPane = Boolean(taskPane || documentComments.open);
  const activeTaskPaneElement = () =>
    workspaceRef.current?.querySelector<HTMLElement>(
      documentComments.open
        ? '.work-document-comments-panel'
        : '.work-document-task-pane',
    ) ?? null;
  useDialogFocusScope<HTMLElement>({
    active: !preview && taskPaneModal && activeTaskPane,
    onEscape: () => {
      if (documentComments.open) void closeCommentsPanel();
      else void closeTaskPane();
    },
    escapeDisabled: Boolean(documentComments.draft),
    passThroughCommandKeys: ['enter', 'f', 'h'],
    initialFocus: () => {
      const pane = activeTaskPaneElement();
      if (documentComments.draft) {
        return (
          pane?.querySelector<HTMLElement>('[aria-label="批注内容"]') ?? null
        );
      }
      if (taskPane === 'find' || taskPane === 'replace') {
        return (
          pane?.querySelector<HTMLElement>('[aria-label="查找内容"]') ?? null
        );
      }
      if (taskPane === 'navigation') {
        return (
          pane?.querySelector<HTMLElement>('[aria-label="搜索文档"]') ?? null
        );
      }
      return pane?.querySelector<HTMLElement>('.ds-icon-button.close') ?? null;
    },
    getActiveScope: activeTaskPaneElement,
    restoreFocusTarget: taskPaneRestoreTarget,
  });

  useOfficeTaskPaneEscape(Boolean(taskPane) && !taskPaneModal, closeTaskPane);
  useOfficeTaskPaneEscape(
    documentComments.open && !documentComments.draft && !taskPaneModal,
    closeCommentsPanel,
  );

  const restoreDocumentBodyFocus = useCallback(() => {
    restoreDocumentEditorFocus(() => {
      if (!editor || editor.isDestroyed) return null;
      return editor.view.dom;
    });
  }, [editor]);

  const replaceDocumentText = useCallback(
    (from: number, to: number, replacement: string) => {
      if (!editor) return false;
      if (trackChangesRef.current) {
        return editor.commands.replaceDocumentTextWithTrackedChange(
          from,
          to,
          replacement,
        );
      }
      return editor
        .chain()
        .setTextSelection({ from, to })
        .insertContent(documentPlainTextAsHtml(replacement))
        .run();
    },
    [editor],
  );

  useEffect(() => {
    if (!editor) return;
    const applyEditableState = () => {
      if (editor.isDestroyed) return;
      editor.setEditable(!preview);
      const editorDom = editor.view.dom;
      editorDom.setAttribute('role', preview ? 'document' : 'textbox');
      if (preview) {
        editorDom.removeAttribute('aria-multiline');
      } else {
        editorDom.setAttribute('aria-multiline', 'true');
      }
    };
    applyEditableState();
    editor.on('mount', applyEditableState);
    return () => {
      editor.off('mount', applyEditableState);
    };
  }, [editor, preview]);

  useEffect(() => {
    if (!editor) return;
    const applySpellcheckState = () => {
      if (!editor.isDestroyed) {
        editor.view.dom.setAttribute('spellcheck', String(spellcheckEnabled));
      }
    };
    applySpellcheckState();
    editor.on('mount', applySpellcheckState);
    return () => {
      editor.off('mount', applySpellcheckState);
    };
  }, [editor, spellcheckEnabled]);

  useEffect(() => {
    if (!editor || appliedSourceKeyRef.current === editorInput.sourceKey)
      return;
    const currentContent = normalizeDocumentHtml({
      ...content,
      html: editor.getHTML(),
    });
    if (
      typeof editorInput.source === 'string' &&
      currentContent === normalizedContent
    ) {
      appliedSourceKeyRef.current = editorInput.sourceKey;
      return;
    }
    appliedSourceKeyRef.current = editorInput.sourceKey;
    editor
      .chain()
      .setMeta('addToHistory', false)
      .setContent(editorInput.source, { emitUpdate: false })
      .run();
  }, [content, editor, editorInput, normalizedContent]);

  const section = editor ? activeDocumentSection(editor) : null;
  const layout = section?.layout ?? documentInitialSectionLayout(content);
  const margins = documentMargins({
    ...content,
    pageSize: layout.pageSize,
    margins: layout.margins,
  });
  const marginPixels = {
    top: millimetersToPixels(margins.top),
    right: millimetersToPixels(margins.right),
    bottom: millimetersToPixels(margins.bottom),
    left: millimetersToPixels(margins.left),
  };
  const pageChrome = normalizeDocumentPageChrome(layout.pageChrome, layout);
  const kernelPage = useMemo(
    () => documentPageMetrics(layout),
    [
      layout.margins.bottom,
      layout.margins.left,
      layout.margins.right,
      layout.margins.top,
      layout.orientation,
      layout.pageSize,
    ],
  );
  const pagination = useDocumentPagination({
    editor,
    documentRevision: editorInput.revision,
    enabled: Boolean(editor && viewMode === 'page'),
    layoutKey: [
      layout.breakAfter,
      layout.columns.count,
      layout.columns.spacing,
      layout.columns.separator,
      layoutOpen,
      JSON.stringify(pageChrome),
    ].join(':'),
    page: kernelPage,
    selectionVersion,
    wasmUrl: kernelWasmUrl,
    layoutFonts,
    loadedLayoutFontIds,
  });
  const documentInsert = useDocumentInsertCommands({
    contentRef,
    editor,
    resolveFieldContext: pagination.resolveFieldContext,
  });
  useEffect(() => {
    if (!editor || !pagination.resolveFieldContext) return;
    editor.commands.refreshDocumentFields(contentRef.current, {
      resolveContext: pagination.resolveFieldContext,
      addToHistory: false,
      updateClock: false,
    });
  }, [contentRef, editor, pagination.resolveFieldContext]);

  const pageCount = editor
    ? (pagination.pageCount ?? documentPageCount(editor))
    : 1;
  const paginationSurfaceHeight = pagination.pageCount
    ? documentPaginationSurfaceHeight(pagination.pageCount, kernelPage)
    : undefined;
  const currentPage = editor
    ? Math.min(pageCount, pagination.currentPage ?? documentCurrentPage(editor))
    : 1;
  const currentPageDescriptor =
    pagination.currentPageDescriptor ??
    fallbackPaginationPageDescriptor(
      section?.id,
      section?.index,
      layout,
      currentPage,
    );
  const firstPageDescriptor =
    pagination.pages[0] ??
    fallbackPaginationPageDescriptor(section?.id, section?.index, layout, 1);
  const lastPageDescriptor = pagination.pages.at(-1) ?? currentPageDescriptor;
  const {
    chromeEditor: pageChromeEditor,
    close: closePageChrome,
    edit: editPageChrome,
    editing: pageChromeEditing,
    footerChrome,
    headerChrome,
    reset: resetPageChrome,
    setChromeEditor: setPageChromeEditor,
    togglePageNumber: toggleVisiblePageNumber,
    update: updateVisiblePageChrome,
    visibleChrome,
  } = useDocumentPageChrome({
    editor,
    enabled: !preview && viewMode === 'page',
    firstPage: firstPageDescriptor,
    footerRef: pageFooterRef,
    headerRef: pageHeaderRef,
    lastPage: lastPageDescriptor,
    layout,
    onBeforeEdit: () => requestEditorViewChange(null, true),
    visiblePage: currentPageDescriptor,
  });
  useOfficeEditorWheelZoom({
    enabled: Boolean(editor),
    scopeRef: workspaceRef,
    onZoomIn: () => setZoom((current) => clampDocumentZoom(current + 10)),
    onZoomOut: () => setZoom((current) => clampDocumentZoom(current - 10)),
  });

  if (!editor) {
    return <WorkEditorLoadingState title="正在准备文字编辑器" />;
  }

  const finalPageNumber =
    pagination.pages.at(-1)?.pageNumber ??
    Math.max(1, layout.pageNumberStart ?? 1) + pageCount - 1;
  const changes = collectDocumentChanges(editor.state.doc);
  const citationCount = documentCitationCount(editor);
  const textStatistics = documentTextStatistics(editor);
  const paragraphIndent = documentParagraphIndent(editor);
  const paragraphTabStops = documentParagraphTabStops(editor);
  const updateLayout = (next: typeof layout) => {
    editor.commands.updateActiveDocumentSection(next);
  };
  const openDocumentStatistics = () => {
    const activeElement = document.activeElement;
    statisticsInvokerRef.current =
      activeElement instanceof HTMLElement && activeElement.isConnected
        ? activeElement
        : editor.view.dom;
    setStatisticsOpen(true);
  };
  const updateToolbarLayout = (next: typeof layout) => {
    updateLayout(next);
    restoreDocumentBodyFocus();
  };
  const openLayoutPanel = (target: DocumentLayoutPanelTab) => {
    setLayoutPanelTab(target);
    if (layoutOpen) return;
    void toggleTaskPane('layout').then((accepted) => {
      if (accepted) resetPageChrome();
    });
  };
  const toggleLayoutPanel = () => {
    setLayoutPanelTab('page');
    void toggleTaskPane('layout').then((accepted) => {
      if (accepted) resetPageChrome();
    });
  };
  const addSection = () => {
    editor.commands.insertDocumentSection(layout.breakAfter);
    restoreDocumentBodyFocus();
  };
  const updatePageColor = (value: string) => {
    const pageColor = normalizeDocumentPageColor(value);
    if (!pageColor) return;
    if (pageColor !== contentRef.current.pageColor) {
      const next = { ...contentRef.current, pageColor };
      contentRef.current = next;
      onChangeRef.current(next);
    }
    restoreDocumentBodyFocus();
  };
  const changeViewMode = (nextViewMode: DocumentViewMode) => {
    setViewMode(nextViewMode);
    restoreDocumentBodyFocus();
  };
  const changeToolbarZoom = (nextZoom: number) => {
    setZoom(clampDocumentZoom(nextZoom));
    restoreDocumentBodyFocus();
  };
  const fitToolbarZoom = (fit: DocumentZoomFit) => {
    setViewMode('page');
    requestAnimationFrame(() => {
      const viewport = workspaceRef.current?.querySelector<HTMLElement>(
        '.work-document-scroll',
      );
      if (!viewport) return;
      const style = getComputedStyle(viewport);
      setZoom(
        documentZoomForFit(fit, {
          pageHeight: kernelPage.height,
          pageWidth: kernelPage.width,
          viewportHeight: viewport.clientHeight,
          viewportWidth: viewport.clientWidth,
          viewportPadding: {
            top: cssPixelValue(style.paddingTop),
            right: cssPixelValue(style.paddingRight),
            bottom: cssPixelValue(style.paddingBottom),
            left: cssPixelValue(style.paddingLeft),
          },
        }),
      );
      restoreDocumentBodyFocus();
    });
  };
  const changeSpellcheck = (enabled: boolean) => {
    setSpellcheckEnabled(enabled);
    restoreDocumentBodyFocus();
  };
  const toggleBodyPageNumbers = () => {
    toggleVisiblePageNumber();
    restoreDocumentBodyFocus();
  };
  const refreshDocumentFields = () => {
    documentInsert.refreshFields();
    restoreDocumentBodyFocus();
  };
  const openSelectionContextMenu = (
    event: WorkspaceContextMenuEvent,
  ): boolean => {
    if (preview || (!getSelectionMenuItems && !onAgentRequest)) return false;
    const snapshot = createWorkDocumentSelectionSnapshot(
      editor,
      contentRef.current,
    );
    if (!snapshot) return false;
    event.preventDefault();
    event.stopPropagation();
    const position = workspaceContextMenuPosition(event);
    if (getSelectionMenuItems) {
      const items = getSelectionMenuItems(snapshot);
      setSelectionMenu(
        items.length
          ? {
              kind: 'custom',
              x: position.x,
              y: position.y,
              snapshot,
              items,
            }
          : null,
      );
      return items.length > 0;
    }
    setSelectionMenu({
      kind: 'agent',
      x: position.x,
      y: position.y,
      snapshot,
    });
    return true;
  };

  return (
    <section
      className={`work-document-editor${preview ? ' preview' : ''}`}
      data-work-pdf-artifact={artifactId}
      data-work-pdf-surface={artifactId ? 'live' : undefined}
    >
      {!preview && (
        <OfficeFileInput
          ref={imageInputRef}
          accept="image/bmp,image/gif,image/jpeg,image/png,image/webp"
          aria-label="插入文档图片"
          onFileSelect={documentInsert.insertImage}
        />
      )}
      {preview ? (
        <WorkOfficePreviewBar
          ariaLabel="文字预览工具"
          label="只读预览"
          detail={`${pageCount} 页`}
          fileActions={fileActions}
          className="work-document-ribbon"
        />
      ) : (
        <DocumentToolbar
          editor={editor}
          fileActions={fileActions}
          layout={layout}
          layoutFonts={layoutFonts}
          layoutOpen={layoutOpen}
          navigationOpen={navigationOpen}
          pageColor={documentPageColor(content.pageColor)}
          showPageNumbers={visibleChrome.showPageNumber}
          showRulers={showRulers}
          spellcheckEnabled={spellcheckEnabled}
          viewMode={viewMode}
          zoom={zoom}
          pageChromeEditor={pageChromeEditor}
          pageChromeEditingPart={pageChromeEditing?.part ?? null}
          pageChromeShowPageNumber={visibleChrome.showPageNumber}
          onRequestImage={() => imageInputRef.current?.click()}
          onPageChromeEditingPartChange={editPageChrome}
          onClosePageChrome={closePageChrome}
          onTogglePageChromePageNumber={toggleVisiblePageNumber}
          onToggleLayout={toggleLayoutPanel}
          onLayoutChange={updateToolbarLayout}
          onOpenLayout={openLayoutPanel}
          onToggleNavigation={() => void toggleTaskPane('navigation')}
          onToggleRulers={() => {
            setShowRulers((value) => !value);
            restoreDocumentBodyFocus();
          }}
          onPageColorChange={updatePageColor}
          onToggleSpellcheck={() => changeSpellcheck(!spellcheckEnabled)}
          onViewModeChange={changeViewMode}
          onZoomChange={changeToolbarZoom}
          onZoomFit={fitToolbarZoom}
          onTogglePageNumbers={toggleBodyPageNumbers}
          onInsertSection={addSection}
          onInsertNote={documentInsert.insertNote}
          onInsertCaption={documentInsert.insertCaption}
          onInsertCrossReference={documentInsert.insertCrossReference}
          citationsOpen={citationsOpen}
          citationSourceCount={content.bibliography?.sources.length ?? 0}
          onToggleCitations={() => void toggleTaskPane('citations')}
          onInsertField={documentInsert.insertField}
          onRefreshFields={refreshDocumentFields}
          canInsertComment={documentComments.canInsert}
          onInsertComment={() => void startCommentDraft()}
          commentsOpen={documentComments.open}
          commentCount={documentComments.comments.length}
          onToggleComments={() => void toggleCommentsPanel()}
          trackChanges={Boolean(content.trackChanges)}
          changesOpen={changesOpen}
          changeCount={changes.length}
          findReplaceMode={findReplaceMode}
          onRibbonTabChange={(tab) => {
            const keepCurrentPane =
              taskPane === 'navigation' ||
              (taskPane === 'layout' && tab === 'page') ||
              (taskPane === 'citations' && tab === 'references') ||
              (taskPane === 'changes' && tab === 'review') ||
              ((taskPane === 'find' || taskPane === 'replace') &&
                tab === 'home');
            return requestEditorViewChange(
              keepCurrentPane ? taskPane : null,
              tab !== 'review',
            );
          }}
          onToggleTrackChanges={() => {
            editor.commands.toggleDocumentTrackChanges();
            restoreDocumentBodyFocus();
          }}
          onToggleChanges={() => void toggleTaskPane('changes')}
          onOpenWordCount={openDocumentStatistics}
          onOpenFindReplace={openFindReplace}
        />
      )}
      <div
        ref={workspaceRef}
        className={`work-document-workspace${!preview && taskPane ? ' task-pane-open' : ''}`}
        onFocusCapture={rememberTaskPaneDraftFocus}
      >
        {!preview && navigationOpen && (
          <DocumentNavigationPanel
            currentPage={currentPage}
            editor={editor}
            modal={taskPaneModal}
            pageThumbnailSource={
              pageSurfaceRef.current && pagination.pageCount
                ? {
                    element: pageSurfaceRef.current,
                    pageCount: pagination.pageCount,
                    pageGap: kernelPage.pageGap,
                    pageHeight: kernelPage.height,
                    pageWidth: kernelPage.width,
                    revision: editorInput.sourceKey,
                  }
                : undefined
            }
            pages={(pagination.pages.length
              ? pagination.pages
              : [firstPageDescriptor]
            ).map((page) => ({
              backgroundColor: documentPageColor(content.pageColor),
              orientation: page.layout.orientation,
              pageNumber: page.pageNumber,
              physicalPage: page.physicalPage,
              previewText:
                page.previewText ||
                editor.getText().replaceAll(/\s+/g, ' ').trim().slice(0, 320),
              selectionPosition: page.selectionPosition,
            }))}
            onClose={closeTaskPane}
          />
        )}
        <div
          className={`work-document-scroll ${viewMode}${
            !preview && viewMode === 'page' && showRulers
              ? ' rulers-visible'
              : ''
          }`}
        >
          <div
            ref={reviewSurfaceRef}
            className={`work-document-review-surface${!preview && documentComments.open ? ' comments-open' : ''}`}
          >
            <div
              className={`work-document-page-stage ${layout.pageSize} ${layout.orientation} ${viewMode}`}
              data-testid="document-page-stage"
              style={
                { '--work-document-zoom': String(zoom / 100) } as CSSProperties
              }
            >
              {!preview && viewMode === 'page' && showRulers && (
                <DocumentRuler
                  layout={layout}
                  paragraphIndent={paragraphIndent}
                  tabStops={paragraphTabStops}
                  onParagraphIndentChange={(nextParagraphIndent) =>
                    editor.commands.setDocumentParagraphIndent(
                      nextParagraphIndent,
                      {
                        restoreFocus: false,
                      },
                    )
                  }
                  onTabStopsChange={(nextTabStops) =>
                    editor.commands.setDocumentParagraphTabStops(nextTabStops, {
                      restoreFocus: false,
                    })
                  }
                  onLayoutChange={updateLayout}
                />
              )}
              <div className="work-document-page-frame">
                {!preview && viewMode === 'page' && showRulers && (
                  <DocumentVerticalRuler
                    layout={layout}
                    onLayoutChange={updateLayout}
                  />
                )}
                <article
                  ref={pageSurfaceRef}
                  className={`work-document-page${preview ? ' work-document-preview-page' : ''} ${layout.pageSize} ${layout.orientation}${pagination.pageCount ? ' paginated' : ''}${!preview && pageChromeEditing ? ' page-chrome-editing' : ''}`}
                  data-work-pdf-live-document={
                    artifactId && pagination.pageCount ? '' : undefined
                  }
                  data-pdf-orientation={
                    artifactId && pagination.pageCount
                      ? layout.orientation
                      : undefined
                  }
                  data-pdf-page-count={
                    artifactId ? pagination.pageCount : undefined
                  }
                  data-pdf-page-gap={
                    artifactId && pagination.pageCount
                      ? kernelPage.pageGap
                      : undefined
                  }
                  data-pdf-page-height={
                    artifactId && pagination.pageCount
                      ? kernelPage.height
                      : undefined
                  }
                  data-pdf-page-size={
                    artifactId && pagination.pageCount
                      ? layout.pageSize
                      : undefined
                  }
                  data-pdf-page-width={
                    artifactId && pagination.pageCount
                      ? kernelPage.width
                      : undefined
                  }
                  aria-label={preview ? '文字预览' : '文字页面'}
                  style={
                    {
                      padding: `${marginPixels.top}px ${marginPixels.right}px ${marginPixels.bottom}px ${marginPixels.left}px`,
                      backgroundColor: documentPageColor(content.pageColor),
                      minHeight: paginationSurfaceHeight,
                      '--work-document-page-color': documentPageColor(
                        content.pageColor,
                      ),
                      '--work-document-page-margin-left': `${marginPixels.left}px`,
                      '--work-document-page-margin-right': `${marginPixels.right}px`,
                      '--work-document-page-margin-top': `${marginPixels.top}px`,
                      '--work-document-page-margin-bottom': `${marginPixels.bottom}px`,
                    } as CSSProperties
                  }
                >
                  {viewMode === 'page' && pagination.pageCount && (
                    <DocumentPageStack
                      pageColor={documentPageColor(content.pageColor)}
                      pageCount={pagination.pageCount}
                      pageGap={kernelPage.pageGap}
                      pageHeight={kernelPage.height}
                    />
                  )}
                  {viewMode === 'page' &&
                    (!preview || Boolean(headerChrome.headerHtml)) && (
                      // biome-ignore lint/a11y/noStaticElementInteractions: Double-click mirrors desktop Office; keyboard users use the Insert-ribbon commands.
                      <header
                        ref={pageHeaderRef}
                        className={`work-document-page-header${!preview && pageChromeEditing?.part === 'header' ? ' editing' : ''}`}
                        data-document-page-chrome={
                          firstPageDescriptor.pageChrome.variant
                        }
                        onDoubleClick={
                          preview
                            ? undefined
                            : (event) => {
                                if (pageChromeEditing?.part === 'header')
                                  return;
                                event.preventDefault();
                                void editPageChrome('header');
                              }
                        }
                      >
                        {!preview && pageChromeEditing?.part === 'header' ? (
                          <DocumentPageChromeRichTextEditor
                            key={`${pageChromeEditing.sectionId}-${pageChromeEditing.variant}-header`}
                            autoFocus
                            className="work-document-page-chrome-inline-editor"
                            label="页内页眉"
                            value={headerChrome.headerHtml}
                            showToolbar={false}
                            onChange={(headerHtml) =>
                              updateVisiblePageChrome({ headerHtml })
                            }
                            onEditorChange={setPageChromeEditor}
                            onExit={closePageChrome}
                          />
                        ) : headerChrome.headerHtml ? (
                          <div
                            className="work-document-page-chrome-html"
                            dangerouslySetInnerHTML={{
                              __html: headerChrome.headerHtml,
                            }}
                          />
                        ) : null}
                      </header>
                    )}
                  <section
                    className={`work-document-editable ${viewMode}`}
                    aria-label="文档内容编辑区域"
                    onDoubleClick={() => {
                      if (!preview && pageChromeEditing) closePageChrome();
                    }}
                    onContextMenu={openSelectionContextMenu}
                    onKeyDownCapture={(event) => {
                      if (!isWorkspaceContextMenuKeyboardEvent(event)) return;
                      openSelectionContextMenu(event);
                    }}
                  >
                    <EditorContent editor={editor} />
                    {!preview && (
                      <DocumentSelectionToolbar
                        editor={editor}
                        canInsertComment={documentComments.canInsert}
                        layoutFonts={layoutFonts}
                        onInsertComment={() => void startCommentDraft()}
                      />
                    )}
                  </section>
                  {viewMode === 'page' &&
                    (!preview ||
                      Boolean(footerChrome.footerHtml) ||
                      footerChrome.showPageNumber) && (
                      // biome-ignore lint/a11y/noStaticElementInteractions: Double-click mirrors desktop Office; keyboard users use the Insert-ribbon commands.
                      <footer
                        ref={pageFooterRef}
                        className={`work-document-page-footer${!preview && pageChromeEditing?.part === 'footer' ? ' editing' : ''}`}
                        data-document-page-chrome={
                          lastPageDescriptor.pageChrome.variant
                        }
                        onDoubleClick={
                          preview
                            ? undefined
                            : (event) => {
                                if (pageChromeEditing?.part === 'footer')
                                  return;
                                event.preventDefault();
                                void editPageChrome('footer');
                              }
                        }
                      >
                        <div className="work-document-page-footer-content">
                          {!preview && pageChromeEditing?.part === 'footer' ? (
                            <DocumentPageChromeRichTextEditor
                              key={`${pageChromeEditing.sectionId}-${pageChromeEditing.variant}-footer`}
                              autoFocus
                              className="work-document-page-chrome-inline-editor"
                              label="页内页脚"
                              value={footerChrome.footerHtml}
                              showToolbar={false}
                              onChange={(footerHtml) =>
                                updateVisiblePageChrome({ footerHtml })
                              }
                              onEditorChange={setPageChromeEditor}
                              onExit={closePageChrome}
                            />
                          ) : footerChrome.footerHtml ? (
                            <div
                              className="work-document-page-chrome-html"
                              dangerouslySetInnerHTML={{
                                __html: footerChrome.footerHtml,
                              }}
                            />
                          ) : null}
                        </div>
                        {footerChrome.showPageNumber && (
                          <span className="work-document-page-number">
                            {lastPageDescriptor.pageNumber} / {finalPageNumber}
                          </span>
                        )}
                      </footer>
                    )}
                </article>
              </div>
            </div>
            {!preview && documentComments.open && (
              <DocumentCommentsPanel
                editor={editor}
                comments={documentComments.comments}
                draft={documentComments.draft}
                surfaceRef={reviewSurfaceRef}
                onReply={documentComments.reply}
                onToggleResolved={documentComments.toggleResolved}
                onDelete={documentComments.deleteComment}
                onCancelDraft={documentComments.closeDraft}
                onSubmitDraft={documentComments.submitDraft}
                onClose={() => void closeCommentsPanel()}
                onDirtyChange={setCommentsDirty}
              />
            )}
          </div>
        </div>
        {!preview && citationsOpen && (
          <DocumentCitationsPanel
            editor={editor}
            content={content}
            onClose={closeTaskPane}
            onDirtyChange={setCitationsDirty}
          />
        )}
        {!preview && changesOpen && (
          <DocumentChangesPanel
            editor={editor}
            changes={changes}
            trackChanges={Boolean(content.trackChanges)}
            onTrackChangesChange={(enabled) =>
              editor.commands.setDocumentTrackChanges(enabled)
            }
            onClose={closeTaskPane}
          />
        )}
        {!preview && layoutOpen && section && (
          <DocumentLayoutPanel
            layout={layout}
            activeTab={layoutPanelTab}
            sectionIndex={section.index}
            sectionCount={section.count}
            onActiveTabChange={setLayoutPanelTab}
            onChange={updateLayout}
            onInsertSection={addSection}
            onMergeSection={() =>
              editor.commands.mergeDocumentSectionWithPrevious()
            }
            onClose={closeTaskPane}
          />
        )}
        {!preview && findReplaceMode && (
          <DocumentFindReplacePanel
            editor={editor}
            mode={findReplaceMode}
            focusRequest={findReplaceFocusRequest}
            onModeChange={(mode) => void openFindReplace(mode)}
            onReplaceText={replaceDocumentText}
            onClose={closeTaskPane}
          />
        )}
      </div>
      {preview ? (
        <WorkOfficeStatusBar className="work-document-footer">
          <output aria-label="页数状态">{pageCount} 页</output>
          <output aria-label="分节状态">{section?.count ?? 1} 节</output>
        </WorkOfficeStatusBar>
      ) : (
        <DocumentStatusBar
          bibliographyCount={content.bibliography?.sources.length ?? 0}
          citationCount={citationCount}
          currentPage={currentPage}
          pageCount={pageCount}
          saveStatus={saveStatus}
          sectionCount={section?.count ?? 1}
          sectionIndex={section?.index ?? 0}
          spellcheckEnabled={spellcheckEnabled}
          viewMode={viewMode}
          wordCount={textStatistics.wordCount}
          zoom={zoom}
          onSpellcheckChange={changeSpellcheck}
          onOpenWordCount={openDocumentStatistics}
          onViewModeChange={changeViewMode}
          onZoomChange={setZoom}
        />
      )}
      {!preview && selectionMenu && (
        <DocumentSelectionContextMenu
          editor={editor}
          menu={selectionMenu}
          getTrackChanges={() => trackChangesRef.current}
          onAgentRequest={onAgentRequest}
          onClose={() => setSelectionMenu(null)}
        />
      )}
      {!preview && documentInsert.dialog}
      {!preview && taskPaneDialog.dialog}
      {!preview && statisticsOpen && (
        <DocumentStatisticsDialog
          pageCount={pageCount}
          statistics={textStatistics}
          restoreFocusTarget={() => statisticsInvokerRef.current}
          onClose={() => setStatisticsOpen(false)}
        />
      )}
    </section>
  );
}

function cssPixelValue(value: string): number {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : 0;
}
