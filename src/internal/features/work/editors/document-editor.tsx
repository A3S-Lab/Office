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
import type { WorkOfficeCollaborationSession } from '../../../collaboration/office-collaboration';
import {
  createWorkOfficeDocumentCollaborationBinding,
  readWorkOfficeDocumentCollaboration,
  type WorkOfficeDocumentCollaborationBinding,
} from '../../../collaboration/office-document-collaboration';
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
import { millimetersToPixels } from '../work-document-layout';
import {
  createWorkDocumentModel,
  resolveWorkDocumentEditorInput,
} from '../work-document-model';
import { normalizeDocumentPageChrome } from '../work-document-page-chrome';
import {
  documentPageColor,
  normalizeDocumentPageColor,
} from '../work-document-page-color';
import { documentPageSurfaceGeometry } from '../work-document-page-frames';
import { resolveDocumentPageMargins } from '../work-document-page-margins';
import { resolveDocumentPageSize } from '../work-document-page-size';
import {
  DocumentPagination,
  documentPageMetrics,
} from '../work-document-pagination';
import { documentParagraphIndent } from '../work-document-paragraph-formatting';
import type {
  WorkDocumentReviewConflict,
  WorkDocumentReviewConflictEvent,
} from '../work-document-review-conflicts';
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
import {
  DocumentReviewConflictNotice,
  useDocumentReviewConflicts,
} from './use-document-review-conflicts';
import { useOfficeEditorWheelZoom } from './use-office-editor-wheel-zoom';
import {
  type WorkOfficeFileAction,
  WorkOfficePreviewBar,
  WorkOfficeStatusBar,
} from './work-office-chrome';

export interface DocumentEditorProps {
  artifactId?: string;
  collaboration?: WorkOfficeCollaborationSession;
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
  onReviewConflict?: (event: WorkDocumentReviewConflictEvent) => void;
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

interface DocumentCollaborationBridge {
  getContent(): WorkDocumentContent | null;
  isTracking(): boolean;
  onContentChange(content: WorkDocumentContent): void;
  onTrackingChange(enabled: boolean): void;
}

interface DocumentEditorSurfaceProps extends DocumentEditorProps {
  collaborationBinding?: WorkOfficeDocumentCollaborationBinding;
  collaborationBridge?: { current: DocumentCollaborationBridge };
  collaborationInitialContent?: WorkDocumentContent;
}

export function DocumentEditor(props: DocumentEditorProps) {
  if (!props.collaboration) return <DocumentEditorSurface {...props} />;
  return (
    <CollaborativeDocumentEditor
      {...props}
      collaboration={props.collaboration}
    />
  );
}

function CollaborativeDocumentEditor(
  props: DocumentEditorProps & {
    collaboration: WorkOfficeCollaborationSession;
  },
) {
  const { collaboration, extensions = EMPTY_DOCUMENT_EXTENSIONS } = props;
  const collaborationRef = useRef(collaboration);
  const extensionsRef = useRef(extensions);
  if (collaborationRef.current !== collaboration) {
    throw new Error(
      'DocumentEditor collaboration sessions cannot be replaced while mounted. Remount the editor for another shared document.',
    );
  }
  if (extensionsRef.current !== extensions) {
    throw new Error(
      'DocumentEditor extensions cannot be replaced while a collaboration session is mounted. Remount the editor to change its schema.',
    );
  }
  const initial = useRef<WorkDocumentContent | null>(null);
  if (!initial.current) {
    initial.current = readWorkOfficeDocumentCollaboration(collaboration);
  }
  const bridge = useRef<DocumentCollaborationBridge>({
    getContent: () => initial.current,
    isTracking: () => Boolean(initial.current?.trackChanges),
    onContentChange: () => undefined,
    onTrackingChange: () => undefined,
  });
  const [binding, setBinding] =
    useState<WorkOfficeDocumentCollaborationBinding>();
  const bindingRef = useRef<WorkOfficeDocumentCollaborationBinding | undefined>(
    undefined,
  );
  const disposeTokenRef = useRef<object | undefined>(undefined);

  useEffect(() => {
    disposeTokenRef.current = undefined;
    const current =
      bindingRef.current ??
      createWorkOfficeDocumentCollaborationBinding(collaboration, {
        additionalExtensions: extensionsRef.current,
        workExtensions: {
          getContent: () => bridge.current.getContent(),
          isTracking: () => bridge.current.isTracking(),
          createChange: createTrackedDocumentChange,
          onContentChange: (content: WorkDocumentContent) =>
            bridge.current.onContentChange(content),
          onTrackingChange: (enabled: boolean) =>
            bridge.current.onTrackingChange(enabled),
        },
      });
    bindingRef.current = current;
    setBinding(current);
    return () => {
      const token = {};
      disposeTokenRef.current = token;
      queueMicrotask(() => {
        if (disposeTokenRef.current !== token) return;
        if (bindingRef.current === current) bindingRef.current = undefined;
        current.destroy();
      });
    };
  }, [collaboration]);

  if (!binding) {
    return <WorkEditorLoadingState title="正在准备协作文档" />;
  }
  return (
    <DocumentEditorSurface
      {...props}
      collaborationBinding={binding}
      collaborationBridge={bridge}
      collaborationInitialContent={initial.current}
    />
  );
}

function DocumentEditorSurface({
  artifactId,
  collaboration,
  collaborationBinding,
  collaborationBridge,
  collaborationInitialContent,
  content,
  extensions = EMPTY_DOCUMENT_EXTENSIONS,
  preview: requestedPreview,
  saveStatus = '已自动保存',
  kernelWasmUrl,
  layoutFonts = EMPTY_DOCUMENT_LAYOUT_FONTS,
  fileActions,
  getSelectionMenuItems,
  onChange,
  onAgentRequest,
  onReviewConflict,
}: DocumentEditorSurfaceProps) {
  const collaborationEditable = !collaboration || collaboration.mode === 'edit';
  const preview = requestedPreview || !collaborationEditable;
  const readOnly = preview;
  const initialContentRef = useRef(collaborationInitialContent ?? content);
  const effectiveContent = collaboration ? initialContentRef.current : content;
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
  const contentRef = useRef(effectiveContent);
  const onChangeRef = useRef(onChange);
  const trackChangesRef = useRef(Boolean(effectiveContent.trackChanges));
  const collaborationBindingRef = useRef(collaborationBinding);
  const normalizedContent = useMemo(
    () => normalizeDocumentHtml(effectiveContent),
    [effectiveContent],
  );
  const editorInput = useMemo(
    () => resolveWorkDocumentEditorInput(effectiveContent, normalizedContent),
    [effectiveContent, normalizedContent],
  );
  const initialEditorSourceRef = useRef(editorInput.source);
  const appliedSourceKeyRef = useRef(editorInput.sourceKey);
  const activeReviewConflictsRef = useRef<WorkDocumentReviewConflict[]>([]);
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
  if (!collaboration) contentRef.current = content;
  onChangeRef.current = onChange;
  trackChangesRef.current = Boolean(contentRef.current.trackChanges);
  const commitContentChange = useCallback((next: WorkDocumentContent) => {
    const previous = contentRef.current;
    const binding = collaborationBindingRef.current;
    binding?.updateSidecars(previous, next);
    contentRef.current = next;
    if (!binding) onChangeRef.current(next);
  }, []);
  collaborationBindingRef.current = collaborationBinding;
  if (collaborationBridge) {
    collaborationBridge.current = {
      getContent: () => contentRef.current,
      isTracking: () => trackChangesRef.current,
      onContentChange: (next) => commitContentChange(next),
      onTrackingChange: (trackChanges) => {
        trackChangesRef.current = trackChanges;
        commitContentChange({ ...contentRef.current, trackChanges });
      },
    };
  }
  const editorExtensions = useMemo(
    () =>
      mergeOfficeTiptapExtensions(
        'DocumentEditor',
        [
          ...(collaborationBinding?.extensions ??
            createWorkDocumentExtensions({
              getContent: () => contentRef.current,
              isTracking: () => trackChangesRef.current,
              createChange: createTrackedDocumentChange,
              onContentChange: (next) => {
                commitContentChange(next);
              },
              onTrackingChange: (trackChanges) => {
                trackChangesRef.current = trackChanges;
                const next = { ...contentRef.current, trackChanges };
                commitContentChange(next);
              },
            })),
          Placeholder.configure({ placeholder: '在这里开始输入…' }),
          DocumentPagination,
        ],
        collaboration ? EMPTY_DOCUMENT_EXTENSIONS : extensions,
      ),
    [collaboration, collaborationBinding, commitContentChange, extensions],
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
    content: collaboration ? undefined : initialEditorSourceRef.current,
    editable: !readOnly,
    editorProps,
    onUpdate: ({ editor: current }) => {
      const anchors = collectDocumentCommentAnchors(current.state.doc);
      const retainedCommentIds = new Set(
        activeReviewConflictsRef.current
          .filter((conflict) => conflict.kind === 'comment')
          .map((conflict) => conflict.id),
      );
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
          retainedCommentIds,
        ),
      };
      appliedSourceKeyRef.current = `model:${model.revision}:${model.htmlFingerprint}`;
      if (!collaborationBinding) {
        contentRef.current = next;
        onChangeRef.current(next);
      }
    },
    onSelectionUpdate: () => setSelectionVersion((value) => value + 1),
  });
  const reviewConflicts = useDocumentReviewConflicts({
    activeConflictsRef: activeReviewConflictsRef,
    appliedSourceKeyRef,
    artifactId,
    content: collaboration ? contentRef.current : content,
    editor,
    editorInput,
    normalizedContent,
    onReviewConflict,
    reconcileControlledUpdates: !collaboration,
  });
  const documentComments = useDocumentComments({
    contentRef,
    editor,
    onBeforeDraft: () => setTaskPane(null),
  });
  const [collaborationVersion, setCollaborationVersion] = useState(0);
  const handledCollaborationVersionRef = useRef(0);
  const pendingCollaborationContentRef = useRef<
    WorkDocumentContent | undefined
  >(undefined);
  useEffect(() => {
    if (!collaborationBinding) return;
    const unsubscribeChange = collaborationBinding.subscribe((change) => {
      pendingCollaborationContentRef.current = change.content;
      contentRef.current = change.content;
      trackChangesRef.current = Boolean(change.content.trackChanges);
      setCollaborationVersion((value) => value + 1);
    });
    const unsubscribeError = collaborationBinding.subscribeError((error) => {
      queueMicrotask(() => {
        throw error;
      });
    });
    const unsubscribeHistory = collaborationBinding.subscribeHistory(() => {
      setCollaborationVersion((value) => value + 1);
    });
    return () => {
      unsubscribeChange();
      unsubscribeError();
      unsubscribeHistory();
    };
  }, [collaborationBinding]);
  useEffect(() => {
    if (
      collaborationVersion === handledCollaborationVersionRef.current ||
      !pendingCollaborationContentRef.current
    ) {
      return;
    }
    handledCollaborationVersionRef.current = collaborationVersion;
    const next = pendingCollaborationContentRef.current;
    pendingCollaborationContentRef.current = undefined;
    contentRef.current = next;
    trackChangesRef.current = Boolean(next.trackChanges);
    onChangeRef.current(next);
  }, [collaborationVersion]);
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
      editor.setEditable(!readOnly);
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
  }, [editor, preview, readOnly]);

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

  const section = editor ? activeDocumentSection(editor) : null;
  const currentContent = contentRef.current;
  const layout =
    section?.layout ?? documentInitialSectionLayout(currentContent);
  const resolvedPageSize = resolveDocumentPageSize(layout);
  const resolvedMargins = resolveDocumentPageMargins(layout, 1);
  const margins = resolvedMargins.body;
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
      JSON.stringify(layout.pageGeometry),
      JSON.stringify(layout.pageMargins),
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
      JSON.stringify(layout.pageBorders),
      JSON.stringify(layout.pageMargins),
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
  const paginationGeometry = pagination.pageCount
    ? documentPageSurfaceGeometry(
        pagination.pages.map((page) => page.page),
        kernelPage,
        pagination.pageCount,
      )
    : null;
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
  const firstPageFrame = paginationGeometry?.frames[0];
  const lastPageFrame = paginationGeometry?.frames.at(-1);
  const firstPageMargins = resolveDocumentPageMargins(
    firstPageDescriptor.layout,
    firstPageDescriptor.physicalPage,
  );
  const lastPageMargins = resolveDocumentPageMargins(
    lastPageDescriptor.layout,
    lastPageDescriptor.physicalPage,
  );
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
      commitContentChange(next);
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
          pageHeight: currentPageDescriptor.page.height,
          pageWidth: currentPageDescriptor.page.width,
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
          history={
            collaborationBinding
              ? {
                  canRedo: collaborationBinding.canRedo(),
                  canUndo: collaborationBinding.canUndo(),
                  redo: () => collaborationBinding.redo(),
                  undo: () => collaborationBinding.undo(),
                }
              : undefined
          }
          fileActions={fileActions}
          layout={layout}
          layoutFonts={layoutFonts}
          layoutOpen={layoutOpen}
          navigationOpen={navigationOpen}
          pageColor={documentPageColor(currentContent.pageColor)}
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
          citationSourceCount={currentContent.bibliography?.sources.length ?? 0}
          onToggleCitations={() => void toggleTaskPane('citations')}
          onInsertField={documentInsert.insertField}
          onRefreshFields={refreshDocumentFields}
          canInsertComment={documentComments.canInsert}
          onInsertComment={() => void startCommentDraft()}
          commentsOpen={documentComments.open}
          commentCount={documentComments.comments.length}
          onToggleComments={() => void toggleCommentsPanel()}
          trackChanges={Boolean(currentContent.trackChanges)}
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
      {!preview && (
        <DocumentReviewConflictNotice
          conflicts={reviewConflicts.visibleConflicts}
          onDismiss={reviewConflicts.dismiss}
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
                    pages: paginationGeometry?.frames,
                    revision: editorInput.sourceKey,
                  }
                : undefined
            }
            pages={(pagination.pages.length
              ? pagination.pages
              : [firstPageDescriptor]
            ).map((page) => ({
              backgroundColor: documentPageColor(currentContent.pageColor),
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
                  data-pdf-page-width-points={
                    artifactId && pagination.pageCount
                      ? resolvedPageSize.widthPoints
                      : undefined
                  }
                  data-pdf-page-height-points={
                    artifactId && pagination.pageCount
                      ? resolvedPageSize.heightPoints
                      : undefined
                  }
                  data-pdf-page-width={
                    artifactId && pagination.pageCount
                      ? kernelPage.width
                      : undefined
                  }
                  data-document-page-gutter-position={
                    resolvedMargins.gutterPosition
                  }
                  data-document-page-margins-bounded={String(
                    resolvedMargins.bounded,
                  )}
                  data-document-page-mirror-margins={String(
                    resolvedMargins.mirrorMargins,
                  )}
                  data-document-page-top-margin-mode={resolvedMargins.topMode}
                  data-document-page-bottom-margin-mode={
                    resolvedMargins.bottomMode
                  }
                  aria-label={preview ? '文字预览' : '文字页面'}
                  style={
                    {
                      padding: paginationGeometry
                        ? `${firstPageDescriptor.page.marginTop}px 0 ${lastPageDescriptor.page.marginBottom}px`
                        : `${marginPixels.top}px ${marginPixels.right}px ${marginPixels.bottom}px ${marginPixels.left}px`,
                      backgroundColor: documentPageColor(
                        currentContent.pageColor,
                      ),
                      width:
                        viewMode === 'page'
                          ? (paginationGeometry?.width ?? kernelPage.width)
                          : undefined,
                      minHeight:
                        viewMode === 'page'
                          ? (paginationGeometry?.height ?? kernelPage.height)
                          : undefined,
                      '--work-document-page-color': documentPageColor(
                        currentContent.pageColor,
                      ),
                      '--work-document-page-margin-left': `${marginPixels.left}px`,
                      '--work-document-page-margin-right': `${marginPixels.right}px`,
                      '--work-document-page-margin-top': `${marginPixels.top}px`,
                      '--work-document-page-margin-bottom': `${marginPixels.bottom}px`,
                      '--work-document-page-header-distance': `${millimetersToPixels(
                        resolvedMargins.headerDistance,
                      )}px`,
                      '--work-document-page-footer-distance': `${millimetersToPixels(
                        resolvedMargins.footerDistance,
                      )}px`,
                      '--work-document-page-header-height': `${Math.max(
                        0,
                        marginPixels.top -
                          millimetersToPixels(resolvedMargins.headerDistance),
                      )}px`,
                      '--work-document-page-footer-height': `${Math.max(
                        0,
                        marginPixels.bottom -
                          millimetersToPixels(resolvedMargins.footerDistance),
                      )}px`,
                    } as CSSProperties
                  }
                >
                  {viewMode === 'page' && pagination.pageCount && (
                    <DocumentPageStack
                      pageColor={documentPageColor(currentContent.pageColor)}
                      pageCount={pagination.pageCount}
                      pageGap={kernelPage.pageGap}
                      pageHeight={kernelPage.height}
                      pageWidth={kernelPage.width}
                      pages={pagination.pages}
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
                        style={
                          firstPageFrame && paginationGeometry
                            ? {
                                height: firstPageDescriptor.page.headerHeight,
                                left:
                                  firstPageFrame.left +
                                  firstPageDescriptor.page.marginLeft,
                                right:
                                  paginationGeometry.width -
                                  firstPageFrame.left -
                                  firstPageFrame.width +
                                  firstPageDescriptor.page.marginRight,
                                top: millimetersToPixels(
                                  firstPageMargins.headerDistance,
                                ),
                              }
                            : undefined
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
                        style={
                          lastPageFrame && paginationGeometry
                            ? {
                                bottom: millimetersToPixels(
                                  lastPageMargins.footerDistance,
                                ),
                                height: lastPageDescriptor.page.footerHeight,
                                left:
                                  lastPageFrame.left +
                                  lastPageDescriptor.page.marginLeft,
                                right:
                                  paginationGeometry.width -
                                  lastPageFrame.left -
                                  lastPageFrame.width +
                                  lastPageDescriptor.page.marginRight,
                              }
                            : undefined
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
            content={currentContent}
            onClose={closeTaskPane}
            onDirtyChange={setCitationsDirty}
          />
        )}
        {!preview && changesOpen && (
          <DocumentChangesPanel
            editor={editor}
            changes={changes}
            trackChanges={Boolean(currentContent.trackChanges)}
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
          bibliographyCount={currentContent.bibliography?.sources.length ?? 0}
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
