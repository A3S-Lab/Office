import type { Editor } from '@tiptap/core';
import {
  ArrowUpRight,
  Bookmark as BookmarkIcon,
  Braces,
  Check,
  CheckCheck,
  ChevronDown,
  ChevronUp,
  Eye,
  FileDiff,
  FilePlus2,
  FileStack,
  FileText,
  Filter,
  GitCompareArrows,
  Globe2,
  Hash,
  Image as ImageIcon,
  Languages,
  Link2,
  ListChecks,
  MessageSquarePlus,
  MessagesSquare,
  PanelBottomOpen,
  PanelLeftOpen,
  PanelTopOpen,
  Redo2,
  Ruler,
  Scan,
  SlidersHorizontal,
  StretchHorizontal,
  TextCursorInput,
  TextSelect,
  Undo2,
  Users,
  XCircle,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import {
  type MouseEvent as ReactMouseEvent,
  type ReactNode,
  useCallback,
  useEffect,
  useState,
} from 'react';
import { Popover } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import {
  activeDocumentBookmark,
  DOCUMENT_BOOKMARK_DUPLICATE_MESSAGE,
  documentBookmarkNameExists,
  validateDocumentBookmarkName,
} from '../work-document-bookmarks';
import type { WorkDocumentCaptionKind } from '../work-document-captions';
import {
  collectDocumentChanges,
  type WorkDocumentChange,
} from '../work-document-changes';
import type { DocumentComparisonMode } from '../work-document-compare';
import type { WorkDocumentFieldKind } from '../work-document-fields';
import type { WorkDocumentLayoutFont } from '../work-document-fonts';
import {
  DOCUMENT_LINK_VALIDATION_MESSAGE,
  normalizeDocumentHref,
} from '../work-document-links';
import type { WorkDocumentNoteKind } from '../work-document-notes';
import type { WorkDocumentSectionLayout } from '../work-types';
import {
  type DocumentRibbonTabId,
  documentConnectorRibbonTab,
  documentPageChromeRibbonTab,
  documentPictureRibbonTab,
  documentRibbonTabs,
  documentTableRibbonTabs,
  documentTextBoxRibbonTab,
  getDocumentCommandDefinition,
} from './document-command-catalog';
import {
  documentCommandLabel,
  localizeDocumentRibbonTab,
} from './document-command-i18n';
import { DocumentConnectorRibbon } from './document-connector-ribbon';
import { synchronizeDocumentEditorSelectionFromDom } from './document-dom-selection';
import { documentHasRefreshableFields } from './document-editor-support';
import type { DocumentFindReplaceMode } from './document-find-replace-panel';
import { DocumentFontDialog } from './document-font-dialog';
import {
  applyDocumentFontDialogPatch,
  type DocumentFontDialogSource,
  documentFontDialogSource,
} from './document-font-dialog-model';
import { DocumentHomeRibbon } from './document-home-ribbon';
import type { DocumentLayoutPanelTab } from './document-layout-panel';
import {
  type DocumentPageChromeEditingPart,
  DocumentPageChromeRibbon,
} from './document-page-chrome-ribbon';
import { DocumentPageLayoutRibbon } from './document-page-layout-ribbon';
import { DocumentPictureRibbon } from './document-picture-ribbon';
import { DocumentProofingDialog } from './document-proofing-dialog';
import {
  applyDocumentProofingDialogPatch,
  type DocumentProofingDialogSource,
  documentProofingDialogSource,
} from './document-proofing-dialog-model';
import { DocumentReferencesRibbon } from './document-references-ribbon';
import {
  actionableDocumentChangeIndex,
  adjacentDocumentChangeIndex,
} from './document-review-navigation';
import { DocumentTableInsertPopover } from './document-table-insert-popover';
import {
  DocumentTableDesignRibbon,
  DocumentTableLayoutRibbon,
} from './document-table-ribbon';
import { DocumentTextBoxRibbon } from './document-text-box-ribbon';
import { runDocumentWpsShortcut } from './document-wps-shortcuts';
import {
  type DocumentZoomFit,
  MAX_DOCUMENT_ZOOM,
  MIN_DOCUMENT_ZOOM,
} from './document-zoom';
import { useOfficeDialog } from './office-controls';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import { useOfficeMessages } from './office-messages-context';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import {
  type WorkOfficeFileAction,
  WorkOfficeRibbon,
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

export type DocumentViewMode = 'page' | 'web';

interface DocumentFontDialogRequest {
  editor: Editor;
  selection: { from: number; to: number };
  source: DocumentFontDialogSource;
}

interface DocumentProofingDialogRequest {
  editor: Editor;
  selection: { from: number; to: number };
  source: DocumentProofingDialogSource;
}

interface DocumentToolbarProps {
  editor: Editor;
  defaultRibbonCollapsed?: boolean;
  reviewOnly?: boolean;
  suggestionOnly?: boolean;
  history?: {
    canRedo: boolean;
    canUndo: boolean;
    redo: () => boolean;
    undo: () => boolean;
  };
  layoutOpen: boolean;
  layout: WorkDocumentSectionLayout;
  layoutFonts?: readonly WorkDocumentLayoutFont[];
  navigationOpen: boolean;
  pageColor: string;
  showPageNumbers: boolean;
  showFieldCodes: boolean;
  showHiddenText: boolean;
  showRulers: boolean;
  spellcheckEnabled: boolean;
  viewMode: DocumentViewMode;
  zoom: number;
  pageChromeEditor: Editor | null;
  pageChromeEditingPart: DocumentPageChromeEditingPart | null;
  pageChromeShowPageNumber: boolean;
  onRequestImage: () => void;
  onInsertTextBox?: () => void;
  onInsertConnector?: () => void;
  onInsertContentControl?: () => void;
  onImportMailMergeRecipients?: () => void;
  onOpenMailMergeRecipientFilter?: () => void;
  onPageChromeEditingPartChange: (part: DocumentPageChromeEditingPart) => void;
  onClosePageChrome: () => void;
  onTogglePageChromePageNumber: () => void;
  onToggleLayout: () => void;
  onLayoutChange: (layout: WorkDocumentSectionLayout) => void;
  onOpenLayout: (target: DocumentLayoutPanelTab) => void;
  onToggleNavigation: () => void;
  onTogglePageNumbers: () => void;
  onToggleFieldCodes: () => void;
  onToggleSelectedFieldCodes: () => boolean;
  onUnlinkFields: () => boolean;
  onLockFields: () => boolean;
  onUnlockFields: () => boolean;
  onToggleHiddenText: () => void;
  onToggleRulers: () => void;
  onPageColorChange: (color: string) => void;
  onToggleSpellcheck: () => void;
  onViewModeChange: (mode: DocumentViewMode) => void;
  onZoomChange: (zoom: number) => void;
  onZoomFit: (fit: DocumentZoomFit) => void;
  onInsertSection: () => void;
  onInsertNote: (kind: WorkDocumentNoteKind) => void;
  onInsertCaption: (kind: WorkDocumentCaptionKind) => void;
  onInsertCrossReference: () => void;
  onOpenTableOfContents: () => void;
  onOpenIndexEntry: () => void;
  onOpenIndex: () => void;
  citationsOpen: boolean;
  citationSourceCount: number;
  onToggleCitations: () => void;
  onInsertField: (kind: WorkDocumentFieldKind) => void;
  onOpenField: () => void;
  onRefreshFields: () => void;
  onRefreshIndex: () => void;
  onRefreshTableOfContents: () => void;
  canInsertComment: boolean;
  onInsertComment: () => void;
  commentsOpen: boolean;
  commentCount: number;
  onToggleComments: () => void;
  canAcceptComment?: boolean;
  canProcessComment?: boolean;
  canWithdrawComment?: boolean;
  onAcceptComment?: () => void;
  onProcessComment?: () => void;
  onWithdrawComment?: () => void;
  trackChanges: boolean;
  changesOpen: boolean;
  changeCount: number;
  findReplaceMode: DocumentFindReplaceMode | null;
  fileActions?: readonly WorkOfficeFileAction[];
  onRibbonTabChange?: (
    tab: DocumentRibbonTabId,
  ) => boolean | undefined | Promise<boolean | undefined>;
  onToggleTrackChanges: () => void;
  onToggleChanges: () => void;
  onOpenComparison: (mode: DocumentComparisonMode) => void;
  onDecideChange?: (
    change: WorkDocumentChange,
    decision: 'accept' | 'reject',
  ) => boolean;
  onOpenWordCount: () => void;
  onOpenFindReplace: (mode: DocumentFindReplaceMode) => void;
}

export function DocumentToolbar({
  editor,
  defaultRibbonCollapsed = false,
  reviewOnly = false,
  suggestionOnly = false,
  history,
  layoutOpen,
  layout,
  layoutFonts = [],
  navigationOpen,
  pageColor,
  showPageNumbers,
  showFieldCodes,
  showHiddenText,
  showRulers,
  spellcheckEnabled,
  viewMode,
  zoom,
  pageChromeEditor,
  pageChromeEditingPart,
  pageChromeShowPageNumber,
  onRequestImage,
  onInsertTextBox,
  onInsertConnector,
  onInsertContentControl,
  onImportMailMergeRecipients,
  onOpenMailMergeRecipientFilter,
  onPageChromeEditingPartChange,
  onClosePageChrome,
  onTogglePageChromePageNumber,
  onToggleLayout,
  onLayoutChange,
  onOpenLayout,
  onToggleNavigation,
  onTogglePageNumbers,
  onToggleFieldCodes,
  onToggleSelectedFieldCodes,
  onUnlinkFields,
  onLockFields,
  onUnlockFields,
  onToggleHiddenText,
  onToggleRulers,
  onPageColorChange,
  onToggleSpellcheck,
  onViewModeChange,
  onZoomChange,
  onZoomFit,
  onInsertSection,
  onInsertNote,
  onInsertCaption,
  onInsertCrossReference,
  onOpenTableOfContents,
  onOpenIndexEntry,
  onOpenIndex,
  citationsOpen,
  citationSourceCount,
  onToggleCitations,
  onInsertField,
  onOpenField,
  onRefreshFields,
  onRefreshIndex,
  onRefreshTableOfContents,
  canInsertComment,
  onInsertComment,
  commentsOpen,
  commentCount,
  onToggleComments,
  canAcceptComment = false,
  canProcessComment = false,
  canWithdrawComment = false,
  onAcceptComment,
  onProcessComment,
  onWithdrawComment,
  trackChanges,
  changesOpen,
  changeCount,
  findReplaceMode,
  fileActions,
  onRibbonTabChange,
  onToggleTrackChanges,
  onToggleChanges,
  onOpenComparison,
  onDecideChange,
  onOpenWordCount,
  onOpenFindReplace,
}: DocumentToolbarProps) {
  const messages = useOfficeMessages();
  const [activeTab, setActiveTab] = useState<DocumentRibbonTabId>(
    reviewOnly ? 'review' : 'home',
  );
  const [fontDialogRequest, setFontDialogRequest] =
    useState<DocumentFontDialogRequest | null>(null);
  const [proofingDialogRequest, setProofingDialogRequest] =
    useState<DocumentProofingDialogRequest | null>(null);
  const officeDialog = useOfficeDialog();
  const prompt = officeDialog.prompt;
  const imageSelected = editor.isActive('image');
  const textBoxSelected = editor.isActive('documentTextBox');
  const connectorSelected = editor.isActive('documentConnector');
  const tableSelected = editor.isActive('table');
  const activeBookmark = activeDocumentBookmark(editor);
  const hasRefreshableFields = documentHasRefreshableFields(editor);
  const documentChanges = collectDocumentChanges(editor.state.doc);
  const previousChangeIndex = adjacentDocumentChangeIndex(
    documentChanges,
    editor.state.selection,
    -1,
  );
  const nextChangeIndex = adjacentDocumentChangeIndex(
    documentChanges,
    editor.state.selection,
    1,
  );
  const actionableChangeIndex = actionableDocumentChangeIndex(
    documentChanges,
    editor.state.selection,
  );
  const undoCommand = getDocumentCommandDefinition('undo');
  const redoCommand = getDocumentCommandDefinition('redo');
  const spellingCommand = getDocumentCommandDefinition('spelling');
  const insertCommentCommand = getDocumentCommandDefinition('insertComment');
  const trackChangesCommand = getDocumentCommandDefinition('trackChanges');
  const openFontDialog = useCallback((target: Editor) => {
    if (target.isDestroyed) return;
    synchronizeDocumentEditorSelectionFromDom(target);
    const { from, to } = target.state.selection;
    setFontDialogRequest({
      editor: target,
      selection: { from, to },
      source: documentFontDialogSource(target),
    });
  }, []);
  const openProofingDialog = useCallback((target: Editor) => {
    if (target.isDestroyed) return;
    synchronizeDocumentEditorSelectionFromDom(target);
    const { from, to } = target.state.selection;
    setProofingDialogRequest({
      editor: target,
      selection: { from, to },
      source: documentProofingDialogSource(target),
    });
  }, []);
  const ribbonTabs = (
    reviewOnly
      ? documentRibbonTabs.filter(({ id }) => id === 'review' || id === 'view')
      : pageChromeEditor
        ? [...documentRibbonTabs, documentPageChromeRibbonTab]
        : connectorSelected
          ? [...documentRibbonTabs, documentConnectorRibbonTab]
          : textBoxSelected
            ? [...documentRibbonTabs, documentTextBoxRibbonTab]
            : imageSelected
              ? [...documentRibbonTabs, documentPictureRibbonTab]
              : tableSelected
                ? [...documentRibbonTabs, ...documentTableRibbonTabs]
                : documentRibbonTabs
  ).map((tab) => localizeDocumentRibbonTab(tab, messages));
  const toggleLink = useCallback(async () => {
    if (editor.isActive('link')) {
      // Keep ribbon focus on remove-link. chain().focus() schedules into the
      // editor and breaks L2 loops.
      editor.chain().extendMarkRange('link').unsetLink().run();
      return;
    }
    // Capture the editing selection before the modal steals focus; restoring
    // only `editor.view.dom` can collapse Shift+End ranges and leave setLink
    // with an empty selection (no visible <a href>).
    let { from, to } = editor.state.selection;
    if (from === to) {
      // Cursor-only: wrap the current textblock so the link is visible, matching
      // daily Writer expectation when Ctrl+K is pressed without a range.
      const $pos = editor.state.doc.resolve(from);
      const blockFrom = $pos.start();
      const blockTo = $pos.end();
      if (blockFrom < blockTo) {
        from = blockFrom;
        to = blockTo;
      }
    }
    const href = await prompt({
      title: officeMessage(messages, 'document.link.add'),
      description: officeMessage(messages, 'document.link.prompt.description'),
      fieldLabel: officeMessage(messages, 'document.link.prompt.field'),
      initialValue: editor.getAttributes('link').href ?? 'https://',
      placeholder: 'https://',
      inputMode: 'url',
      confirmLabel: officeMessage(messages, 'document.link.add'),
      required: officeMessage(messages, 'document.link.prompt.required'),
      validate: (value) =>
        normalizeDocumentHref(value) ? null : DOCUMENT_LINK_VALIDATION_MESSAGE,
      restoreFocusTarget: () => editor.view.dom,
    });
    if (href === null) return;
    const normalized = normalizeDocumentHref(href);
    if (!normalized) return;
    const docSize = editor.state.doc.content.size;
    editor
      .chain()
      .focus()
      .setTextSelection({
        from: Math.min(from, docSize),
        to: Math.min(to, docSize),
      })
      .setLink({ href: normalized })
      .run();
  }, [editor, messages, prompt]);
  const toggleBookmark = useCallback(async () => {
    if (activeBookmark) {
      // Keep ribbon focus on delete-bookmark. chain().focus() schedules into
      // the editor and breaks L2 loops.
      editor.commands.deleteDocumentBookmark(activeBookmark.id);
      return;
    }
    const name = await prompt({
      title: officeMessage(messages, 'document.bookmark.add'),
      description: officeMessage(
        messages,
        'document.bookmark.prompt.description',
      ),
      fieldLabel: officeMessage(messages, 'document.bookmark.prompt.field'),
      initialValue: '',
      placeholder: officeMessage(
        messages,
        'document.bookmark.prompt.placeholder',
      ),
      confirmLabel: officeMessage(messages, 'document.bookmark.add'),
      required: officeMessage(messages, 'document.bookmark.prompt.required'),
      validate: (value) =>
        validateDocumentBookmarkName(value) ??
        (documentBookmarkNameExists(editor, value)
          ? DOCUMENT_BOOKMARK_DUPLICATE_MESSAGE
          : null),
      restoreFocusTarget: () => editor.view.dom,
    });
    if (name === null || editor.isDestroyed) return;
    editor.chain().focus().insertDocumentBookmark(name.trim()).run();
  }, [activeBookmark, editor, messages, prompt]);
  useEffect(() => {
    setActiveTab((current) => {
      if (reviewOnly) return current === 'view' ? 'view' : 'review';
      if (pageChromeEditor) return 'pageChrome';
      if (connectorSelected) return 'connector';
      if (textBoxSelected) return 'textBox';
      if (imageSelected) return 'picture';
      if (tableSelected) {
        return current === 'tableDesign' || current === 'tableLayout'
          ? current
          : 'tableDesign';
      }
      return current === 'picture' ||
        current === 'textBox' ||
        current === 'connector' ||
        current === 'tableDesign' ||
        current === 'tableLayout' ||
        current === 'pageChrome'
        ? 'home'
        : current;
    });
  }, [
    imageSelected,
    connectorSelected,
    pageChromeEditor,
    reviewOnly,
    tableSelected,
    textBoxSelected,
  ]);

  useEffect(() => {
    let editorDom: HTMLElement | null = null;
    let root: HTMLElement | null = null;
    const onKeyDown = (event: KeyboardEvent) => {
      if (
        event.defaultPrevented ||
        event.repeat ||
        isOfficeShortcutBlocked(event.target)
      ) {
        return;
      }
      const key = event.key.toLowerCase();
      const insideEditor =
        event.target instanceof Node &&
        Boolean(editorDom?.contains(event.target));
      const insidePageChromeEditor =
        event.target instanceof Node &&
        Boolean(pageChromeEditor?.view.dom.contains(event.target));
      if (reviewOnly) {
        if (
          insideEditor &&
          (event.metaKey || event.ctrlKey) &&
          event.altKey &&
          !event.shiftKey &&
          key === 'm'
        ) {
          event.preventDefault();
          if (canInsertComment) onInsertComment();
        }
        return;
      }
      if (
        insidePageChromeEditor &&
        (event.metaKey || event.ctrlKey) &&
        !event.altKey &&
        !event.shiftKey &&
        key === 'd'
      ) {
        event.preventDefault();
        openFontDialog(pageChromeEditor!);
        return;
      }
      if (
        insideEditor &&
        runDocumentWpsShortcut(editor, event, {
          canInsertComment,
          canRefreshFields: hasRefreshableFields,
          onInsertComment,
          onOpenFontDialog: () => openFontDialog(editor),
          onOpenWordCount,
          onRefreshFields,
          onToggleFieldCodes,
          onToggleSelectedFieldCodes,
          onUnlinkFields,
          onLockFields,
          onUnlockFields,
          onToggleSpellcheck,
          onToggleTrackChanges,
        })
      ) {
        event.preventDefault();
        return;
      }
      if (event.altKey || !(event.metaKey || event.ctrlKey)) return;
      const documentHistoryTarget =
        insideEditor || !isDocumentNativeTextUndoTarget(event.target);
      if (documentHistoryTarget && key === 'z') {
        event.preventDefault();
        if (history) {
          if (event.shiftKey) history.redo();
          else history.undo();
          return;
        }
        const command = event.shiftKey
          ? editor.chain().focus().redo()
          : editor.chain().focus().undo();
        command.run();
        return;
      }
      if (documentHistoryTarget && key === 'y' && !event.shiftKey) {
        event.preventDefault();
        if (history) {
          history.redo();
          return;
        }
        editor.chain().focus().redo().run();
        return;
      }
      if (
        !event.shiftKey &&
        insideEditor &&
        (key === 'b' || key === 'i' || key === 'u')
      ) {
        event.preventDefault();
        if (key === 'b') editor.chain().focus().toggleBold().run();
        else if (key === 'i') editor.chain().focus().toggleItalic().run();
        else editor.chain().focus().toggleUnderline().run();
        return;
      }
      if (insideEditor && key === 'k' && !event.shiftKey) {
        event.preventDefault();
        void toggleLink();
        return;
      }
      if ((key === 'f' || key === 'h') && !event.shiftKey) {
        event.preventDefault();
        onOpenFindReplace(key === 'h' ? 'replace' : 'find');
        return;
      }
      if (key !== 'enter' || event.shiftKey || !insideEditor) {
        return;
      }
      event.preventDefault();
      editor.chain().focus().insertContent({ type: 'pageBreak' }).run();
    };
    const detach = () => {
      root?.removeEventListener('keydown', onKeyDown, true);
      root = null;
      editorDom = null;
    };
    const attach = () => {
      detach();
      if (editor.isDestroyed) return;
      editorDom = editor.view.dom;
      root = editorDom.closest<HTMLElement>('.work-document-editor');
      root?.addEventListener('keydown', onKeyDown, true);
    };
    attach();
    editor.on('mount', attach);
    editor.on('unmount', detach);
    return () => {
      editor.off('mount', attach);
      editor.off('unmount', detach);
      detach();
    };
  }, [
    canInsertComment,
    editor,
    hasRefreshableFields,
    history,
    onInsertComment,
    openFontDialog,
    onOpenFindReplace,
    onOpenWordCount,
    onRefreshFields,
    reviewOnly,
    onToggleFieldCodes,
    onToggleSelectedFieldCodes,
    onUnlinkFields,
    onLockFields,
    onUnlockFields,
    onToggleSpellcheck,
    onToggleTrackChanges,
    pageChromeEditor,
    toggleLink,
  ]);

  const selectDocumentChange = (change: WorkDocumentChange) => {
    const maximum = editor.state.doc.content.size;
    editor
      .chain()
      .focus()
      .setTextSelection({
        from: Math.min(change.from, maximum),
        to: Math.min(change.to, maximum),
      })
      .scrollIntoView()
      .run();
  };
  const navigateDocumentChange = (direction: -1 | 1) => {
    const changes = collectDocumentChanges(editor.state.doc);
    const index = adjacentDocumentChangeIndex(
      changes,
      editor.state.selection,
      direction,
    );
    if (index !== null) selectDocumentChange(changes[index]);
  };
  const decideDocumentChange = (decision: 'accept' | 'reject') => {
    const changes = collectDocumentChanges(editor.state.doc);
    const index = actionableDocumentChangeIndex(
      changes,
      editor.state.selection,
    );
    if (index === null) return;
    const change = changes[index];
    const nextIds = [changes[index + 1]?.id, changes[index - 1]?.id].filter(
      (id): id is string => Boolean(id),
    );
    const handled = onDecideChange
      ? onDecideChange(change, decision)
      : decision === 'accept'
        ? editor.commands.acceptDocumentChange(change.id)
        : editor.commands.rejectDocumentChange(change.id);
    if (!handled) return;
    const remaining = collectDocumentChanges(editor.state.doc);
    const next = nextIds
      .map((id) => remaining.find((candidate) => candidate.id === id))
      .find((candidate) => Boolean(candidate));
    if (next) selectDocumentChange(next);
    else {
      editor
        .chain()
        .focus()
        .setTextSelection(Math.min(change.from, editor.state.doc.content.size))
        .scrollIntoView()
        .run();
    }
  };

  return (
    <>
      <WorkOfficeRibbon
        ariaLabel={officeMessage(messages, 'document.toolbar.aria')}
        tabs={ribbonTabs}
        defaultTab={reviewOnly ? 'review' : 'home'}
        activeTab={activeTab}
        onTabChange={(tab) => {
          const accepted = onRibbonTabChange?.(tab);
          if (accepted instanceof Promise) {
            void accepted.then((result) => {
              if (result !== false) setActiveTab(tab);
            });
          } else if (accepted !== false) {
            setActiveTab(tab);
          }
        }}
        adaptive
        collapsible
        fileActions={fileActions}
        quickAccessActions={
          reviewOnly && !suggestionOnly
            ? []
            : [
                {
                  id: undoCommand.id,
                  label: documentCommandLabel('undo', messages),
                  icon: <Undo2 size={15} />,
                  shortcut: undoCommand.shortcut?.label,
                  ariaKeyShortcuts: undoCommand.shortcut?.aria,
                  disabled:
                    editor.isDestroyed ||
                    (history
                      ? !history.canUndo
                      : !editor.can().chain().undo().run()),
                  onSelect: () => {
                    if (history) {
                      history.undo();
                      return;
                    }
                    // Keep QAT focus on undo. chain().focus() schedules into
                    // the editor and breaks L2 loops.
                    editor.commands.undo();
                  },
                },
                {
                  id: redoCommand.id,
                  label: documentCommandLabel('redo', messages),
                  icon: <Redo2 size={15} />,
                  shortcut: redoCommand.shortcut?.label,
                  ariaKeyShortcuts: redoCommand.shortcut?.aria,
                  disabled:
                    editor.isDestroyed ||
                    (history
                      ? !history.canRedo
                      : !editor.can().chain().redo().run()),
                  onSelect: () => {
                    if (history) {
                      history.redo();
                      return;
                    }
                    // Keep QAT focus on redo. chain().focus() schedules into
                    // the editor and breaks L2 loops.
                    editor.commands.redo();
                  },
                },
              ]
        }
        defaultCollapsed={defaultRibbonCollapsed}
        className="work-document-ribbon"
        toolbarClassName="document-toolbar"
        panels={{
          home: reviewOnly ? null : (
            <DocumentHomeRibbon
              editor={editor}
              findReplaceMode={findReplaceMode}
              layoutFonts={layoutFonts}
              onOpenFontDialog={() => openFontDialog(editor)}
              onFindText={(replace) =>
                onOpenFindReplace(replace ? 'replace' : 'find')
              }
            />
          ),
          insert: reviewOnly ? null : (
            <>
              <RibbonGroup
                label={officeMessage(messages, 'document.group.page')}
                priority="high"
              >
                <ToolbarButton
                  label={documentCommandLabel('insertPageBreak', messages)}
                  shortcut="Cmd/Ctrl+Enter"
                  ariaKeyShortcuts="Control+Enter Meta+Enter"
                  displayLabel
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => {
                    // Keep ribbon focus on Insert page-break. chain().focus()
                    // schedules into the editor and breaks L2 loops.
                    editor.commands.insertContent({ type: 'pageBreak' });
                  }}
                >
                  <FilePlus2 size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup
                label={officeMessage(messages, 'document.group.table')}
                priority="high"
              >
                <DocumentTableInsertPopover editor={editor} />
              </RibbonGroup>
              <RibbonGroup
                label={officeMessage(messages, 'document.group.illustrations')}
              >
                <ToolbarButton
                  label={officeMessage(messages, 'document.insert.picture')}
                  displayLabel
                  onClick={onRequestImage}
                >
                  <ImageIcon size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup
                label={officeMessage(messages, 'document.group.links')}
                priority="low"
              >
                <ToolbarButton
                  label={
                    editor.isActive('link')
                      ? officeMessage(messages, 'document.link.remove')
                      : officeMessage(messages, 'document.link.add')
                  }
                  shortcut="Cmd/Ctrl+K"
                  ariaKeyShortcuts="Control+K Meta+K"
                  displayLabel
                  active={editor.isActive('link')}
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => void toggleLink()}
                >
                  <Link2 size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label={
                    activeBookmark
                      ? officeMessage(messages, 'document.bookmark.remove')
                      : officeMessage(messages, 'document.bookmark.add')
                  }
                  displayLabel
                  active={Boolean(activeBookmark)}
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => void toggleBookmark()}
                >
                  <BookmarkIcon size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup
                label={officeMessage(messages, 'document.group.headerFooter')}
              >
                <ToolbarButton
                  label={officeMessage(messages, 'document.pageChrome.header')}
                  displayLabel
                  onClick={() => onPageChromeEditingPartChange('header')}
                >
                  <PanelTopOpen size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label={officeMessage(messages, 'document.pageChrome.footer')}
                  displayLabel
                  onClick={() => onPageChromeEditingPartChange('footer')}
                >
                  <PanelBottomOpen size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label={officeMessage(
                    messages,
                    'document.pageChrome.pageNumber',
                  )}
                  displayLabel
                  active={showPageNumbers}
                  onClick={onTogglePageNumbers}
                >
                  <Hash size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup
                label={officeMessage(messages, 'document.group.text')}
                priority="low"
              >
                <ToolbarButton
                  label={documentCommandLabel('insertTextBox', messages)}
                  displayLabel
                  onClick={() => onInsertTextBox?.()}
                >
                  <TextCursorInput size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label={documentCommandLabel('insertConnector', messages)}
                  displayLabel
                  onClick={() => onInsertConnector?.()}
                >
                  <ArrowUpRight size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label={documentCommandLabel('insertContentControl', messages)}
                  displayLabel
                  onClick={() => onInsertContentControl?.()}
                >
                  <TextSelect size={19} />
                </ToolbarButton>
                <DocumentFieldInsertMenu onInsertField={onInsertField} />
                <ToolbarButton
                  label={officeMessage(messages, 'document.field.settings')}
                  title={officeMessage(
                    messages,
                    'document.field.settingsTitle',
                  )}
                  displayLabel
                  onClick={onOpenField}
                >
                  <SlidersHorizontal size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label={officeMessage(
                    messages,
                    'document.mailMerge.importRecipients',
                  )}
                  title={officeMessage(
                    messages,
                    'document.mailMerge.importRecipientsTitle',
                  )}
                  displayLabel
                  onClick={() => onImportMailMergeRecipients?.()}
                >
                  <Users size={19} />
                </ToolbarButton>
                {onOpenMailMergeRecipientFilter ? (
                  <ToolbarButton
                    label={officeMessage(
                      messages,
                      'document.mailMerge.filterRecipients',
                    )}
                    title={officeMessage(
                      messages,
                      'document.mailMerge.filterRecipientsTitle',
                    )}
                    displayLabel
                    onClick={onOpenMailMergeRecipientFilter}
                  >
                    <Filter size={19} />
                  </ToolbarButton>
                ) : null}
              </RibbonGroup>
            </>
          ),
          page: reviewOnly ? null : (
            <DocumentPageLayoutRibbon
              editor={editor}
              layout={layout}
              layoutOpen={layoutOpen}
              pageColor={pageColor}
              onLayoutChange={onLayoutChange}
              onOpenLayout={onOpenLayout}
              onToggleLayout={onToggleLayout}
              onPageColorChange={onPageColorChange}
              onInsertSection={onInsertSection}
            />
          ),
          references: reviewOnly ? null : (
            <DocumentReferencesRibbon
              editor={editor}
              citationsOpen={citationsOpen}
              citationSourceCount={citationSourceCount}
              onInsertNote={onInsertNote}
              onInsertCaption={onInsertCaption}
              onInsertCrossReference={onInsertCrossReference}
              onOpenTableOfContents={onOpenTableOfContents}
              onRefreshTableOfContents={onRefreshTableOfContents}
              onOpenIndexEntry={onOpenIndexEntry}
              onOpenIndex={onOpenIndex}
              onRefreshIndex={onRefreshIndex}
              onToggleCitations={onToggleCitations}
              onRefreshFields={onRefreshFields}
              onUnlinkFields={() => {
                onUnlinkFields();
              }}
              onLockFields={() => {
                onLockFields();
              }}
              onUnlockFields={() => {
                onUnlockFields();
              }}
            />
          ),
          review: (
            <>
              {!reviewOnly && (
                <RibbonGroup
                  label={officeMessage(messages, 'document.group.proofing')}
                  priority="high"
                >
                  <ToolbarButton
                    label={documentCommandLabel('spelling', messages)}
                    displayLabel
                    shortcut={spellingCommand.shortcut?.label}
                    ariaKeyShortcuts={spellingCommand.shortcut?.aria}
                    active={spellcheckEnabled}
                    onClick={onToggleSpellcheck}
                  >
                    <CheckCheck size={19} />
                  </ToolbarButton>
                  <ToolbarButton
                    label={officeMessage(
                      messages,
                      'document.proofing.language',
                    )}
                    displayLabel
                    onClick={() => openProofingDialog(editor)}
                  >
                    <Languages size={19} />
                  </ToolbarButton>
                </RibbonGroup>
              )}
              {!suggestionOnly && (
                <RibbonGroup
                  label={officeMessage(messages, 'document.group.comments')}
                  priority="high"
                >
                  <ToolbarButton
                    label={documentCommandLabel('insertComment', messages)}
                    displayLabel
                    shortcut={insertCommentCommand.shortcut?.label}
                    ariaKeyShortcuts={insertCommentCommand.shortcut?.aria}
                    disabled={!canInsertComment}
                    title={
                      canInsertComment
                        ? officeMessage(
                            messages,
                            'document.comment.insertTitle',
                            {
                              shortcut:
                                insertCommentCommand.shortcut?.label ?? '',
                            },
                          )
                        : officeMessage(
                            messages,
                            'document.comment.selectFirst',
                          )
                    }
                    onClick={onInsertComment}
                  >
                    <MessageSquarePlus size={19} />
                  </ToolbarButton>
                  <ToolbarButton
                    label={
                      commentCount
                        ? officeMessage(
                            messages,
                            'document.comment.viewWithCount',
                            {
                              count: String(commentCount),
                            },
                          )
                        : officeMessage(messages, 'document.comment.view')
                    }
                    displayLabel
                    active={commentsOpen}
                    onClick={onToggleComments}
                  >
                    <MessagesSquare size={19} />
                  </ToolbarButton>
                  <ToolbarButton
                    label={officeMessage(messages, 'document.comment.accept')}
                    displayLabel
                    disabled={!canAcceptComment}
                    title={officeMessage(
                      messages,
                      'document.comment.acceptTitle',
                    )}
                    onClick={() => onAcceptComment?.()}
                  >
                    <Check size={19} />
                  </ToolbarButton>
                  <ToolbarButton
                    label={officeMessage(messages, 'document.comment.process')}
                    displayLabel
                    disabled={!canProcessComment}
                    title={officeMessage(
                      messages,
                      'document.comment.processTitle',
                    )}
                    onClick={() => onProcessComment?.()}
                  >
                    <ListChecks size={19} />
                  </ToolbarButton>
                  <ToolbarButton
                    label={officeMessage(messages, 'document.comment.withdraw')}
                    displayLabel
                    disabled={!canWithdrawComment}
                    title={officeMessage(
                      messages,
                      canWithdrawComment
                        ? 'document.comment.withdrawTitle'
                        : 'document.comment.withdrawBlocked',
                    )}
                    onClick={() => onWithdrawComment?.()}
                  >
                    <Undo2 size={19} />
                  </ToolbarButton>
                </RibbonGroup>
              )}
              {(!reviewOnly || suggestionOnly) && (
                <>
                  <RibbonGroup
                    label={officeMessage(messages, 'document.group.revisions')}
                    priority="high"
                  >
                    <ToolbarButton
                      label={
                        suggestionOnly
                          ? officeMessage(
                              messages,
                              'document.track.suggestionMode',
                            )
                          : documentCommandLabel('trackChanges', messages)
                      }
                      displayLabel
                      shortcut={trackChangesCommand.shortcut?.label}
                      ariaKeyShortcuts={trackChangesCommand.shortcut?.aria}
                      active={suggestionOnly || trackChanges}
                      disabled={suggestionOnly}
                      title={
                        suggestionOnly
                          ? officeMessage(
                              messages,
                              'document.track.suggestionModeHint',
                            )
                          : undefined
                      }
                      onClick={onToggleTrackChanges}
                    >
                      <FileDiff size={19} />
                    </ToolbarButton>
                    <ToolbarButton
                      label={
                        changeCount
                          ? officeMessage(
                              messages,
                              'document.changes.viewWithCount',
                              {
                                count: String(changeCount),
                              },
                            )
                          : officeMessage(messages, 'document.changes.view')
                      }
                      displayLabel
                      active={changesOpen}
                      onClick={onToggleChanges}
                    >
                      <ListChecks size={19} />
                    </ToolbarButton>
                  </RibbonGroup>
                  {!suggestionOnly && (
                    <RibbonGroup
                      label={officeMessage(messages, 'document.group.changes')}
                      priority="high"
                    >
                      <ToolbarButton
                        label={officeMessage(
                          messages,
                          'document.changes.accept',
                        )}
                        displayLabel
                        disabled={actionableChangeIndex === null}
                        title={officeMessage(
                          messages,
                          'document.changes.acceptTitle',
                        )}
                        onClick={() => decideDocumentChange('accept')}
                      >
                        <Check size={19} />
                      </ToolbarButton>
                      <ToolbarButton
                        label={officeMessage(
                          messages,
                          'document.changes.reject',
                        )}
                        displayLabel
                        disabled={actionableChangeIndex === null}
                        title={officeMessage(
                          messages,
                          'document.changes.rejectTitle',
                        )}
                        onClick={() => decideDocumentChange('reject')}
                      >
                        <XCircle size={19} />
                      </ToolbarButton>
                      <ToolbarButton
                        label={officeMessage(
                          messages,
                          'document.changes.previous',
                        )}
                        displayLabel
                        disabled={previousChangeIndex === null}
                        onClick={() => navigateDocumentChange(-1)}
                      >
                        <ChevronUp size={19} />
                      </ToolbarButton>
                      <ToolbarButton
                        label={officeMessage(messages, 'document.changes.next')}
                        displayLabel
                        disabled={nextChangeIndex === null}
                        onClick={() => navigateDocumentChange(1)}
                      >
                        <ChevronDown size={19} />
                      </ToolbarButton>
                    </RibbonGroup>
                  )}
                </>
              )}
              {!reviewOnly && !suggestionOnly && (
                <RibbonGroup
                  label={officeMessage(messages, 'document.group.compare')}
                  priority="high"
                >
                  <ToolbarButton
                    label={officeMessage(
                      messages,
                      'document.compare.documents',
                    )}
                    displayLabel
                    title={officeMessage(
                      messages,
                      'document.compare.documentsTitle',
                    )}
                    onClick={() => onOpenComparison('compare')}
                  >
                    <GitCompareArrows size={19} />
                  </ToolbarButton>
                  <ToolbarButton
                    label={officeMessage(messages, 'document.compare.merge')}
                    displayLabel
                    title={officeMessage(
                      messages,
                      'document.compare.mergeTitle',
                    )}
                    onClick={() => onOpenComparison('combine')}
                  >
                    <FileStack size={19} />
                  </ToolbarButton>
                </RibbonGroup>
              )}
            </>
          ),
          view: (
            <>
              <RibbonGroup
                label={officeMessage(messages, 'document.group.documentViews')}
                priority="high"
              >
                <ToolbarButton
                  label={officeMessage(messages, 'document.view.page')}
                  displayLabel
                  active={viewMode === 'page'}
                  onClick={() => onViewModeChange('page')}
                >
                  <FileText size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label={officeMessage(messages, 'document.view.web')}
                  displayLabel
                  active={viewMode === 'web'}
                  onClick={() => onViewModeChange('web')}
                >
                  <Globe2 size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup
                label={officeMessage(messages, 'document.group.show')}
                priority="high"
              >
                <ToolbarButton
                  label={officeMessage(messages, 'document.view.ruler')}
                  displayLabel
                  active={showRulers}
                  disabled={viewMode !== 'page'}
                  title={
                    viewMode === 'page'
                      ? officeMessage(messages, 'document.view.rulerTitle')
                      : officeMessage(messages, 'document.view.rulerDisabled')
                  }
                  onClick={onToggleRulers}
                >
                  <Ruler size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label={documentCommandLabel('navigationPane', messages)}
                  displayLabel
                  active={navigationOpen}
                  onClick={onToggleNavigation}
                >
                  <PanelLeftOpen size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label={documentCommandLabel('toggleFieldCodes', messages)}
                  displayLabel
                  active={showFieldCodes}
                  shortcut={
                    getDocumentCommandDefinition('toggleFieldCodes').shortcut
                      ?.label
                  }
                  ariaKeyShortcuts={
                    getDocumentCommandDefinition('toggleFieldCodes').shortcut
                      ?.aria
                  }
                  title={officeMessage(
                    messages,
                    'document.view.toggleFieldCodesTitle',
                  )}
                  onClick={onToggleFieldCodes}
                >
                  <Braces size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label={documentCommandLabel('showHiddenText', messages)}
                  displayLabel
                  active={showHiddenText}
                  title={officeMessage(
                    messages,
                    'document.view.showHiddenTitle',
                  )}
                  onClick={onToggleHiddenText}
                >
                  <Eye size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup
                label={officeMessage(messages, 'document.group.zoom', {
                  percent: String(zoom),
                })}
                priority="low"
              >
                <ToolbarButton
                  label={officeMessage(messages, 'document.zoom.out')}
                  disabled={zoom <= MIN_DOCUMENT_ZOOM}
                  onClick={() => onZoomChange(zoom - 10)}
                >
                  <ZoomOut size={17} />
                </ToolbarButton>
                <ToolbarButton
                  label={officeMessage(messages, 'document.zoom.100')}
                  active={zoom === 100}
                  onClick={() => onZoomChange(100)}
                >
                  100%
                </ToolbarButton>
                <ToolbarButton
                  label={officeMessage(messages, 'document.zoom.onePage')}
                  displayLabel
                  onClick={() => onZoomFit('page')}
                >
                  <Scan size={17} />
                </ToolbarButton>
                <ToolbarButton
                  label={officeMessage(messages, 'document.zoom.pageWidth')}
                  displayLabel
                  onClick={() => onZoomFit('width')}
                >
                  <StretchHorizontal size={17} />
                </ToolbarButton>
                <ToolbarButton
                  label={officeMessage(messages, 'document.zoom.in')}
                  disabled={zoom >= MAX_DOCUMENT_ZOOM}
                  onClick={() => onZoomChange(zoom + 10)}
                >
                  <ZoomIn size={17} />
                </ToolbarButton>
              </RibbonGroup>
            </>
          ),
          picture: imageSelected ? (
            <DocumentPictureRibbon editor={editor} />
          ) : null,
          tableDesign: tableSelected ? (
            <DocumentTableDesignRibbon editor={editor} />
          ) : null,
          tableLayout: tableSelected ? (
            <DocumentTableLayoutRibbon editor={editor} />
          ) : null,
          pageChrome:
            pageChromeEditor && pageChromeEditingPart ? (
              <DocumentPageChromeRibbon
                editor={pageChromeEditor}
                editingPart={pageChromeEditingPart}
                showPageNumber={pageChromeShowPageNumber}
                onEditingPartChange={onPageChromeEditingPartChange}
                onTogglePageNumber={onTogglePageChromePageNumber}
                onOpenFontDialog={() => openFontDialog(pageChromeEditor)}
                onOpenProofingDialog={() =>
                  openProofingDialog(pageChromeEditor)
                }
                onClose={onClosePageChrome}
              />
            ) : null,
          textBox: textBoxSelected ? (
            <DocumentTextBoxRibbon editor={editor} />
          ) : null,
          connector: connectorSelected ? (
            <DocumentConnectorRibbon editor={editor} />
          ) : null,
        }}
      />
      {fontDialogRequest && (
        <DocumentFontDialog
          source={fontDialogRequest.source}
          layoutFonts={layoutFonts}
          restoreFocusTarget={() =>
            fontDialogRequest.editor.isDestroyed
              ? null
              : fontDialogRequest.editor.view.dom
          }
          onApply={(patch) =>
            applyDocumentFontDialogPatch(
              fontDialogRequest.editor,
              fontDialogRequest.selection,
              patch,
            )
          }
          onClose={() => {
            const { editor: dialogEditor, selection } = fontDialogRequest;
            setFontDialogRequest(null);
            requestAnimationFrame(() => {
              if (dialogEditor.isDestroyed) return;
              dialogEditor
                .chain()
                .setTextSelection(selection)
                .focus(null, { scrollIntoView: false })
                .run();
            });
          }}
        />
      )}
      {proofingDialogRequest && (
        <DocumentProofingDialog
          source={proofingDialogRequest.source}
          restoreFocusTarget={() =>
            proofingDialogRequest.editor.isDestroyed
              ? null
              : proofingDialogRequest.editor.view.dom
          }
          onApply={(patch) =>
            applyDocumentProofingDialogPatch(
              proofingDialogRequest.editor,
              proofingDialogRequest.selection,
              patch,
            )
          }
          onClose={() => {
            const { editor: dialogEditor, selection } = proofingDialogRequest;
            setProofingDialogRequest(null);
            requestAnimationFrame(() => {
              if (dialogEditor.isDestroyed) return;
              dialogEditor
                .chain()
                .setTextSelection(selection)
                .focus(null, { scrollIntoView: false })
                .run();
            });
          }}
        />
      )}
      {officeDialog.dialog}
    </>
  );
}

function isDocumentNativeTextUndoTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  return (
    (target instanceof HTMLInputElement &&
      documentNativeUndoInputTypes.has(target.type)) ||
    target instanceof HTMLTextAreaElement ||
    target instanceof HTMLSelectElement ||
    target.isContentEditable ||
    Boolean(target.closest('[contenteditable="true"]'))
  );
}

const documentNativeUndoInputTypes = new Set([
  'date',
  'datetime-local',
  'email',
  'month',
  'number',
  'password',
  'search',
  'tel',
  'text',
  'time',
  'url',
  'week',
]);

function ToolbarButton({
  label,
  title,
  shortcut,
  ariaKeyShortcuts,
  active = false,
  disabled = false,
  displayLabel = false,
  onMouseDown,
  onClick,
  children,
}: {
  label: string;
  title?: string;
  shortcut?: string;
  ariaKeyShortcuts?: string;
  active?: boolean;
  disabled?: boolean;
  displayLabel?: boolean;
  onMouseDown?: (event: ReactMouseEvent<HTMLButtonElement>) => void;
  onClick: () => void;
  children: ReactNode;
}) {
  return (
    <WorkOfficeRibbonButton
      label={label}
      visibleLabel={label.replace(/（\d+）$/, '')}
      title={title ?? (shortcut ? `${label}（${shortcut}）` : label)}
      aria-keyshortcuts={ariaKeyShortcuts}
      active={active}
      displayLabel={displayLabel}
      disabled={disabled}
      onMouseDown={onMouseDown}
      onClick={onClick}
    >
      {children}
    </WorkOfficeRibbonButton>
  );
}

const RibbonGroup = WorkOfficeRibbonGroup;

const DOCUMENT_FIELD_INSERT_KINDS = [
  'page',
  'numPages',
  'section',
  'sectionPages',
  'date',
  'time',
  'createDate',
  'saveDate',
  'printDate',
  'wordCount',
  'characterCount',
  'fileName',
  'author',
  'title',
  'subject',
  'keywords',
  'lastSavedBy',
  'comments',
] as const satisfies readonly WorkDocumentFieldKind[];

function DocumentFieldInsertMenu({
  onInsertField,
}: {
  onInsertField: (kind: WorkDocumentFieldKind) => void;
}) {
  const messages = useOfficeMessages();
  const documentFieldInsertActions = DOCUMENT_FIELD_INSERT_KINDS.map(
    (value) => ({
      value,
      label: officeMessage(messages, `document.field.kind.${value}`),
    }),
  );
  return (
    <Popover
      label={officeMessage(messages, 'document.field.insertMenu')}
      panelLabel={officeMessage(messages, 'document.field.insertMenu')}
      panelRole="menu"
      portal
      className="work-document-field-insert-menu"
      panelClassName="work-office-context-menu work-document-field-insert-menu-panel"
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          type="button"
          className={`work-document-field-insert-trigger${open ? ' open' : ''}`}
          title={officeMessage(messages, 'document.field.insertMenu')}
        >
          <span>{officeMessage(messages, 'document.field.insertLabel')}</span>
          <ChevronDown size={14} aria-hidden="true" />
        </button>
      )}
    >
      {(close) =>
        documentFieldInsertActions.map(({ value, label }) => (
          <button
            key={value}
            type="button"
            role="menuitem"
            tabIndex={-1}
            onClick={() => {
              close();
              onInsertField(value);
            }}
          >
            <span>{label}</span>
          </button>
        ))
      }
    </Popover>
  );
}
