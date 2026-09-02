import type { Editor } from '@tiptap/core';
import {
  Bookmark as BookmarkIcon,
  Check,
  CheckCheck,
  ChevronDown,
  ChevronUp,
  FileDiff,
  FilePlus2,
  FileStack,
  FileText,
  GitCompareArrows,
  Globe2,
  Hash,
  Eye,
  Image as ImageIcon,
  Link2,
  ListChecks,
  Languages,
  MessageSquarePlus,
  MessagesSquare,
  PanelBottomOpen,
  PanelLeftOpen,
  PanelTopOpen,
  Redo2,
  Ruler,
  Scan,
  StretchHorizontal,
  TextCursorInput,
  Undo2,
  XCircle,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import { type ReactNode, useCallback, useEffect, useState } from 'react';
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
  documentPageChromeRibbonTab,
  documentPictureRibbonTab,
  documentRibbonTabs,
  documentTableRibbonTabs,
  documentTextBoxRibbonTab,
  getDocumentCommandDefinition,
} from './document-command-catalog';
import { synchronizeDocumentEditorSelectionFromDom } from './document-dom-selection';
import { documentHasRefreshableFields } from './document-editor-support';
import type { DocumentFindReplaceMode } from './document-find-replace-panel';
import { DocumentProofingDialog } from './document-proofing-dialog';
import {
  applyDocumentProofingDialogPatch,
  documentProofingDialogSource,
  type DocumentProofingDialogSource,
} from './document-proofing-dialog-model';
import { DocumentHomeRibbon } from './document-home-ribbon';
import { DocumentFontDialog } from './document-font-dialog';
import {
  applyDocumentFontDialogPatch,
  documentFontDialogSource,
  type DocumentFontDialogSource,
} from './document-font-dialog-model';
import type { DocumentLayoutPanelTab } from './document-layout-panel';
import {
  type DocumentPageChromeEditingPart,
  DocumentPageChromeRibbon,
} from './document-page-chrome-ribbon';
import { DocumentPageLayoutRibbon } from './document-page-layout-ribbon';
import { DocumentPictureRibbon } from './document-picture-ribbon';
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
import { OfficeSelect, useOfficeDialog } from './office-controls';
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
  onPageChromeEditingPartChange: (part: DocumentPageChromeEditingPart) => void;
  onClosePageChrome: () => void;
  onTogglePageChromePageNumber: () => void;
  onToggleLayout: () => void;
  onLayoutChange: (layout: WorkDocumentSectionLayout) => void;
  onOpenLayout: (target: DocumentLayoutPanelTab) => void;
  onToggleNavigation: () => void;
  onTogglePageNumbers: () => void;
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
  onRefreshFields: () => void;
  onRefreshIndex: () => void;
  onRefreshTableOfContents: () => void;
  canInsertComment: boolean;
  onInsertComment: () => void;
  commentsOpen: boolean;
  commentCount: number;
  onToggleComments: () => void;
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
  onPageChromeEditingPartChange,
  onClosePageChrome,
  onTogglePageChromePageNumber,
  onToggleLayout,
  onLayoutChange,
  onOpenLayout,
  onToggleNavigation,
  onTogglePageNumbers,
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
  onRefreshFields,
  onRefreshIndex,
  onRefreshTableOfContents,
  canInsertComment,
  onInsertComment,
  commentsOpen,
  commentCount,
  onToggleComments,
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
  const ribbonTabs = reviewOnly
    ? documentRibbonTabs.filter(({ id }) => id === 'review' || id === 'view')
    : pageChromeEditor
      ? [...documentRibbonTabs, documentPageChromeRibbonTab]
      : textBoxSelected
        ? [...documentRibbonTabs, documentTextBoxRibbonTab]
        : imageSelected
          ? [...documentRibbonTabs, documentPictureRibbonTab]
          : tableSelected
            ? [...documentRibbonTabs, ...documentTableRibbonTabs]
            : documentRibbonTabs;
  const toggleLink = useCallback(async () => {
    if (editor.isActive('link')) {
      editor.chain().focus().unsetLink().run();
      return;
    }
    const href = await prompt({
      title: '添加链接',
      description: '输入网页、邮箱地址，或使用 #书签名称 跳转到文档内位置。',
      fieldLabel: '链接地址',
      initialValue: editor.getAttributes('link').href ?? 'https://',
      placeholder: 'https://',
      inputMode: 'url',
      confirmLabel: '添加链接',
      required: '请输入链接地址。',
      validate: (value) =>
        normalizeDocumentHref(value) ? null : DOCUMENT_LINK_VALIDATION_MESSAGE,
      restoreFocusTarget: () => editor.view.dom,
    });
    if (href === null) return;
    const normalized = normalizeDocumentHref(href);
    if (normalized) editor.chain().focus().setLink({ href: normalized }).run();
  }, [editor, prompt]);
  const toggleBookmark = useCallback(async () => {
    if (activeBookmark) {
      editor.chain().focus().deleteDocumentBookmark(activeBookmark.id).run();
      return;
    }
    const name = await prompt({
      title: '添加书签',
      description: '为当前光标位置或选中内容创建文档内链接目标。',
      fieldLabel: '书签名称',
      initialValue: '',
      placeholder: '例如 Architecture_2',
      confirmLabel: '添加书签',
      required: '请输入书签名称。',
      validate: (value) =>
        validateDocumentBookmarkName(value) ??
        (documentBookmarkNameExists(editor, value)
          ? DOCUMENT_BOOKMARK_DUPLICATE_MESSAGE
          : null),
      restoreFocusTarget: () => editor.view.dom,
    });
    if (name === null || editor.isDestroyed) return;
    editor.chain().focus().insertDocumentBookmark(name.trim()).run();
  }, [activeBookmark, editor, prompt]);
  useEffect(() => {
    setActiveTab((current) => {
      if (reviewOnly) return current === 'view' ? 'view' : 'review';
      if (pageChromeEditor) return 'pageChrome';
      if (textBoxSelected) return 'textBox';
      if (imageSelected) return 'picture';
      if (tableSelected) {
        return current === 'tableDesign' || current === 'tableLayout'
          ? current
          : 'tableDesign';
      }
      return current === 'picture' ||
        current === 'textBox' ||
        current === 'tableDesign' ||
        current === 'tableLayout' ||
        current === 'pageChrome'
        ? 'home'
        : current;
    });
  }, [
    imageSelected,
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
        ariaLabel="文字功能区"
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
                  label: undoCommand.label,
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
                    editor.chain().focus().undo().run();
                  },
                },
                {
                  id: redoCommand.id,
                  label: redoCommand.label,
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
                    editor.chain().focus().redo().run();
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
              <RibbonGroup label="页面" priority="high">
                <ToolbarButton
                  label="插入分页符"
                  shortcut="Cmd/Ctrl+Enter"
                  ariaKeyShortcuts="Control+Enter Meta+Enter"
                  displayLabel
                  onClick={() =>
                    editor
                      .chain()
                      .focus()
                      .insertContent({ type: 'pageBreak' })
                      .run()
                  }
                >
                  <FilePlus2 size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup label="表格" priority="high">
                <DocumentTableInsertPopover editor={editor} />
              </RibbonGroup>
              <RibbonGroup label="插图">
                <ToolbarButton
                  label="插入图片"
                  displayLabel
                  onClick={onRequestImage}
                >
                  <ImageIcon size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup label="链接" priority="low">
                <ToolbarButton
                  label={editor.isActive('link') ? '取消链接' : '添加链接'}
                  shortcut="Cmd/Ctrl+K"
                  ariaKeyShortcuts="Control+K Meta+K"
                  displayLabel
                  active={editor.isActive('link')}
                  onClick={() => void toggleLink()}
                >
                  <Link2 size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label={activeBookmark ? '删除书签' : '添加书签'}
                  displayLabel
                  active={Boolean(activeBookmark)}
                  onClick={() => void toggleBookmark()}
                >
                  <BookmarkIcon size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup label="页眉和页脚">
                <ToolbarButton
                  label="页眉"
                  displayLabel
                  onClick={() => onPageChromeEditingPartChange('header')}
                >
                  <PanelTopOpen size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label="页脚"
                  displayLabel
                  onClick={() => onPageChromeEditingPartChange('footer')}
                >
                  <PanelBottomOpen size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label="页码"
                  displayLabel
                  active={showPageNumbers}
                  onClick={onTogglePageNumbers}
                >
                  <Hash size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup label="文本" priority="low">
                <ToolbarButton
                  label="插入文本框"
                  displayLabel
                  onClick={() => onInsertTextBox?.()}
                >
                  <TextCursorInput size={19} />
                </ToolbarButton>
                <DocumentFieldSelect onInsertField={onInsertField} />
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
            />
          ),
          review: (
            <>
              {!reviewOnly && (
                <RibbonGroup label="校对" priority="high">
                  <ToolbarButton
                    label="拼写检查"
                    displayLabel
                    shortcut={spellingCommand.shortcut?.label}
                    ariaKeyShortcuts={spellingCommand.shortcut?.aria}
                    active={spellcheckEnabled}
                    onClick={onToggleSpellcheck}
                  >
                    <CheckCheck size={19} />
                  </ToolbarButton>
                  <ToolbarButton
                    label="设置校对语言"
                    displayLabel
                    onClick={() => openProofingDialog(editor)}
                  >
                    <Languages size={19} />
                  </ToolbarButton>
                </RibbonGroup>
              )}
              {!suggestionOnly && (
                <RibbonGroup label="批注" priority="high">
                  <ToolbarButton
                    label="添加批注"
                    displayLabel
                    shortcut={insertCommentCommand.shortcut?.label}
                    ariaKeyShortcuts={insertCommentCommand.shortcut?.aria}
                    disabled={!canInsertComment}
                    title={
                      canInsertComment
                        ? `添加批注（${insertCommentCommand.shortcut?.label}）`
                        : '请先选择未批注的文字'
                    }
                    onClick={onInsertComment}
                  >
                    <MessageSquarePlus size={19} />
                  </ToolbarButton>
                  <ToolbarButton
                    label={`查看批注${commentCount ? `（${commentCount}）` : ''}`}
                    displayLabel
                    active={commentsOpen}
                    onClick={onToggleComments}
                  >
                    <MessagesSquare size={19} />
                  </ToolbarButton>
                </RibbonGroup>
              )}
              {(!reviewOnly || suggestionOnly) && (
                <>
                  <RibbonGroup label="修订" priority="high">
                    <ToolbarButton
                      label={suggestionOnly ? '建议模式' : '修订模式'}
                      displayLabel
                      shortcut={trackChangesCommand.shortcut?.label}
                      ariaKeyShortcuts={trackChangesCommand.shortcut?.aria}
                      active={suggestionOnly || trackChanges}
                      disabled={suggestionOnly}
                      title={
                        suggestionOnly
                          ? '建议模式始终记录带身份的文字修订'
                          : undefined
                      }
                      onClick={onToggleTrackChanges}
                    >
                      <FileDiff size={19} />
                    </ToolbarButton>
                    <ToolbarButton
                      label={`查看修订${changeCount ? `（${changeCount}）` : ''}`}
                      displayLabel
                      active={changesOpen}
                      onClick={onToggleChanges}
                    >
                      <ListChecks size={19} />
                    </ToolbarButton>
                  </RibbonGroup>
                  {!suggestionOnly && (
                    <RibbonGroup label="更改" priority="high">
                      <ToolbarButton
                        label="接受修订"
                        displayLabel
                        disabled={actionableChangeIndex === null}
                        title="接受当前修订并转到下一处"
                        onClick={() => decideDocumentChange('accept')}
                      >
                        <Check size={19} />
                      </ToolbarButton>
                      <ToolbarButton
                        label="拒绝修订"
                        displayLabel
                        disabled={actionableChangeIndex === null}
                        title="拒绝当前修订并转到下一处"
                        onClick={() => decideDocumentChange('reject')}
                      >
                        <XCircle size={19} />
                      </ToolbarButton>
                      <ToolbarButton
                        label="上一处修订"
                        displayLabel
                        disabled={previousChangeIndex === null}
                        onClick={() => navigateDocumentChange(-1)}
                      >
                        <ChevronUp size={19} />
                      </ToolbarButton>
                      <ToolbarButton
                        label="下一处修订"
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
                <RibbonGroup label="比较" priority="high">
                  <ToolbarButton
                    label="比较文档"
                    displayLabel
                    title="将另一个版本转换为可接受或拒绝的修订"
                    onClick={() => onOpenComparison('compare')}
                  >
                    <GitCompareArrows size={19} />
                  </ToolbarButton>
                  <ToolbarButton
                    label="合并文档"
                    displayLabel
                    title="合入与当前基线一致的带修订审阅副本"
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
              <RibbonGroup label="文档视图" priority="high">
                <ToolbarButton
                  label="页面视图"
                  displayLabel
                  active={viewMode === 'page'}
                  onClick={() => onViewModeChange('page')}
                >
                  <FileText size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label="网页视图"
                  displayLabel
                  active={viewMode === 'web'}
                  onClick={() => onViewModeChange('web')}
                >
                  <Globe2 size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup label="显示" priority="high">
                <ToolbarButton
                  label="标尺"
                  displayLabel
                  active={showRulers}
                  disabled={viewMode !== 'page'}
                  title={
                    viewMode === 'page'
                      ? '显示或隐藏标尺'
                      : '标尺仅用于页面视图'
                  }
                  onClick={onToggleRulers}
                >
                  <Ruler size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label="导航窗格"
                  displayLabel
                  active={navigationOpen}
                  onClick={onToggleNavigation}
                >
                  <PanelLeftOpen size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label="显示隐藏文字"
                  displayLabel
                  active={showHiddenText}
                  title="显示或隐藏以隐藏文字格式标记的内容"
                  onClick={onToggleHiddenText}
                >
                  <Eye size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup label={`缩放 ${zoom}%`} priority="low">
                <ToolbarButton
                  label="缩小文档"
                  disabled={zoom <= MIN_DOCUMENT_ZOOM}
                  onClick={() => onZoomChange(zoom - 10)}
                >
                  <ZoomOut size={17} />
                </ToolbarButton>
                <ToolbarButton
                  label="缩放至 100%"
                  active={zoom === 100}
                  onClick={() => onZoomChange(100)}
                >
                  100%
                </ToolbarButton>
                <ToolbarButton
                  label="单页"
                  displayLabel
                  onClick={() => onZoomFit('page')}
                >
                  <Scan size={17} />
                </ToolbarButton>
                <ToolbarButton
                  label="页宽"
                  displayLabel
                  onClick={() => onZoomFit('width')}
                >
                  <StretchHorizontal size={17} />
                </ToolbarButton>
                <ToolbarButton
                  label="放大文档"
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
      onClick={onClick}
    >
      {children}
    </WorkOfficeRibbonButton>
  );
}

const RibbonGroup = WorkOfficeRibbonGroup;

function DocumentFieldSelect({
  onInsertField,
}: {
  onInsertField: (kind: WorkDocumentFieldKind) => void;
}) {
  return (
    <OfficeSelect
      ariaLabel="插入页码或日期"
      value=""
      options={[
        { value: '', label: '页码或日期', disabled: true },
        { value: 'page', label: '页码' },
        { value: 'numPages', label: '总页数' },
        { value: 'section', label: '当前节号' },
        { value: 'sectionPages', label: '本节页数' },
        { value: 'date', label: '当前日期' },
        { value: 'time', label: '当前时间' },
      ]}
      onValueChange={(kind) => {
        if (kind) onInsertField(kind);
      }}
    />
  );
}
