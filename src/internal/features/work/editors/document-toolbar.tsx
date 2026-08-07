import type { Editor } from '@tiptap/core';
import {
  BookOpen,
  CheckCheck,
  Columns3,
  FileDiff,
  FilePlus2,
  FileText,
  Globe2,
  Hash,
  Image as ImageIcon,
  Link2,
  ListChecks,
  MessageSquarePlus,
  MessagesSquare,
  Palette,
  PanelBottomOpen,
  PanelLeftOpen,
  PanelTopOpen,
  RefreshCw,
  Redo2,
  Ruler,
  Settings2,
  Table2,
  Undo2,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import { type ReactNode, useCallback, useEffect, useState } from 'react';
import type { WorkDocumentCaptionKind } from '../work-document-captions';
import type { WorkDocumentFieldKind } from '../work-document-fields';
import {
  DOCUMENT_LINK_VALIDATION_MESSAGE,
  normalizeDocumentHref,
} from '../work-document-links';
import type { WorkDocumentNoteKind } from '../work-document-notes';
import type { WorkDocumentLayoutFont } from '../work-document-fonts';
import {
  MAX_DOCUMENT_ZOOM,
  MIN_DOCUMENT_ZOOM,
} from './document-editor-support';
import {
  documentPageChromeRibbonTab,
  documentPictureRibbonTab,
  documentRibbonTabs,
  documentTableRibbonTabs,
  type DocumentRibbonTabId,
  getDocumentCommandDefinition,
} from './document-command-catalog';
import { DocumentHomeRibbon } from './document-home-ribbon';
import {
  type DocumentPageChromeEditingPart,
  DocumentPageChromeRibbon,
} from './document-page-chrome-ribbon';
import { DocumentParagraphSpacingPopover } from './document-paragraph-spacing-popover';
import { DocumentPaginationPopover } from './document-pagination-popover';
import { DocumentPictureRibbon } from './document-picture-ribbon';
import { DocumentTableInsertPopover } from './document-table-insert-popover';
import {
  DocumentTableDesignRibbon,
  DocumentTableLayoutRibbon,
} from './document-table-ribbon';
import type { DocumentFindReplaceMode } from './document-find-replace-panel';
import {
  OfficeColorPicker,
  OfficeSelect,
  useOfficeDialog,
} from './office-controls';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import {
  type WorkOfficeFileAction,
  WorkOfficeRibbon,
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

export type DocumentViewMode = 'page' | 'web';

interface DocumentToolbarProps {
  editor: Editor;
  layoutOpen: boolean;
  layoutFonts?: readonly WorkDocumentLayoutFont[];
  navigationOpen: boolean;
  pageColor: string;
  showPageNumbers: boolean;
  showRulers: boolean;
  spellcheckEnabled: boolean;
  viewMode: DocumentViewMode;
  zoom: number;
  pageChromeEditor: Editor | null;
  pageChromeEditingPart: DocumentPageChromeEditingPart | null;
  pageChromeShowPageNumber: boolean;
  onRequestImage: () => void;
  onPageChromeEditingPartChange: (part: DocumentPageChromeEditingPart) => void;
  onClosePageChrome: () => void;
  onTogglePageChromePageNumber: () => void;
  onToggleLayout: () => void;
  onToggleNavigation: () => void;
  onTogglePageNumbers: () => void;
  onToggleRulers: () => void;
  onPageColorChange: (color: string) => void;
  onToggleSpellcheck: () => void;
  onViewModeChange: (mode: DocumentViewMode) => void;
  onZoomChange: (zoom: number) => void;
  onInsertSection: () => void;
  onInsertNote: (kind: WorkDocumentNoteKind) => void;
  onInsertCaption: (kind: WorkDocumentCaptionKind) => void;
  onInsertCrossReference: () => void;
  citationsOpen: boolean;
  citationSourceCount: number;
  onToggleCitations: () => void;
  onInsertField: (kind: WorkDocumentFieldKind) => void;
  onRefreshFields: () => void;
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
  onOpenFindReplace: (mode: DocumentFindReplaceMode) => void;
}

export function DocumentToolbar({
  editor,
  layoutOpen,
  layoutFonts = [],
  navigationOpen,
  pageColor,
  showPageNumbers,
  showRulers,
  spellcheckEnabled,
  viewMode,
  zoom,
  pageChromeEditor,
  pageChromeEditingPart,
  pageChromeShowPageNumber,
  onRequestImage,
  onPageChromeEditingPartChange,
  onClosePageChrome,
  onTogglePageChromePageNumber,
  onToggleLayout,
  onToggleNavigation,
  onTogglePageNumbers,
  onToggleRulers,
  onPageColorChange,
  onToggleSpellcheck,
  onViewModeChange,
  onZoomChange,
  onInsertSection,
  onInsertNote,
  onInsertCaption,
  onInsertCrossReference,
  citationsOpen,
  citationSourceCount,
  onToggleCitations,
  onInsertField,
  onRefreshFields,
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
  onOpenFindReplace,
}: DocumentToolbarProps) {
  const [activeTab, setActiveTab] = useState<DocumentRibbonTabId>('home');
  const officeDialog = useOfficeDialog();
  const prompt = officeDialog.prompt;
  const imageSelected = editor.isActive('image');
  const tableSelected = editor.isActive('table');
  const hasRefreshableFields = documentHasRefreshableFields(editor);
  const undoCommand = getDocumentCommandDefinition('undo');
  const redoCommand = getDocumentCommandDefinition('redo');
  const ribbonTabs = pageChromeEditor
    ? [...documentRibbonTabs, documentPageChromeRibbonTab]
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
      description: '为当前选中的文字设置跳转地址。',
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
  useEffect(() => {
    setActiveTab((current) => {
      if (pageChromeEditor) return 'pageChrome';
      if (imageSelected) return 'picture';
      if (tableSelected) {
        return current === 'tableDesign' || current === 'tableLayout'
          ? current
          : 'tableDesign';
      }
      return current === 'picture' ||
        current === 'tableDesign' ||
        current === 'tableLayout' ||
        current === 'pageChrome'
        ? 'home'
        : current;
    });
  }, [imageSelected, pageChromeEditor, tableSelected]);

  useEffect(() => {
    let editorDom: HTMLElement | null = null;
    let root: HTMLElement | null = null;
    const onKeyDown = (event: KeyboardEvent) => {
      if (
        event.defaultPrevented ||
        event.repeat ||
        event.altKey ||
        isOfficeShortcutBlocked(event.target) ||
        !(event.metaKey || event.ctrlKey)
      ) {
        return;
      }
      const key = event.key.toLowerCase();
      const insideEditor =
        event.target instanceof Node &&
        Boolean(editorDom?.contains(event.target));
      const documentHistoryTarget =
        insideEditor || !isDocumentNativeTextUndoTarget(event.target);
      if (documentHistoryTarget && key === 'z') {
        event.preventDefault();
        const command = event.shiftKey
          ? editor.chain().focus().redo()
          : editor.chain().focus().undo();
        command.run();
        return;
      }
      if (documentHistoryTarget && key === 'y' && !event.shiftKey) {
        event.preventDefault();
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
      root?.removeEventListener('keydown', onKeyDown);
      root = null;
      editorDom = null;
    };
    const attach = () => {
      detach();
      if (editor.isDestroyed) return;
      editorDom = editor.view.dom;
      root = editorDom.closest<HTMLElement>('.work-document-editor');
      root?.addEventListener('keydown', onKeyDown);
    };
    attach();
    editor.on('mount', attach);
    editor.on('unmount', detach);
    return () => {
      editor.off('mount', attach);
      editor.off('unmount', detach);
      detach();
    };
  }, [editor, onOpenFindReplace, toggleLink]);

  return (
    <>
      <WorkOfficeRibbon
        ariaLabel="文字功能区"
        tabs={ribbonTabs}
        defaultTab="home"
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
        collapsible
        fileActions={fileActions}
        quickAccessActions={[
          {
            id: undoCommand.id,
            label: undoCommand.label,
            icon: <Undo2 size={15} />,
            shortcut: undoCommand.shortcut?.label,
            ariaKeyShortcuts: undoCommand.shortcut?.aria,
            disabled: editor.isDestroyed || !editor.can().chain().undo().run(),
            onSelect: () => {
              editor.chain().focus().undo().run();
            },
          },
          {
            id: redoCommand.id,
            label: redoCommand.label,
            icon: <Redo2 size={15} />,
            shortcut: redoCommand.shortcut?.label,
            ariaKeyShortcuts: redoCommand.shortcut?.aria,
            disabled: editor.isDestroyed || !editor.can().chain().redo().run(),
            onSelect: () => {
              editor.chain().focus().redo().run();
            },
          },
        ]}
        className="work-document-ribbon"
        toolbarClassName="document-toolbar"
        panels={{
          home: (
            <DocumentHomeRibbon
              editor={editor}
              findReplaceMode={findReplaceMode}
              layoutFonts={layoutFonts}
              onFindText={(replace) =>
                onOpenFindReplace(replace ? 'replace' : 'find')
              }
            />
          ),
          insert: (
            <>
              <RibbonGroup label="插图与表格">
                <ToolbarButton
                  label="插入图片"
                  displayLabel
                  onClick={onRequestImage}
                >
                  <ImageIcon size={19} />
                </ToolbarButton>
                <DocumentTableInsertPopover editor={editor} />
              </RibbonGroup>
              <RibbonGroup label="页面">
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
              </RibbonGroup>
              <RibbonGroup label="链接">
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
              </RibbonGroup>
              <RibbonGroup label="页码和日期">
                <DocumentFieldSelect onInsertField={onInsertField} />
              </RibbonGroup>
            </>
          ),
          page: (
            <>
              <RibbonGroup label="段落">
                <DocumentParagraphSpacingPopover editor={editor} />
                <DocumentPaginationPopover editor={editor} />
              </RibbonGroup>
              <RibbonGroup label="页面设置">
                <ToolbarButton
                  label="页面设置"
                  displayLabel
                  active={layoutOpen}
                  onClick={onToggleLayout}
                >
                  <Settings2 size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label="显示页码"
                  displayLabel
                  active={showPageNumbers}
                  onClick={onTogglePageNumbers}
                >
                  <Hash size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup label="页面背景">
                <OfficeColorPicker
                  ariaLabel="页面颜色"
                  className="work-document-page-color-picker"
                  triggerLabel="页面颜色"
                  triggerIcon={<Palette size={18} />}
                  value={pageColor}
                  onValueChange={onPageColorChange}
                />
              </RibbonGroup>
              <RibbonGroup label="分隔符">
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
                <ToolbarButton
                  label="插入分节符"
                  displayLabel
                  onClick={onInsertSection}
                >
                  <Columns3 size={19} />
                </ToolbarButton>
              </RibbonGroup>
            </>
          ),
          references: (
            <>
              <RibbonGroup label="脚注与尾注">
                <ToolbarButton
                  label="插入脚注"
                  displayLabel
                  onClick={() => onInsertNote('footnote')}
                >
                  <span className="work-ribbon-glyph">¹</span>
                </ToolbarButton>
                <ToolbarButton
                  label="插入尾注"
                  displayLabel
                  onClick={() => onInsertNote('endnote')}
                >
                  <span className="work-ribbon-glyph">ⅰ</span>
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup label="题注">
                <ToolbarButton
                  label="插入图片题注"
                  displayLabel
                  onClick={() => onInsertCaption('figure')}
                >
                  <ImageIcon size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label="插入表格题注"
                  displayLabel
                  onClick={() => onInsertCaption('table')}
                >
                  <Table2 size={19} />
                </ToolbarButton>
                <ToolbarButton
                  label="插入交叉引用"
                  displayLabel
                  onClick={onInsertCrossReference}
                >
                  <Link2 size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup label="文献">
                <ToolbarButton
                  label={`文献库${citationSourceCount ? `（${citationSourceCount}）` : ''}`}
                  displayLabel
                  active={citationsOpen}
                  onClick={onToggleCitations}
                >
                  <BookOpen size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup label="更新">
                <ToolbarButton
                  label="更新页码和日期"
                  displayLabel
                  disabled={!hasRefreshableFields}
                  title={
                    hasRefreshableFields
                      ? '更新页码和日期'
                      : '文档中没有可更新的页码或日期'
                  }
                  onClick={onRefreshFields}
                >
                  <RefreshCw size={19} />
                </ToolbarButton>
              </RibbonGroup>
            </>
          ),
          review: (
            <>
              <RibbonGroup label="校对">
                <ToolbarButton
                  label="拼写检查"
                  displayLabel
                  active={spellcheckEnabled}
                  onClick={onToggleSpellcheck}
                >
                  <CheckCheck size={19} />
                </ToolbarButton>
              </RibbonGroup>
              <RibbonGroup label="批注">
                <ToolbarButton
                  label="添加批注"
                  displayLabel
                  disabled={!canInsertComment}
                  title={canInsertComment ? '添加批注' : '请先选择未批注的文字'}
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
              <RibbonGroup label="修订">
                <ToolbarButton
                  label="修订模式"
                  displayLabel
                  active={trackChanges}
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
            </>
          ),
          view: (
            <>
              <RibbonGroup label="显示">
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
              </RibbonGroup>
              <RibbonGroup label="文档视图">
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
              <RibbonGroup label={`缩放 ${zoom}%`}>
                <ToolbarButton
                  label="缩小文档"
                  disabled={zoom <= MIN_DOCUMENT_ZOOM}
                  onClick={() => onZoomChange(zoom - 10)}
                >
                  <ZoomOut size={17} />
                </ToolbarButton>
                {[75, 100, 125].map((value) => (
                  <ToolbarButton
                    key={value}
                    label={`缩放至 ${value}%`}
                    active={zoom === value}
                    onClick={() => onZoomChange(value)}
                  >
                    {value}%
                  </ToolbarButton>
                ))}
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
                onClose={onClosePageChrome}
              />
            ) : null,
        }}
      />
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

function documentHasRefreshableFields(editor: Editor): boolean {
  let found = false;
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'documentField') return !found;
    found = true;
    return false;
  });
  return found;
}
