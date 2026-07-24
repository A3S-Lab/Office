import type { Editor } from '@tiptap/core';
import Placeholder from '@tiptap/extension-placeholder';
import { EditorContent, useEditor } from '@tiptap/react';
import {
  type CSSProperties,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { WorkspaceContextMenu } from '../../workspace/components/workspace-context-menu';
import { WorkDocumentPreview } from '../components/work-document-pages';
import { WorkEditorLoadingState } from '../components/work-editor-loading-state';
import type { WorkEditorAgentRequest } from '../work-agent-request';
import {
  editorDocumentCaptionTargets,
  insertDocumentCaption,
  insertDocumentCrossReference,
} from '../work-document-caption-editor';
import type { WorkDocumentLayoutFont } from '../work-document-fonts';
import {
  collectDocumentChanges,
  replaceDocumentTextWithTrackedChange,
  type WorkDocumentChangeKind,
} from '../work-document-changes';
import { documentCitationCount } from '../work-document-citation-editor';
import {
  appendDocumentCommentReply,
  collectDocumentCommentAnchors,
  documentCommentViews,
  insertDocumentComment,
  removeDocumentComment,
  removeDocumentCommentRecord,
  retainAnchoredDocumentComments,
  toggleDocumentCommentResolved,
} from '../work-document-comments';
import { createWorkDocumentExtensions } from '../work-document-extensions';
import {
  insertDocumentField,
  refreshDocumentFields,
} from '../work-document-field-editor';
import { documentMargins, millimetersToPixels } from '../work-document-layout';
import {
  createWorkDocumentModel,
  resolveWorkDocumentEditorInput,
} from '../work-document-model';
import { insertDocumentNote } from '../work-document-note-editor';
import {
  documentParagraphIndent,
  setDocumentParagraphIndent,
} from '../work-document-paragraph-formatting';
import {
  documentParagraphTabStops,
  setDocumentParagraphTabStops,
} from '../work-document-tab-stops';
import {
  documentPageMetrics,
  DocumentPagination,
} from '../work-document-pagination';
import {
  documentPageChromeLegacyFields,
  normalizeDocumentPageChrome,
  updateDocumentPageChromeVariant,
} from '../work-document-page-chrome';
import {
  documentInitialSectionLayout,
  normalizeDocumentHtml,
  syncDocumentContentFromHtml,
} from '../work-document-section';
import {
  activeDocumentSection,
  documentSectionById,
  insertDocumentSection,
  mergeDocumentSectionWithPrevious,
  updateActiveDocumentSection,
  updateDocumentSection,
} from '../work-document-section-editor';
import { createWorkId } from '../work-templates';
import type {
  WorkDocumentContent,
  WorkDocumentNode,
  WorkDocumentPageChromeContent,
  WorkDocumentPageChromeVariant,
} from '../work-types';
import { DocumentChangesPanel } from './document-changes-panel';
import { DocumentCitationsPanel } from './document-citations-panel';
import { DocumentCommentsPanel } from './document-comments-panel';
import {
  clampDocumentZoom,
  documentAgentMenuItems,
  documentCurrentPage,
  documentPageCount,
  documentWordCount,
  fileToDataUrl,
  plainTextAsHtml,
} from './document-editor-support';
import { fallbackPaginationPageDescriptor } from './document-editor-pagination';
import { DocumentLayoutPanel } from './document-layout-panel';
import { DocumentPageChromeRichTextEditor } from './document-page-chrome-editor';
import type { DocumentPageChromeEditingPart } from './document-page-chrome-ribbon';
import { DocumentRuler } from './document-ruler';
import { DocumentStatusBar } from './document-status-bar';
import { DocumentToolbar, type DocumentViewMode } from './document-toolbar';
import { DocumentVerticalRuler } from './document-vertical-ruler';
import { OfficeFileInput, useOfficeDialog } from './office-controls';
import {
  type WorkOfficeFileAction,
  WorkOfficePreviewBar,
  WorkOfficeStatusBar,
} from './work-office-chrome';
import { useDocumentPagination } from './use-document-pagination';
import { useDocumentLayoutFonts } from './use-document-layout-fonts';

export interface DocumentEditorProps {
  content: WorkDocumentContent;
  preview: boolean;
  saveStatus?: string;
  kernelWasmUrl?: string;
  layoutFonts?: readonly WorkDocumentLayoutFont[];
  fileActions?: readonly WorkOfficeFileAction[];
  onChange: (content: WorkDocumentContent) => void;
  onAgentRequest?: (request: WorkEditorAgentRequest) => void | Promise<void>;
}

const EMPTY_DOCUMENT_LAYOUT_FONTS: readonly WorkDocumentLayoutFont[] = [];

interface DocumentPageChromeEditingState {
  part: DocumentPageChromeEditingPart;
  sectionId: string;
  variant: WorkDocumentPageChromeVariant;
}

function createTrackedDocumentChange(_kind: WorkDocumentChangeKind) {
  return {
    id: createWorkId('change'),
    author: 'A3S Work 用户',
    date: new Date().toISOString(),
  };
}

export function DocumentEditor({
  content,
  preview,
  saveStatus = '已自动保存',
  kernelWasmUrl,
  layoutFonts = EMPTY_DOCUMENT_LAYOUT_FONTS,
  fileActions,
  onChange,
  onAgentRequest,
}: DocumentEditorProps) {
  const imageInputRef = useRef<HTMLInputElement>(null);
  const pageHeaderRef = useRef<HTMLElement>(null);
  const pageFooterRef = useRef<HTMLElement>(null);
  const reviewSurfaceRef = useRef<HTMLDivElement>(null);
  const contentRef = useRef(content);
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
  const [layoutOpen, setLayoutOpen] = useState(false);
  const [changesOpen, setChangesOpen] = useState(false);
  const [commentsOpen, setCommentsOpen] = useState(false);
  const [citationsOpen, setCitationsOpen] = useState(false);
  const [spellcheckEnabled, setSpellcheckEnabled] = useState(true);
  const [viewMode, setViewMode] = useState<DocumentViewMode>('page');
  const [zoom, setZoom] = useState(90);
  const [pageChromeEditing, setPageChromeEditing] =
    useState<DocumentPageChromeEditingState | null>(null);
  const [pageChromeEditor, setPageChromeEditor] = useState<Editor | null>(null);
  const officeDialog = useOfficeDialog();
  const [agentMenu, setAgentMenu] = useState<{
    x: number;
    y: number;
    selection: string;
    rawSelection: string;
    from: number;
    to: number;
  } | null>(null);
  const [selectionVersion, setSelectionVersion] = useState(0);
  const loadedLayoutFontIds = useDocumentLayoutFonts(layoutFonts);
  contentRef.current = content;
  trackChangesRef.current = Boolean(content.trackChanges);
  const editorExtensions = useMemo(
    () => [
      ...createWorkDocumentExtensions({
        isTracking: () => trackChangesRef.current,
        createChange: createTrackedDocumentChange,
      }),
      Placeholder.configure({ placeholder: '在这里开始输入…' }),
      DocumentPagination,
    ],
    [],
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

  useEffect(() => {
    if (!editor) return;
    const applyEditableState = () => {
      if (!editor.isDestroyed) editor.setEditable(!preview);
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
    if (viewMode === 'page') return;
    setPageChromeEditing(null);
    setPageChromeEditor(null);
  }, [viewMode]);

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
    editor.commands.setContent(editorInput.source, { emitUpdate: false });
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
    enabled: Boolean(editor && !preview && viewMode === 'page'),
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

  if (!editor) {
    return <WorkEditorLoadingState title="正在准备文字编辑器" />;
  }

  const pageCount = pagination.pageCount ?? documentPageCount(editor);
  const currentPage = Math.min(
    pageCount,
    pagination.currentPage ?? documentCurrentPage(editor),
  );
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
  const editingSection = pageChromeEditing
    ? documentSectionById(editor, pageChromeEditing.sectionId)
    : null;
  const editingChrome = pageChromeEditing
    ? normalizeDocumentPageChrome(
        editingSection?.layout.pageChrome,
        editingSection?.layout,
      )[pageChromeEditing.variant]
    : null;
  const visibleChrome = editingChrome ?? currentPageDescriptor.pageChrome;
  const headerChrome =
    pageChromeEditing?.part === 'header'
      ? visibleChrome
      : firstPageDescriptor.pageChrome;
  const footerChrome =
    pageChromeEditing?.part === 'footer'
      ? visibleChrome
      : lastPageDescriptor.pageChrome;
  const finalPageNumber =
    pagination.pages.at(-1)?.pageNumber ??
    Math.max(1, layout.pageNumberStart ?? 1) + pageCount - 1;
  const changes = collectDocumentChanges(editor.state.doc);
  const commentAnchors = collectDocumentCommentAnchors(editor.state.doc);
  const comments = documentCommentViews(content.comments ?? [], commentAnchors);
  const citationCount = documentCitationCount(editor);
  const paragraphIndent = documentParagraphIndent(editor);
  const paragraphTabStops = documentParagraphTabStops(editor);
  const commitComments = (nextComments: WorkDocumentContent['comments']) => {
    const next = { ...contentRef.current, comments: nextComments };
    contentRef.current = next;
    onChange(next);
  };
  const replyToComment = (id: string, text: string) => {
    commitComments(
      appendDocumentCommentReply(contentRef.current.comments ?? [], id, {
        id: createWorkId('comment-reply'),
        author: 'A3S Work 用户',
        date: new Date().toISOString(),
        text,
      }),
    );
  };
  const toggleResolvedComment = (id: string) =>
    commitComments(
      toggleDocumentCommentResolved(contentRef.current.comments ?? [], id),
    );
  const deleteComment = (id: string) => {
    const next = {
      ...contentRef.current,
      comments: removeDocumentCommentRecord(
        contentRef.current.comments ?? [],
        id,
      ),
    };
    contentRef.current = next;
    if (!removeDocumentComment(editor, id)) onChange(next);
  };
  const updateLayout = (next: typeof layout) => {
    updateActiveDocumentSection(editor, next);
  };
  const updateVisiblePageChrome = (
    patch: Partial<WorkDocumentPageChromeContent>,
  ) => {
    const target =
      pageChromeEditing ??
      ({
        part: 'header',
        sectionId: currentPageDescriptor.sectionId,
        variant: currentPageDescriptor.pageChrome.variant,
      } satisfies DocumentPageChromeEditingState);
    const targetSection = documentSectionById(editor, target.sectionId);
    const targetLayout = targetSection?.layout ?? layout;
    const targetPageChrome = normalizeDocumentPageChrome(
      targetLayout.pageChrome,
      targetLayout,
    );
    const nextPageChrome = updateDocumentPageChromeVariant(
      targetPageChrome,
      target.variant,
      patch,
    );
    updateDocumentSection(editor, target.sectionId, {
      ...targetLayout,
      pageChrome: nextPageChrome,
      ...documentPageChromeLegacyFields(nextPageChrome),
    });
  };
  const editPageChrome = (part: DocumentPageChromeEditingPart) => {
    const target = part === 'header' ? firstPageDescriptor : lastPageDescriptor;
    setLayoutOpen(false);
    setPageChromeEditor(null);
    setPageChromeEditing({
      part,
      sectionId: target.sectionId,
      variant: target.pageChrome.variant,
    });
    requestAnimationFrame(() => {
      const element =
        part === 'header' ? pageHeaderRef.current : pageFooterRef.current;
      element?.scrollIntoView?.({ block: 'center' });
    });
  };
  const closePageChrome = () => {
    setPageChromeEditing(null);
    setPageChromeEditor(null);
    requestAnimationFrame(() => {
      if (!editor.isDestroyed) editor.commands.focus();
    });
  };
  const toggleVisiblePageNumber = () =>
    updateVisiblePageChrome({
      showPageNumber: !visibleChrome.showPageNumber,
    });
  const addSection = () => {
    insertDocumentSection(editor, layout.breakAfter);
  };

  if (preview) {
    return (
      <section className="work-document-editor preview">
        <WorkOfficePreviewBar
          ariaLabel="文字预览工具"
          label="只读预览"
          detail={`${pageCount} 页`}
          fileActions={fileActions}
          className="work-document-ribbon"
        />
        <WorkDocumentPreview content={content} />
        <WorkOfficeStatusBar className="work-document-footer">
          <output aria-label="页数状态">{pageCount} 页</output>
          <output aria-label="分节状态">{section?.count ?? 1} 节</output>
        </WorkOfficeStatusBar>
      </section>
    );
  }

  return (
    <section className="work-document-editor">
      <OfficeFileInput
        ref={imageInputRef}
        accept="image/*"
        aria-label="插入文档图片"
        onChange={(event) => {
          const file = event.target.files?.[0];
          event.target.value = '';
          if (!file) return;
          if (file.size > 8 * 1024 * 1024) {
            void officeDialog.notice({
              title: '图片过大',
              description: '单张图片不能超过 8 MiB。',
            });
            return;
          }
          void fileToDataUrl(file).then((src) =>
            editor
              .chain()
              .focus()
              .setImage({ src, alt: file.name, title: file.name })
              .run(),
          );
        }}
      />
      <DocumentToolbar
        editor={editor}
        fileActions={fileActions}
        layoutOpen={layoutOpen}
        showPageNumbers={visibleChrome.showPageNumber}
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
        onToggleLayout={() => {
          setPageChromeEditing(null);
          setPageChromeEditor(null);
          setLayoutOpen((value) => !value);
        }}
        onToggleSpellcheck={() => setSpellcheckEnabled((value) => !value)}
        onViewModeChange={setViewMode}
        onZoomChange={(nextZoom) => setZoom(clampDocumentZoom(nextZoom))}
        onTogglePageNumbers={toggleVisiblePageNumber}
        onInsertSection={addSection}
        onInsertNote={(kind) => insertDocumentNote(editor, kind)}
        onInsertCaption={(kind) =>
          void officeDialog
            .prompt({
              title: kind === 'figure' ? '图片题注文字' : '表格题注文字',
              confirmLabel: '插入题注',
            })
            .then((title) => {
              if (title !== null) insertDocumentCaption(editor, kind, title);
            })
        }
        onInsertCrossReference={() =>
          void (async () => {
            const targets = editorDocumentCaptionTargets(editor);
            if (!targets.length) {
              await officeDialog.notice({
                title: '还没有题注',
                description: '请先插入图片或表格题注。',
              });
              return;
            }
            const choice = await officeDialog.prompt({
              title: '引用题注',
              description: targets
                .map((target) => `${target.display} ${target.title}`.trim())
                .join('；'),
              initialValue: targets[0].display,
              confirmLabel: '插入引用',
            });
            if (choice === null) return;
            const target = targets.find(
              (item) =>
                item.display === choice.trim() ||
                `${item.display} ${item.title}`.trim() === choice.trim(),
            );
            if (!target) {
              await officeDialog.notice({
                title: '没有找到题注',
                description: '请选择现有的图片或表格题注。',
              });
              return;
            }
            insertDocumentCrossReference(editor, target);
          })()
        }
        citationsOpen={citationsOpen}
        citationSourceCount={content.bibliography?.sources.length ?? 0}
        onToggleCitations={() => setCitationsOpen((value) => !value)}
        onInsertField={(kind) =>
          insertDocumentField(editor, kind, contentRef.current)
        }
        onRefreshFields={() => {
          refreshDocumentFields(editor, contentRef.current);
        }}
        onInsertComment={() =>
          void (async () => {
            if (editor.state.selection.empty) {
              await officeDialog.notice({
                title: '无法添加批注',
                description: '请先选择要批注的文字。',
              });
              return;
            }
            const text = await officeDialog.prompt({
              title: '批注内容',
              multiline: true,
              confirmLabel: '添加批注',
            });
            if (!text?.trim()) return;
            const comment = {
              id: createWorkId('comment'),
              author: 'A3S Work 用户',
              date: new Date().toISOString(),
              text: text.trim(),
              resolved: false,
            };
            const previous = contentRef.current;
            contentRef.current = {
              ...previous,
              comments: [...(previous.comments ?? []), comment],
            };
            if (!insertDocumentComment(editor, comment.id)) {
              contentRef.current = previous;
              await officeDialog.notice({
                title: '无法添加批注',
                description: '所选文字已经包含批注，请选择其他文字。',
              });
              return;
            }
            setCommentsOpen(true);
          })()
        }
        commentsOpen={commentsOpen}
        commentCount={comments.length}
        onToggleComments={() => setCommentsOpen((value) => !value)}
        trackChanges={Boolean(content.trackChanges)}
        changesOpen={changesOpen}
        changeCount={changes.length}
        onRibbonTabChange={(tab) => {
          if (tab !== 'page') setLayoutOpen(false);
          if (tab !== 'references') setCitationsOpen(false);
          if (tab !== 'review') {
            setCommentsOpen(false);
            setChangesOpen(false);
          }
        }}
        onToggleTrackChanges={() => {
          const trackChanges = !trackChangesRef.current;
          trackChangesRef.current = trackChanges;
          onChange({ ...contentRef.current, trackChanges });
        }}
        onToggleChanges={() => setChangesOpen((value) => !value)}
        onReplaceText={(from, to, replacement) => {
          editor.commands.focus();
          if (trackChangesRef.current) {
            return replaceDocumentTextWithTrackedChange(
              editor,
              from,
              to,
              replacement,
              createTrackedDocumentChange,
            );
          }
          return editor
            .chain()
            .focus()
            .setTextSelection({ from, to })
            .insertContent(plainTextAsHtml(replacement))
            .run();
        }}
      />
      {citationsOpen && (
        <DocumentCitationsPanel
          editor={editor}
          content={content}
          onChange={(next) => {
            contentRef.current = next;
            onChange(next);
          }}
          onClose={() => setCitationsOpen(false)}
        />
      )}
      {changesOpen && (
        <DocumentChangesPanel
          editor={editor}
          changes={changes}
          onClose={() => setChangesOpen(false)}
        />
      )}
      {layoutOpen && section && (
        <DocumentLayoutPanel
          layout={layout}
          sectionIndex={section.index}
          sectionCount={section.count}
          onChange={updateLayout}
          onInsertSection={addSection}
          onMergeSection={() => mergeDocumentSectionWithPrevious(editor)}
        />
      )}
      <div className={`work-document-scroll ${viewMode}`}>
        <div
          ref={reviewSurfaceRef}
          className={`work-document-review-surface${commentsOpen ? ' comments-open' : ''}`}
        >
          <div
            className={`work-document-page-stage ${layout.pageSize} ${layout.orientation} ${viewMode}`}
            data-testid="document-page-stage"
            style={
              { '--work-document-zoom': String(zoom / 100) } as CSSProperties
            }
          >
            {viewMode === 'page' && (
              <DocumentRuler
                layout={layout}
                paragraphIndent={paragraphIndent}
                tabStops={paragraphTabStops}
                onParagraphIndentChange={(nextParagraphIndent) =>
                  setDocumentParagraphIndent(editor, nextParagraphIndent, {
                    restoreFocus: false,
                  })
                }
                onTabStopsChange={(nextTabStops) =>
                  setDocumentParagraphTabStops(editor, nextTabStops, {
                    restoreFocus: false,
                  })
                }
                onLayoutChange={updateLayout}
              />
            )}
            <div className="work-document-page-frame">
              {viewMode === 'page' && (
                <DocumentVerticalRuler
                  layout={layout}
                  onLayoutChange={updateLayout}
                />
              )}
              <article
                className={`work-document-page ${layout.pageSize} ${layout.orientation}${pagination.pageCount ? ' paginated' : ''}${pageChromeEditing ? ' page-chrome-editing' : ''}${layoutOpen ? ' page-chrome-panel-open' : ''}`}
                aria-label={preview ? '文字预览' : '文字页面'}
                style={
                  {
                    padding: `${marginPixels.top}px ${marginPixels.right}px ${marginPixels.bottom}px ${marginPixels.left}px`,
                    '--work-document-page-margin-left': `${marginPixels.left}px`,
                    '--work-document-page-margin-right': `${marginPixels.right}px`,
                    '--work-document-page-header-offset': `${Math.max(
                      8,
                      (marginPixels.top - 28) / 2,
                    )}px`,
                    '--work-document-page-footer-offset': `${Math.max(
                      8,
                      (marginPixels.bottom - 28) / 2,
                    )}px`,
                  } as CSSProperties
                }
              >
                {viewMode === 'page' && (
                  <header
                    ref={pageHeaderRef}
                    className={`work-document-page-header${headerChrome.headerHtml ? ' has-content' : ' empty'}${pageChromeEditing?.part === 'header' ? ' editing' : ''}`}
                    data-document-page-chrome={
                      firstPageDescriptor.pageChrome.variant
                    }
                  >
                    {pageChromeEditing?.part === 'header' ? (
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
                    ) : (
                      <>
                        {headerChrome.headerHtml ? (
                          <div
                            className="work-document-page-chrome-html"
                            dangerouslySetInnerHTML={{
                              __html: headerChrome.headerHtml,
                            }}
                          />
                        ) : (
                          <span className="work-document-page-chrome-placeholder">
                            页眉
                          </span>
                        )}
                        <button
                          type="button"
                          className="work-document-page-chrome-activate"
                          aria-label="编辑页眉"
                          title="编辑页眉"
                          onClick={() => editPageChrome('header')}
                        />
                      </>
                    )}
                  </header>
                )}
                <section
                  className={`work-document-editable ${viewMode}`}
                  aria-label="文档内容编辑区域"
                  onDoubleClick={() => {
                    if (pageChromeEditing) closePageChrome();
                  }}
                  onContextMenu={(event) => {
                    if (!onAgentRequest) return;
                    const { from, to, empty } = editor.state.selection;
                    if (empty) return;
                    const rawSelection = editor.state.doc.textBetween(
                      from,
                      to,
                      '\n',
                    );
                    const selection = rawSelection.trim();
                    if (!selection) return;
                    event.preventDefault();
                    setAgentMenu({
                      x: event.clientX,
                      y: event.clientY,
                      selection,
                      rawSelection,
                      from,
                      to,
                    });
                  }}
                >
                  <EditorContent editor={editor} />
                </section>
                {viewMode === 'page' && (
                  <footer
                    ref={pageFooterRef}
                    className={`work-document-page-footer${footerChrome.footerHtml ? ' has-content' : ' empty'}${pageChromeEditing?.part === 'footer' ? ' editing' : ''}`}
                    data-document-page-chrome={
                      lastPageDescriptor.pageChrome.variant
                    }
                  >
                    <div className="work-document-page-footer-content">
                      {pageChromeEditing?.part === 'footer' ? (
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
                      ) : (
                        <>
                          {footerChrome.footerHtml ? (
                            <div
                              className="work-document-page-chrome-html"
                              dangerouslySetInnerHTML={{
                                __html: footerChrome.footerHtml,
                              }}
                            />
                          ) : (
                            <span className="work-document-page-chrome-placeholder">
                              页脚
                            </span>
                          )}
                          <button
                            type="button"
                            className="work-document-page-chrome-activate"
                            aria-label="编辑页脚"
                            title="编辑页脚"
                            onClick={() => editPageChrome('footer')}
                          />
                        </>
                      )}
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
          {commentsOpen && (
            <DocumentCommentsPanel
              editor={editor}
              comments={comments}
              surfaceRef={reviewSurfaceRef}
              onReply={replyToComment}
              onToggleResolved={toggleResolvedComment}
              onDelete={deleteComment}
              onClose={() => setCommentsOpen(false)}
            />
          )}
        </div>
      </div>
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
        wordCount={documentWordCount(editor.getText())}
        zoom={zoom}
        onSpellcheckChange={setSpellcheckEnabled}
        onViewModeChange={setViewMode}
        onZoomChange={setZoom}
      />
      {agentMenu && onAgentRequest && (
        <WorkspaceContextMenu
          label="选中文本 AI 操作"
          x={agentMenu.x}
          y={agentMenu.y}
          items={documentAgentMenuItems(agentMenu.selection, onAgentRequest, {
            target: {
              id: 'document-selection',
              label: '选中文本',
              before: agentMenu.rawSelection,
            },
            apply: (changes) => {
              const change = changes.find(
                (candidate) => candidate.id === 'document-selection',
              );
              if (!change) return { appliedTargetIds: [], conflicts: [] };
              const current = editor.state.doc.textBetween(
                agentMenu.from,
                agentMenu.to,
                '\n',
              );
              if (current !== agentMenu.rawSelection) {
                return {
                  appliedTargetIds: [],
                  conflicts: [
                    {
                      targetId: change.id,
                      label: change.label,
                      message: '选中文本在建议生成后已发生变化。',
                    },
                  ],
                };
              }
              const applied = trackChangesRef.current
                ? replaceDocumentTextWithTrackedChange(
                    editor,
                    agentMenu.from,
                    agentMenu.to,
                    change.after,
                    createTrackedDocumentChange,
                  )
                : editor
                    .chain()
                    .focus()
                    .setTextSelection({
                      from: agentMenu.from,
                      to: agentMenu.to,
                    })
                    .insertContent(plainTextAsHtml(change.after))
                    .run();
              return applied
                ? { appliedTargetIds: [change.id], conflicts: [] }
                : {
                    appliedTargetIds: [],
                    conflicts: [
                      {
                        targetId: change.id,
                        label: change.label,
                        message: '编辑器无法替换当前选区。',
                      },
                    ],
                  };
            },
          })}
          onClose={() => setAgentMenu(null)}
        />
      )}
      {officeDialog.dialog}
    </section>
  );
}
