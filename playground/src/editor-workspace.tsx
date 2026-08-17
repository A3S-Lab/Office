import {
  type DocumentContent,
  type DocumentSelectionContext,
  downloadArtifact,
  downloadArtifactPdf,
  type EditorAgentRequest,
  type GetDocumentSelectionMenuItems,
  type GetMarkdownSelectionMenuItems,
  type MarkdownContent,
  type MarkdownSelectionContext,
  type OfficeArtifact,
  type OfficeArtifactContent,
  type PresentationContent,
  readSourceBlob,
  registerSourceBlob,
  type SpreadsheetContent,
} from '@a3s-lab/office/core';
import {
  DocumentEditor,
  MarkdownEditor,
  PdfViewer,
  PresentationEditor,
  SpreadsheetEditor,
} from '@a3s-lab/office/react';
import {
  ArrowLeft,
  Download,
  Eye,
  MessageSquareText,
  PanelLeftOpen,
  Pencil,
  SendHorizontal,
  Sparkles,
  TriangleAlert,
  X,
} from 'lucide-react';
import {
  type KeyboardEvent as ReactKeyboardEvent,
  type PointerEvent as ReactPointerEvent,
  useCallback,
  useEffect,
  useRef,
  useState,
} from 'react';
import { useDialogFocusScope } from '../../src/internal/design-system/primitives/overlay/dialog-focus-scope';
import { usePlaygroundCollaborationPresenceFixture } from './collaboration-presence-fixture';
import {
  type PlaygroundNativeDocumentSuggestionStage,
  usePlaygroundDocumentSuggestionFixture,
} from './collaboration-suggestion-fixture';
import { FileKindIcon, fileKindExtension, fileKindLabel } from './file-kind';
import {
  type PlaygroundPdfAnnotationStage,
  usePlaygroundPdfCollaborationFixture,
} from './pdf-collaboration-fixture';
import type { NoticeTone } from './playground-types';
import {
  type PlaygroundPresentationElementStage,
  usePlaygroundPresentationCollaborationFixture,
} from './presentation-collaboration-fixture';
import {
  type PlaygroundSpreadsheetCellStage,
  usePlaygroundSpreadsheetCollaborationFixture,
} from './spreadsheet-collaboration-fixture';

const assistantMinimumWidth = 340;
const assistantMaximumWidth = 680;

export function EditorWorkspace({
  artifact,
  collaborationDemo,
  suggestionDemo,
  sidebarOpen,
  assistantModal,
  assistantOpen,
  assistantWidth,
  lastAgentRequest,
  onOpenSidebar,
  onBack,
  onChange,
  onRename,
  onTouch,
  onToggleAssistant,
  onAssistantWidthChange,
  onAgentRequest,
  onNotice,
}: {
  artifact: OfficeArtifact;
  collaborationDemo: boolean;
  suggestionDemo: boolean;
  sidebarOpen: boolean;
  assistantModal: boolean;
  assistantOpen: boolean;
  assistantWidth: number;
  lastAgentRequest: EditorAgentRequest | null;
  onOpenSidebar: () => void;
  onBack: () => void;
  onChange: (content: OfficeArtifactContent) => void;
  onRename: (title: string) => void;
  onTouch: () => void;
  onToggleAssistant: () => void;
  onAssistantWidthChange: (width: number) => void;
  onAgentRequest: (request: EditorAgentRequest) => void;
  onNotice: (message: string, tone?: NoticeTone) => void;
}) {
  const [preview, setPreview] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [assistantQuestion, setAssistantQuestion] =
    useState<PlaygroundAssistantQuestionDraft | null>(null);
  const extension = fileKindExtension(artifact.kind);
  const e2eFixture =
    typeof window !== 'undefined' &&
    new URLSearchParams(window.location.search).get('e2e');
  const loadPdf = useCallback(() => readSourceBlob(artifact), [artifact]);
  const controlledReviewFixture = e2eFixture === 'word-review-conflict';
  const controlledReviewFixtureReady =
    artifact.content.type === 'document' &&
    artifact.content.html.includes(CONTROLLED_REVIEW_COMMENT_ID);
  const collaborationPresenceFixtureEnabled =
    artifact.content.type === 'document' &&
    (collaborationDemo || e2eFixture === 'collaboration-presence');
  const collaborationPresenceFixture =
    usePlaygroundCollaborationPresenceFixture(
      collaborationPresenceFixtureEnabled &&
        artifact.content.type === 'document'
        ? { artifactId: artifact.id, content: artifact.content }
        : undefined,
    );
  const documentSuggestionFixtureEnabled =
    artifact.content.type === 'document' &&
    (suggestionDemo || e2eFixture === 'collaboration-document-suggestions');
  const documentSuggestionFixture = usePlaygroundDocumentSuggestionFixture(
    documentSuggestionFixtureEnabled,
  );
  const pendingSuggestion = Boolean(
    documentSuggestionFixture?.content.html.includes('data-document-change'),
  );
  const suggestionDecisionCount =
    documentSuggestionFixture?.content.changeDecisions?.length ?? 0;
  const pdfCollaborationFixtureEnabled =
    e2eFixture === 'collaboration-pdf-annotations' &&
    artifact.content.type === 'pdf';
  const pdfCollaborationFixture = usePlaygroundPdfCollaborationFixture(
    pdfCollaborationFixtureEnabled
      ? {
          artifactId: artifact.id,
          loadSource: loadPdf,
          pageCount: 4,
        }
      : undefined,
  );
  const spreadsheetCollaborationFixtureEnabled =
    e2eFixture === 'collaboration-spreadsheet-cells' &&
    artifact.content.type === 'spreadsheet';
  const spreadsheetCollaborationFixture =
    usePlaygroundSpreadsheetCollaborationFixture(
      spreadsheetCollaborationFixtureEnabled,
    );
  const presentationCollaborationFixtureEnabled =
    e2eFixture === 'collaboration-presentation-elements' &&
    artifact.content.type === 'presentation';
  const presentationCollaborationFixture =
    usePlaygroundPresentationCollaborationFixture(
      presentationCollaborationFixtureEnabled,
    );

  const handleAgentRequest = useCallback(
    (request: EditorAgentRequest) => {
      setAssistantQuestion(null);
      onAgentRequest(request);
    },
    [onAgentRequest],
  );
  const beginAssistantQuestion = useCallback(
    (draft: PlaygroundAssistantQuestionDraft) => {
      setAssistantQuestion(draft);
      if (!assistantOpen) onToggleAssistant();
    },
    [assistantOpen, onToggleAssistant],
  );

  const exportArtifact = async () => {
    if (exporting) return;
    setExporting(true);
    try {
      await downloadArtifact(artifact);
      onNotice(
        `${artifact.title}.${extension.toLocaleLowerCase()} 已下载`,
        'success',
      );
    } catch (error) {
      onNotice(
        error instanceof Error ? error.message : '导出失败，请重试',
        'danger',
      );
    } finally {
      setExporting(false);
    }
  };

  const exportDocumentPdf = async () => {
    if (exporting || artifact.content.type !== 'document') return;
    setExporting(true);
    try {
      await downloadArtifactPdf(artifact);
      onNotice(`${artifact.title}.pdf 已下载`, 'success');
    } catch (error) {
      onNotice(
        error instanceof Error ? error.message : 'PDF 导出失败，请重试',
        'danger',
      );
    } finally {
      setExporting(false);
    }
  };

  const savePdf = useCallback(
    async (pdf: Blob) => {
      registerSourceBlob(artifact.id, pdf);
      onTouch();
      onNotice('PDF 批注已保存到当前浏览器会话', 'success');
      return true;
    },
    [artifact.id, onNotice, onTouch],
  );
  const getDocumentSelectionMenuItems =
    useCallback<GetDocumentSelectionMenuItems>(
      () =>
        createPlaygroundSelectionMenuItems<DocumentSelectionContext>(
          handleAgentRequest,
          beginAssistantQuestion,
          onNotice,
        ),
      [beginAssistantQuestion, handleAgentRequest, onNotice],
    );
  const getMarkdownSelectionMenuItems =
    useCallback<GetMarkdownSelectionMenuItems>(
      () =>
        createPlaygroundSelectionMenuItems<MarkdownSelectionContext>(
          handleAgentRequest,
          beginAssistantQuestion,
          onNotice,
        ),
      [beginAssistantQuestion, handleAgentRequest, onNotice],
    );

  return (
    <section className={`work-editor-shell ${artifact.kind}`}>
      <div className="playground-editor-row">
        <section
          className="playground-editor-host"
          aria-label={`${fileKindLabel(artifact.kind)}编辑器`}
        >
          <header
            className="work-editor-header playground-editor-header"
            role="toolbar"
            aria-label="文件命令栏"
          >
            <div className="playground-editor-command-start">
              {!sidebarOpen && (
                <button
                  type="button"
                  className="playground-icon-button editor-sidebar-open"
                  aria-label="展开办公侧边栏"
                  title="展开侧边栏"
                  onClick={onOpenSidebar}
                >
                  <PanelLeftOpen size={17} />
                </button>
              )}
              <button
                type="button"
                className="work-editor-back"
                aria-label="返回办公首页"
                title="返回办公首页"
                onClick={onBack}
              >
                <ArrowLeft size={17} />
              </button>
              <span className={`work-file-kind-icon ${artifact.kind}`}>
                <FileKindIcon kind={artifact.kind} size={16} />
              </span>
              <div className="work-editor-identity">
                <input
                  className="work-office-text-field"
                  value={artifact.title}
                  aria-label="文件名"
                  title={`${artifact.title}.${extension.toLocaleLowerCase()}`}
                  onChange={(event) => onRename(event.target.value)}
                  onBlur={() => {
                    if (!artifact.title.trim()) {
                      onRename(`无标题${fileKindLabel(artifact.kind)}`);
                    }
                  }}
                />
              </div>
            </div>
            <div className="work-editor-header-actions">
              {collaborationDemo && (
                <output
                  className="playground-collaboration-mode-status"
                  data-testid="collaboration-comment-mode"
                  aria-live="polite"
                >
                  <MessageSquareText size={14} />
                  <span>实时评论 · 2 人在线</span>
                </output>
              )}
              {documentSuggestionFixtureEnabled && (
                <output
                  className="playground-collaboration-mode-status"
                  data-testid="collaboration-suggestion-mode"
                  aria-live="polite"
                >
                  <Pencil size={14} />
                  <span>建议协作 · 2 人在线</span>
                </output>
              )}
              {documentSuggestionFixture && (
                <>
                  <output
                    className="playground-collaboration-fixture-status"
                    data-testid="native-document-suggestion-status"
                    data-state={documentSuggestionFixture.nativeStage}
                    data-native-editor-deletion={String(
                      documentSuggestionFixture.nativeProjection
                        .editorHasDeletion,
                    )}
                    data-native-editor-insertion={String(
                      documentSuggestionFixture.nativeProjection
                        .editorHasInsertion,
                    )}
                    data-native-suggester-deletion={String(
                      documentSuggestionFixture.nativeProjection
                        .suggesterHasDeletion,
                    )}
                    data-native-suggester-insertion={String(
                      documentSuggestionFixture.nativeProjection
                        .suggesterHasInsertion,
                    )}
                    aria-live="polite"
                  >
                    {nativeDocumentSuggestionStatus(
                      documentSuggestionFixture.nativeStage,
                    )}
                  </output>
                  <button
                    type="button"
                    className="work-editor-ai-button"
                    data-testid="native-document-suggestion-action"
                    data-state={documentSuggestionFixture.nativeStage}
                    aria-label={nativeDocumentSuggestionAction(
                      documentSuggestionFixture.nativeStage,
                    )}
                    disabled={
                      documentSuggestionFixture.nativeStage === 'accepted'
                    }
                    onClick={documentSuggestionFixture.advanceNativeSuggestion}
                  >
                    <Sparkles size={15} />
                    <span>
                      {nativeDocumentSuggestionAction(
                        documentSuggestionFixture.nativeStage,
                      )}
                    </span>
                  </button>
                </>
              )}
              {controlledReviewFixture &&
                artifact.content.type === 'document' && (
                  <button
                    type="button"
                    className="work-editor-ai-button"
                    aria-label={
                      controlledReviewFixtureReady
                        ? '模拟外部审阅更新'
                        : '加载审阅冲突夹具'
                    }
                    onClick={() => {
                      const current = artifact.content;
                      if (current.type !== 'document') return;
                      onChange(nextControlledReviewFixtureContent(current));
                    }}
                  >
                    <TriangleAlert size={15} />
                    <span>
                      {controlledReviewFixtureReady
                        ? '模拟外部更新'
                        : '加载审阅夹具'}
                    </span>
                  </button>
                )}
              {pdfCollaborationFixtureEnabled && (
                <>
                  <output
                    className="playground-collaboration-fixture-status"
                    data-testid="pdf-collaboration-annotation-status"
                    data-state={
                      pdfCollaborationFixture?.annotationStage ?? 'loading'
                    }
                    aria-live="polite"
                  >
                    {pdfCollaborationFixture
                      ? pdfCollaborationAnnotationStatus(
                          pdfCollaborationFixture.annotationStage,
                        )
                      : '正在准备 PDF 协作'}
                  </output>
                  {pdfCollaborationFixture && (
                    <button
                      type="button"
                      className="work-editor-ai-button"
                      aria-label={pdfCollaborationAnnotationAction(
                        pdfCollaborationFixture.annotationStage,
                      )}
                      disabled={
                        pdfCollaborationFixture.annotationStage === 'deleted'
                      }
                      onClick={pdfCollaborationFixture.advanceAnnotation}
                    >
                      <MessageSquareText size={15} />
                      <span>
                        {pdfCollaborationAnnotationAction(
                          pdfCollaborationFixture.annotationStage,
                        )}
                      </span>
                    </button>
                  )}
                </>
              )}
              {spreadsheetCollaborationFixtureEnabled && (
                <>
                  <output
                    className="playground-collaboration-fixture-status"
                    data-testid="spreadsheet-collaboration-cell-status"
                    data-state={
                      spreadsheetCollaborationFixture?.cellStage ?? 'loading'
                    }
                    aria-live="polite"
                  >
                    {spreadsheetCollaborationFixture
                      ? spreadsheetCollaborationCellStatus(
                          spreadsheetCollaborationFixture.cellStage,
                        )
                      : '正在准备 Spreadsheet 协作'}
                  </output>
                  {spreadsheetCollaborationFixture && (
                    <button
                      type="button"
                      className="work-editor-ai-button"
                      aria-label={spreadsheetCollaborationCellAction(
                        spreadsheetCollaborationFixture.cellStage,
                      )}
                      disabled={
                        spreadsheetCollaborationFixture.cellStage === 'deleted'
                      }
                      onClick={spreadsheetCollaborationFixture.advanceCell}
                    >
                      <Pencil size={15} />
                      <span>
                        {spreadsheetCollaborationCellAction(
                          spreadsheetCollaborationFixture.cellStage,
                        )}
                      </span>
                    </button>
                  )}
                </>
              )}
              {presentationCollaborationFixtureEnabled && (
                <>
                  <output
                    className="playground-collaboration-fixture-status"
                    data-testid="presentation-collaboration-element-status"
                    data-state={
                      presentationCollaborationFixture?.elementStage ??
                      'loading'
                    }
                    aria-live="polite"
                  >
                    {presentationCollaborationFixture
                      ? presentationCollaborationElementStatus(
                          presentationCollaborationFixture.elementStage,
                        )
                      : '正在准备 Presentation 协作'}
                  </output>
                  {presentationCollaborationFixture && (
                    <button
                      type="button"
                      className="work-editor-ai-button"
                      aria-label={presentationCollaborationElementAction(
                        presentationCollaborationFixture.elementStage,
                      )}
                      disabled={
                        presentationCollaborationFixture.elementStage ===
                        'deleted'
                      }
                      onClick={presentationCollaborationFixture.advanceElement}
                    >
                      <Pencil size={15} />
                      <span>
                        {presentationCollaborationElementAction(
                          presentationCollaborationFixture.elementStage,
                        )}
                      </span>
                    </button>
                  )}
                </>
              )}
              {artifact.kind !== 'pdf' && !documentSuggestionFixtureEnabled && (
                <fieldset className="playground-mode-switch">
                  <legend className="sr-only">编辑或预览</legend>
                  <button
                    type="button"
                    className={!preview ? 'active' : ''}
                    aria-label="编辑"
                    aria-pressed={!preview}
                    title="编辑"
                    onClick={() => setPreview(false)}
                  >
                    <Pencil size={14} />
                    <span>编辑</span>
                  </button>
                  <button
                    type="button"
                    className={preview ? 'active' : ''}
                    aria-label="预览"
                    aria-pressed={preview}
                    title="预览"
                    onClick={() => setPreview(true)}
                  >
                    <Eye size={15} />
                    <span>预览</span>
                  </button>
                </fieldset>
              )}
              <button
                type="button"
                className={`work-editor-ai-button ${assistantOpen ? 'active' : ''}`}
                aria-label={assistantOpen ? '关闭 AI 助手' : '打开 AI 助手'}
                aria-pressed={assistantOpen}
                onClick={onToggleAssistant}
              >
                <Sparkles size={15} />
                <span>AI 助手</span>
              </button>
              <EditorExportButton
                key={artifact.id}
                kind={artifact.kind}
                exporting={exporting}
                onExport={() => void exportArtifact()}
                onExportPdf={
                  artifact.content.type === 'document'
                    ? () => void exportDocumentPdf()
                    : undefined
                }
              />
            </div>
          </header>
          {artifact.content.type === 'document' &&
            !documentSuggestionFixtureEnabled &&
            (!collaborationPresenceFixtureEnabled ||
              collaborationPresenceFixture) && (
              <DocumentEditor
                artifactId={artifact.id}
                collaboration={collaborationPresenceFixture?.collaboration}
                content={artifact.content}
                getSelectionMenuItems={getDocumentSelectionMenuItems}
                onAgentRequest={handleAgentRequest}
                onChange={(content: DocumentContent) => onChange(content)}
                onReviewConflict={(event) =>
                  onNotice(
                    `检测到 ${event.conflicts.length} 个审阅冲突`,
                    'neutral',
                  )
                }
                preview={preview}
                presence={collaborationPresenceFixture?.presence}
                saveStatus="本次会话已保存"
                theme="light"
              />
            )}
          {artifact.content.type === 'document' &&
            !documentSuggestionFixtureEnabled &&
            collaborationPresenceFixtureEnabled &&
            !collaborationPresenceFixture && (
              <div role="status">正在准备协作测试夹具</div>
            )}
          {artifact.content.type === 'document' &&
            documentSuggestionFixtureEnabled &&
            documentSuggestionFixture && (
              <section
                key={documentSuggestionFixture.nativeStage}
                className="playground-suggestion-demo"
                aria-label="双人建议协作"
              >
                <section
                  className="playground-suggestion-peer suggester"
                  data-testid="suggestion-peer"
                  aria-label="建议者 林澄"
                >
                  <header>
                    <div>
                      <span>建议者</span>
                      <strong>林澄</strong>
                    </div>
                    <small>只能提交或撤回自己的文字建议</small>
                  </header>
                  <DocumentEditor
                    artifactId={documentSuggestionFixture.artifactId}
                    collaboration={documentSuggestionFixture.suggester}
                    content={documentSuggestionFixture.content}
                    onChange={documentSuggestionFixture.updateContent}
                    preview={false}
                    saveStatus="建议已同步"
                    theme="light"
                  />
                </section>
                <section
                  className="playground-suggestion-peer editor"
                  data-testid="editor-peer"
                  aria-label="编辑者 周宁"
                >
                  <header>
                    <div>
                      <span>编辑者</span>
                      <strong>周宁</strong>
                    </div>
                    <small data-testid="suggestion-decision-status">
                      {pendingSuggestion
                        ? '有待处理建议'
                        : suggestionDecisionCount > 0
                          ? `已记录 ${suggestionDecisionCount} 项决定`
                          : '可接受或拒绝建议'}
                    </small>
                  </header>
                  <DocumentEditor
                    artifactId={documentSuggestionFixture.artifactId}
                    collaboration={documentSuggestionFixture.editor}
                    content={documentSuggestionFixture.content}
                    onChange={documentSuggestionFixture.updateContent}
                    onReviewConflict={(event) =>
                      onNotice(
                        `检测到 ${event.conflicts.length} 个建议决定冲突`,
                        'neutral',
                      )
                    }
                    preview={false}
                    saveStatus="决定已同步"
                    theme="light"
                  />
                </section>
              </section>
            )}
          {artifact.content.type === 'document' &&
            documentSuggestionFixtureEnabled &&
            !documentSuggestionFixture && (
              <div role="status">正在准备双人建议协作</div>
            )}
          {artifact.content.type === 'markdown' && (
            <MarkdownEditor
              content={artifact.content}
              getSelectionMenuItems={getMarkdownSelectionMenuItems}
              onChange={(content: MarkdownContent) => onChange(content)}
              preview={preview}
              saveStatus="本次会话已保存"
              theme="light"
            />
          )}
          {artifact.content.type === 'spreadsheet' &&
            (!spreadsheetCollaborationFixtureEnabled ||
              spreadsheetCollaborationFixture) && (
              <SpreadsheetEditor
                collaboration={spreadsheetCollaborationFixture?.collaboration}
                content={artifact.content}
                onAgentRequest={handleAgentRequest}
                onChange={(content: SpreadsheetContent) => onChange(content)}
                preview={preview}
                saveStatus="本次会话已保存"
                theme="light"
              />
            )}
          {artifact.content.type === 'spreadsheet' &&
            spreadsheetCollaborationFixtureEnabled &&
            !spreadsheetCollaborationFixture && (
              <div role="status">正在准备 Spreadsheet 协作测试夹具</div>
            )}
          {artifact.content.type === 'presentation' &&
            (!presentationCollaborationFixtureEnabled ||
              presentationCollaborationFixture) && (
              <PresentationEditor
                collaboration={presentationCollaborationFixture?.collaboration}
                content={artifact.content}
                onAgentRequest={handleAgentRequest}
                onChange={(content: PresentationContent) => onChange(content)}
                preview={preview}
                saveStatus="本次会话已保存"
                theme="light"
              />
            )}
          {artifact.content.type === 'presentation' &&
            presentationCollaborationFixtureEnabled &&
            !presentationCollaborationFixture && (
              <div role="status">正在准备 Presentation 协作测试夹具</div>
            )}
          {artifact.content.type === 'pdf' &&
            (!pdfCollaborationFixtureEnabled || pdfCollaborationFixture) && (
              <PdfViewer
                collaboration={pdfCollaborationFixture?.collaboration}
                fileName={
                  artifact.source?.name ??
                  `${artifact.title.toLocaleLowerCase()}.pdf`
                }
                loadSource={loadPdf}
                onSave={savePdf}
                sourceKey={`${artifact.id}:${artifact.revision}`}
                theme="light"
              />
            )}
          {artifact.content.type === 'pdf' &&
            pdfCollaborationFixtureEnabled &&
            !pdfCollaborationFixture && (
              <div role="status">正在准备 PDF 协作测试夹具</div>
            )}
        </section>

        {assistantOpen && (
          <AssistantPanel
            artifact={artifact}
            lastRequest={lastAgentRequest}
            modal={assistantModal}
            questionDraft={assistantQuestion}
            width={assistantWidth}
            onCancelQuestion={() => setAssistantQuestion(null)}
            onClose={onToggleAssistant}
            onQuestionChange={(question) =>
              setAssistantQuestion((current) =>
                current ? { ...current, question } : current,
              )
            }
            onSubmitQuestion={() => {
              if (!assistantQuestion) return;
              const request =
                createPlaygroundAssistantQuestionRequest(assistantQuestion);
              if (request) handleAgentRequest(request);
            }}
            onWidthChange={onAssistantWidthChange}
          />
        )}
      </div>
    </section>
  );
}

function nativeDocumentSuggestionStatus(
  stage: PlaygroundNativeDocumentSuggestionStage,
): string {
  switch (stage) {
    case 'ready':
      return '原生建议夹具已就绪';
    case 'proposed':
      return 'A3S Agent 已提议把 😀 替换为 reviewed';
    case 'accepted':
      return 'Native Editor 已接受 2 项建议';
  }
}

function nativeDocumentSuggestionAction(
  stage: PlaygroundNativeDocumentSuggestionStage,
): string {
  switch (stage) {
    case 'ready':
      return '运行原生 Agent 提议';
    case 'proposed':
      return '运行原生编辑者接受';
    case 'accepted':
      return '原生建议流程已完成';
  }
}

function pdfCollaborationAnnotationStatus(
  stage: PlaygroundPdfAnnotationStage,
): string {
  switch (stage) {
    case 'ready':
      return 'PDF 协作已连接';
    case 'created':
      return '代理已创建共享批注';
    case 'updated':
      return '代理已更新共享批注';
    case 'deleted':
      return '代理已删除共享批注';
  }
}

function pdfCollaborationAnnotationAction(
  stage: PlaygroundPdfAnnotationStage,
): string {
  switch (stage) {
    case 'ready':
      return '模拟原生创建 PDF 批注';
    case 'created':
      return '模拟原生更新 PDF 批注';
    case 'updated':
      return '模拟原生删除 PDF 批注';
    case 'deleted':
      return '原生 PDF 批注已删除';
  }
}

function spreadsheetCollaborationCellStatus(
  stage: PlaygroundSpreadsheetCellStage,
): string {
  switch (stage) {
    case 'ready':
      return 'Spreadsheet 协作已连接';
    case 'batched':
      return '代理已原子更新、创建并删除 Data 单元格';
    case 'updated':
      return '代理已把 Data!A2 更新为 12';
    case 'created':
      return '代理已创建 Empty!F101 稀疏单元格';
    case 'deleted':
      return '代理已删除 Sparse!D6 稀疏单元格';
  }
}

function spreadsheetCollaborationCellAction(
  stage: PlaygroundSpreadsheetCellStage,
): string {
  switch (stage) {
    case 'ready':
      return '模拟原子批量修改 Spreadsheet';
    case 'batched':
      return '模拟原生更新 Spreadsheet 单元格';
    case 'updated':
      return '模拟原生创建稀疏单元格';
    case 'created':
      return '模拟原生删除稀疏单元格';
    case 'deleted':
      return '原生 Spreadsheet 单元格生命周期已完成';
  }
}

function presentationCollaborationElementStatus(
  stage: PlaygroundPresentationElementStage,
): string {
  switch (stage) {
    case 'ready':
      return 'Presentation 协作已连接';
    case 'updated':
      return '代理已更新共享标题对象';
    case 'created':
      return '代理已创建共享场景对象';
    case 'reordered':
      return '代理已调整共享对象层级';
    case 'deleted':
      return '代理已删除第二页场景对象';
  }
}

function presentationCollaborationElementAction(
  stage: PlaygroundPresentationElementStage,
): string {
  switch (stage) {
    case 'ready':
      return '模拟原生更新 Presentation 对象';
    case 'updated':
      return '模拟原生创建 Presentation 对象';
    case 'created':
      return '模拟原生调整 Presentation 对象层级';
    case 'reordered':
      return '模拟原生删除 Presentation 对象';
    case 'deleted':
      return '原生 Presentation 对象生命周期已完成';
  }
}

const CONTROLLED_REVIEW_COMMENT_ID = 'e2e-controlled-review-comment';
const CONTROLLED_REVIEW_CHANGE_ID = 'e2e-controlled-review-change';
const CONTROLLED_REVIEW_TEXT = '描述当前情况';

function nextControlledReviewFixtureContent(
  content: DocumentContent,
): DocumentContent {
  const comments = (content.comments ?? []).filter(
    (comment) => comment.id !== CONTROLLED_REVIEW_COMMENT_ID,
  );
  const fixtureComment = {
    id: CONTROLLED_REVIEW_COMMENT_ID,
    author: 'External reviewer',
    date: '2026-08-11T00:00:00.000Z',
    text: 'Verify the controlled review boundary.',
    resolved: false,
  };
  if (!content.html.includes(CONTROLLED_REVIEW_COMMENT_ID)) {
    const reviewedText = [
      `<span data-comment-id="${CONTROLLED_REVIEW_COMMENT_ID}" data-document-comment="true">${CONTROLLED_REVIEW_TEXT}</span>`,
      '、核心问题和可衡量的成功标准。',
      `<ins data-change-kind="insertion" data-change-id="${CONTROLLED_REVIEW_CHANGE_ID}" data-change-author="External reviewer" data-change-date="2026-08-11T00:00:00.000Z" data-document-change="true">待确认</ins>`,
    ].join('');
    const sourceText = `${CONTROLLED_REVIEW_TEXT}、核心问题和可衡量的成功标准。`;
    const html = content.html.includes(sourceText)
      ? content.html.replace(sourceText, reviewedText)
      : `${content.html}<p>${reviewedText}</p>`;
    return {
      ...content,
      html,
      model: undefined,
      comments: [...comments, fixtureComment],
    };
  }
  const commentMarkup = `<span data-comment-id="${CONTROLLED_REVIEW_COMMENT_ID}" data-document-comment="true">${CONTROLLED_REVIEW_TEXT}</span>`;
  const changeMarkup = `<ins data-change-kind="insertion" data-change-id="${CONTROLLED_REVIEW_CHANGE_ID}" data-change-author="External reviewer" data-change-date="2026-08-11T00:00:00.000Z" data-document-change="true">待确认</ins>`;
  return {
    ...content,
    html: content.html
      .replace(
        commentMarkup,
        `<span data-comment-id="${CONTROLLED_REVIEW_COMMENT_ID}" data-document-comment="true">外部改写内容</span>`,
      )
      .replace(changeMarkup, '待确认'),
    model: undefined,
    comments: [...comments, fixtureComment],
  };
}

export function EditorExportButton({
  kind,
  exporting,
  onExport,
  onExportPdf,
}: {
  kind: OfficeArtifact['kind'];
  exporting: boolean;
  onExport: () => void;
  onExportPdf?: () => void;
}) {
  const [menuOpen, setMenuOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);
  const triggerRef = useRef<HTMLButtonElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const label = kind === 'pdf' ? '下载 PDF' : '导出';
  const hasFormatMenu = Boolean(onExportPdf);

  useEffect(() => {
    if (!menuOpen || !hasFormatMenu) return;
    menuRef.current?.querySelector<HTMLButtonElement>('button')?.focus();

    const closeFromOutside = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setMenuOpen(false);
    };
    const closeFromKeyboard = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      setMenuOpen(false);
      triggerRef.current?.focus();
    };
    document.addEventListener('pointerdown', closeFromOutside);
    document.addEventListener('keydown', closeFromKeyboard);
    return () => {
      document.removeEventListener('pointerdown', closeFromOutside);
      document.removeEventListener('keydown', closeFromKeyboard);
    };
  }, [hasFormatMenu, menuOpen]);

  const runExport = (action: () => void) => {
    setMenuOpen(false);
    action();
  };

  const moveMenuFocus = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) return;
    const items = Array.from(
      menuRef.current?.querySelectorAll<HTMLButtonElement>(
        '[role="menuitem"]',
      ) ?? [],
    );
    if (!items.length) return;
    event.preventDefault();
    const current = items.indexOf(document.activeElement as HTMLButtonElement);
    const index =
      event.key === 'Home'
        ? 0
        : event.key === 'End'
          ? items.length - 1
          : (current + (event.key === 'ArrowDown' ? 1 : -1) + items.length) %
            items.length;
    items[index]?.focus();
  };

  return (
    <div className="work-export-menu" ref={rootRef}>
      <button
        ref={triggerRef}
        type="button"
        className="work-export-button"
        aria-label={label}
        aria-busy={exporting || undefined}
        aria-expanded={hasFormatMenu ? menuOpen : undefined}
        aria-haspopup={hasFormatMenu ? 'menu' : undefined}
        title={label}
        disabled={exporting}
        onClick={() =>
          hasFormatMenu ? setMenuOpen((current) => !current) : onExport()
        }
      >
        <Download size={15} />
        <span>{label}</span>
      </button>
      {onExportPdf && menuOpen && !exporting && (
        <div
          ref={menuRef}
          className="work-export-menu-panel"
          role="menu"
          aria-label="导出格式"
          onKeyDown={moveMenuFocus}
        >
          <button
            type="button"
            role="menuitem"
            onClick={() => runExport(onExport)}
          >
            <Download size={14} />
            下载 DOCX
          </button>
          <button
            type="button"
            role="menuitem"
            onClick={() => runExport(onExportPdf)}
          >
            <Download size={14} />
            导出 PDF
          </button>
        </div>
      )}
    </div>
  );
}

interface PlaygroundSelectionContext {
  selection: {
    text: string;
    beforeText: string;
    afterText: string;
  };
  document: {
    text: string;
  };
  commands: {
    copyText(): Promise<boolean>;
  };
}

interface PlaygroundSelectionMenuItem<Context> {
  id: string;
  label: string;
  icon: 'copy' | 'message' | 'sparkles' | 'wand';
  separatorBefore?: boolean;
  onSelect(context: Context): void | Promise<void>;
}

export interface PlaygroundAssistantQuestionDraft {
  question: string;
  selectionContext: string;
  selectionPreview: string;
}

export function createPlaygroundSelectionMenuItems<
  Context extends PlaygroundSelectionContext,
>(
  onAgentRequest: (request: EditorAgentRequest) => void,
  onQuestionDraft: (draft: PlaygroundAssistantQuestionDraft) => void,
  onNotice: (message: string, tone?: NoticeTone) => void,
): readonly PlaygroundSelectionMenuItem<Context>[] {
  return [
    {
      id: 'copy',
      label: '复制',
      icon: 'copy',
      onSelect: async ({ commands }) => {
        const copied = await commands.copyText();
        onNotice(
          copied ? '选中文本已复制' : '无法访问剪贴板，请使用系统快捷键',
          copied ? 'success' : 'danger',
        );
      },
    },
    {
      id: 'expand',
      label: '扩写选中内容',
      icon: 'sparkles',
      separatorBefore: true,
      onSelect: (context) =>
        onAgentRequest({
          instruction:
            '请结合全文和相邻段落扩写选中内容，保持原有事实、语气和文档结构；先返回可审阅草稿。',
          selection: selectionPromptContext(context),
        }),
    },
    {
      id: 'polish',
      label: '润色表达',
      icon: 'wand',
      onSelect: (context) =>
        onAgentRequest({
          instruction:
            '请结合全文语气润色选中内容，减少重复和空话，保持事实与专业术语不变；先返回可审阅草稿。',
          selection: selectionPromptContext(context),
        }),
    },
    {
      id: 'ask',
      label: '询问 AI 助手',
      icon: 'message',
      onSelect: (context) => onQuestionDraft(createQuestionDraft(context)),
    },
  ];
}

function createQuestionDraft(
  context: PlaygroundSelectionContext,
): PlaygroundAssistantQuestionDraft {
  const selection = context.selection.text.trim();
  const previewLimit = 240;
  return {
    question: '',
    selectionContext: selectionPromptContext(context),
    selectionPreview:
      selection.length > previewLimit
        ? `${selection.slice(0, previewLimit - 1).trimEnd()}…`
        : selection,
  };
}

export function createPlaygroundAssistantQuestionRequest(
  draft: PlaygroundAssistantQuestionDraft,
): EditorAgentRequest | null {
  const question = draft.question.trim();
  if (!question) return null;
  return {
    instruction: `请结合已附带的文档上下文回答：\n\n${question}`,
    selection: draft.selectionContext,
  };
}

function selectionPromptContext(context: PlaygroundSelectionContext): string {
  return [
    `选中文本：\n${context.selection.text}`,
    `前文：\n${context.selection.beforeText || '（无）'}`,
    `后文：\n${context.selection.afterText || '（无）'}`,
    `完整文档：\n${context.document.text}`,
  ].join('\n\n');
}

export function AssistantPanel({
  artifact,
  lastRequest,
  modal,
  questionDraft,
  width,
  onCancelQuestion,
  onClose,
  onQuestionChange,
  onSubmitQuestion,
  onWidthChange,
}: {
  artifact: OfficeArtifact;
  lastRequest: EditorAgentRequest | null;
  modal: boolean;
  questionDraft: PlaygroundAssistantQuestionDraft | null;
  width: number;
  onCancelQuestion: () => void;
  onClose: () => void;
  onQuestionChange: (question: string) => void;
  onSubmitQuestion: () => void;
  onWidthChange: (width: number) => void;
}) {
  const closeButtonRef = useRef<HTMLButtonElement>(null);
  const questionInputRef = useRef<HTMLTextAreaElement>(null);
  const questionSelectionContext = questionDraft?.selectionContext;
  const focusScope = useDialogFocusScope<HTMLElement>({
    active: modal,
    onEscape: onClose,
    initialFocus: () =>
      questionDraft ? questionInputRef.current : closeButtonRef.current,
  });
  useEffect(() => {
    if (questionSelectionContext) questionInputRef.current?.focus();
  }, [questionSelectionContext]);
  const modalAttributes = modal
    ? ({ role: 'dialog', 'aria-modal': true } as const)
    : {};
  const welcomeCopy = assistantWelcomeCopy(artifact.kind);
  const resize = (event: ReactPointerEvent<HTMLHRElement>) => {
    if (event.button !== 0) return;
    event.preventDefault();
    const startX = event.clientX;
    const startWidth = width;
    document.body.classList.add('playground-resizing');

    const move = (moveEvent: PointerEvent) => {
      const nextWidth = startWidth + startX - moveEvent.clientX;
      onWidthChange(clampAssistantWidth(nextWidth));
    };
    const stop = () => {
      document.body.classList.remove('playground-resizing');
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', stop);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', stop, { once: true });
  };

  return (
    <aside
      {...modalAttributes}
      ref={focusScope.scopeRef}
      className="playground-assistant"
      aria-label="AI 助手"
      style={{ width }}
      onKeyDown={focusScope.handleKeyDown}
    >
      <hr
        className="playground-assistant-resizer"
        aria-label="调整 AI 助手宽度"
        aria-orientation="vertical"
        aria-valuemin={assistantMinimumWidth}
        aria-valuemax={assistantMaximumWidth}
        aria-valuenow={width}
        tabIndex={0}
        onKeyDown={(event) => {
          if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
          event.preventDefault();
          const delta = event.key === 'ArrowLeft' ? 20 : -20;
          onWidthChange(clampAssistantWidth(width + delta));
        }}
        onPointerDown={resize}
      />
      <header>
        <span className="playground-assistant-mark">
          <Sparkles size={15} />
        </span>
        <div>
          <strong>AI 助手</strong>
          <small>{artifact.title}</small>
        </div>
        <button
          ref={closeButtonRef}
          type="button"
          className="playground-icon-button"
          aria-label="关闭 AI 助手"
          onClick={onClose}
        >
          <X size={16} />
        </button>
      </header>
      <div className="playground-assistant-content">
        {questionDraft ? (
          <section className="playground-agent-question">
            <span className="playground-agent-request-icon">
              <MessageSquareText size={18} />
            </span>
            <div>
              <small>询问选中内容</small>
              <h2>你想问什么？</h2>
            </div>
            <blockquote>{questionDraft.selectionPreview}</blockquote>
            <p>已附带选中文本和文档上下文。</p>
            <form
              onSubmit={(event) => {
                event.preventDefault();
                onSubmitQuestion();
              }}
            >
              <label>
                <span>问题</span>
                <textarea
                  ref={questionInputRef}
                  aria-label="向 AI 助手提问"
                  value={questionDraft.question}
                  placeholder="例如：这段结论有哪些依据？"
                  rows={4}
                  onChange={(event) => onQuestionChange(event.target.value)}
                />
              </label>
              <div className="playground-agent-question-actions">
                <button
                  type="button"
                  className="playground-secondary-button"
                  onClick={onCancelQuestion}
                >
                  取消
                </button>
                <button
                  type="submit"
                  className="playground-primary-button"
                  disabled={!questionDraft.question.trim()}
                >
                  <SendHorizontal size={14} />
                  发送问题
                </button>
              </div>
            </form>
          </section>
        ) : lastRequest ? (
          <section className="playground-agent-request">
            <span className="playground-agent-request-icon">
              <Sparkles size={18} />
            </span>
            <div>
              <small>AI 请求</small>
              <h2>{lastRequest.instruction}</h2>
            </div>
            {lastRequest.selection && (
              <details>
                <summary>查看附带上下文</summary>
                <blockquote>{lastRequest.selection}</blockquote>
              </details>
            )}
            <p>请求已准备好，在线 Playground 不会自动上传文件。</p>
          </section>
        ) : (
          <section className="playground-assistant-welcome">
            <span>
              <Sparkles size={21} />
            </span>
            <h2>{welcomeCopy.title}</h2>
            <p>{welcomeCopy.description}</p>
          </section>
        )}
      </div>
    </aside>
  );
}

function assistantWelcomeCopy(kind: OfficeArtifact['kind']): {
  title: string;
  description: string;
} {
  switch (kind) {
    case 'document':
      return {
        title: '从选中文本开始',
        description: '选中文本后，可从右键菜单发起扩写、润色或提问。',
      };
    case 'markdown':
      return {
        title: '从选中内容开始',
        description: '选中 Markdown 内容后，可从右键菜单发起改写或提问。',
      };
    case 'spreadsheet':
      return {
        title: '从单元格开始',
        description: '选中单元格后，可从菜单发起分析或整理。',
      };
    case 'presentation':
      return {
        title: '从幻灯片内容开始',
        description: '选中文本或对象后，可从菜单发起改写或提问。',
      };
    case 'pdf':
      return {
        title: '当前 PDF 暂无 AI 请求',
        description: 'PDF 阅读与批注可直接使用；摘要与问答接入仍在完善。',
      };
  }
}

function clampAssistantWidth(width: number): number {
  return Math.round(
    Math.max(assistantMinimumWidth, Math.min(assistantMaximumWidth, width)),
  );
}
