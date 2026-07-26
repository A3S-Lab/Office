import {
  ArrowLeft,
  Code2,
  Download,
  Eye,
  PanelLeftOpen,
  Pencil,
  Sparkles,
  X,
} from 'lucide-react';
import {
  type PointerEvent as ReactPointerEvent,
  useCallback,
  useState,
} from 'react';
import {
  downloadArtifact,
  readSourceBlob,
  registerSourceBlob,
  type DocumentContent,
  type DocumentSelectionContext,
  type EditorAgentRequest,
  type GetDocumentSelectionMenuItems,
  type GetMarkdownSelectionMenuItems,
  type MarkdownContent,
  type MarkdownSelectionContext,
  type OfficeArtifact,
  type OfficeArtifactContent,
  type PresentationContent,
  type SpreadsheetContent,
} from '@a3s-lab/office/core';
import {
  DocumentEditor,
  MarkdownEditor,
  PdfViewer,
  PresentationEditor,
  SpreadsheetEditor,
} from '@a3s-lab/office/react';
import { FileKindIcon, fileKindExtension, fileKindLabel } from './file-kind';
import type { NoticeTone } from './playground-types';

const assistantMinimumWidth = 340;
const assistantMaximumWidth = 680;

export function EditorWorkspace({
  artifact,
  sidebarOpen,
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
  sidebarOpen: boolean;
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
  const extension = fileKindExtension(artifact.kind);

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

  const savePdf = useCallback(
    async (pdf: Blob) => {
      registerSourceBlob(artifact.id, pdf);
      onTouch();
      onNotice('PDF 批注已保存到当前浏览器会话', 'success');
      return true;
    },
    [artifact.id, onNotice, onTouch],
  );
  const loadPdf = useCallback(() => readSourceBlob(artifact), [artifact]);
  const getDocumentSelectionMenuItems =
    useCallback<GetDocumentSelectionMenuItems>(
      () =>
        playgroundSelectionMenuItems<DocumentSelectionContext>(
          onAgentRequest,
          onNotice,
        ),
      [onAgentRequest, onNotice],
    );
  const getMarkdownSelectionMenuItems =
    useCallback<GetMarkdownSelectionMenuItems>(
      () =>
        playgroundSelectionMenuItems<MarkdownSelectionContext>(
          onAgentRequest,
          onNotice,
        ),
      [onAgentRequest, onNotice],
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
              {artifact.kind !== 'pdf' && (
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
                kind={artifact.kind}
                exporting={exporting}
                onExport={() => void exportArtifact()}
              />
            </div>
          </header>
          {artifact.content.type === 'document' && (
            <DocumentEditor
              content={artifact.content}
              getSelectionMenuItems={getDocumentSelectionMenuItems}
              onAgentRequest={onAgentRequest}
              onChange={(content: DocumentContent) => onChange(content)}
              preview={preview}
              saveStatus="本次会话已保存"
              theme="light"
            />
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
          {artifact.content.type === 'spreadsheet' && (
            <SpreadsheetEditor
              content={artifact.content}
              onAgentRequest={onAgentRequest}
              onChange={(content: SpreadsheetContent) => onChange(content)}
              preview={preview}
              saveStatus="本次会话已保存"
              theme="light"
            />
          )}
          {artifact.content.type === 'presentation' && (
            <PresentationEditor
              content={artifact.content}
              onAgentRequest={onAgentRequest}
              onChange={(content: PresentationContent) => onChange(content)}
              preview={preview}
              saveStatus="本次会话已保存"
              theme="light"
            />
          )}
          {artifact.content.type === 'pdf' && (
            <PdfViewer
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
        </section>

        {assistantOpen && (
          <AssistantPanel
            artifact={artifact}
            lastRequest={lastAgentRequest}
            width={assistantWidth}
            onClose={onToggleAssistant}
            onWidthChange={onAssistantWidthChange}
          />
        )}
      </div>
    </section>
  );
}

export function EditorExportButton({
  kind,
  exporting,
  onExport,
}: {
  kind: OfficeArtifact['kind'];
  exporting: boolean;
  onExport: () => void;
}) {
  const label = kind === 'pdf' ? '下载 PDF' : '导出';
  return (
    <button
      type="button"
      className="work-export-button"
      aria-label={label}
      aria-busy={exporting || undefined}
      title={label}
      disabled={exporting}
      onClick={onExport}
    >
      <Download size={15} />
      <span>{label}</span>
    </button>
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

function playgroundSelectionMenuItems<
  Context extends PlaygroundSelectionContext,
>(
  onAgentRequest: (request: EditorAgentRequest) => void,
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
      onSelect: (context) =>
        onAgentRequest({
          instruction: '请结合文档上下文回答关于选中内容的问题：\n\n问题：',
          selection: selectionPromptContext(context),
        }),
    },
  ];
}

function selectionPromptContext(context: PlaygroundSelectionContext): string {
  return [
    `选中文本：\n${context.selection.text}`,
    `前文：\n${context.selection.beforeText || '（无）'}`,
    `后文：\n${context.selection.afterText || '（无）'}`,
    `完整文档：\n${context.document.text}`,
  ].join('\n\n');
}

function AssistantPanel({
  artifact,
  lastRequest,
  width,
  onClose,
  onWidthChange,
}: {
  artifact: OfficeArtifact;
  lastRequest: EditorAgentRequest | null;
  width: number;
  onClose: () => void;
  onWidthChange: (width: number) => void;
}) {
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
      className="playground-assistant"
      aria-label="AI 助手"
      style={{ width }}
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
          type="button"
          className="playground-icon-button"
          aria-label="关闭 AI 助手"
          onClick={onClose}
        >
          <X size={16} />
        </button>
      </header>
      <div className="playground-assistant-content">
        {lastRequest ? (
          <section className="playground-agent-request">
            <span className="playground-agent-request-icon">
              <Sparkles size={18} />
            </span>
            <div>
              <small>编辑器请求</small>
              <h2>{lastRequest.instruction}</h2>
            </div>
            {lastRequest.selection && (
              <blockquote>{lastRequest.selection}</blockquote>
            )}
            <p>
              请求已经由 <code>onAgentRequest</code> 交给宿主。在线 Playground
              不会把文件发送到任何模型。
            </p>
          </section>
        ) : (
          <section className="playground-assistant-welcome">
            <span>
              <Sparkles size={21} />
            </span>
            <h2>和当前文件一起工作</h2>
            <p>
              在文字、表格或演示中选择内容，再从编辑器菜单发起 AI
              操作，即可在这里查看真实的宿主请求。
            </p>
          </section>
        )}
        <section className="playground-agent-hook">
          <span>
            <Code2 size={15} />
            接入自己的模型
          </span>
          <code>{'<DocumentEditor onAgentRequest={handleRequest} />'}</code>
        </section>
      </div>
    </aside>
  );
}

function clampAssistantWidth(width: number): number {
  return Math.round(
    Math.max(assistantMinimumWidth, Math.min(assistantMaximumWidth, width)),
  );
}
