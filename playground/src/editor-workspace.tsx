import {
  ArrowLeft,
  Check,
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
  type EditorAgentRequest,
  type MarkdownContent,
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

  return (
    <section className={`work-editor-shell ${artifact.kind}`}>
      <header className="work-editor-header playground-editor-header">
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
          <FileKindIcon kind={artifact.kind} size={17} />
        </span>
        <div className="work-editor-identity">
          <input
            className="work-office-text-field"
            value={artifact.title}
            aria-label="文件名"
            onChange={(event) => onRename(event.target.value)}
            onBlur={() => {
              if (!artifact.title.trim()) {
                onRename(`无标题${fileKindLabel(artifact.kind)}`);
              }
            }}
          />
          <span>
            {extension}
            <i aria-hidden="true">·</i>
            <span>
              <Check size={11} />
              本次会话已保存
            </span>
          </span>
        </div>
        <div className="work-editor-header-actions">
          {artifact.kind !== 'pdf' && (
            <fieldset className="playground-mode-switch">
              <legend className="sr-only">编辑或预览</legend>
              <button
                type="button"
                className={!preview ? 'active' : ''}
                aria-pressed={!preview}
                onClick={() => setPreview(false)}
              >
                <Pencil size={14} />
                <span>编辑</span>
              </button>
              <button
                type="button"
                className={preview ? 'active' : ''}
                aria-pressed={preview}
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
          <button
            type="button"
            className="work-export-button"
            disabled={exporting}
            onClick={() => void exportArtifact()}
          >
            <Download size={15} />
            <span>{artifact.kind === 'pdf' ? '下载 PDF' : '导出'}</span>
          </button>
        </div>
      </header>

      <div className="playground-editor-row">
        <section
          className="playground-editor-host"
          aria-label={`${fileKindLabel(artifact.kind)}编辑器`}
        >
          {artifact.content.type === 'document' && (
            <DocumentEditor
              content={artifact.content}
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
              onStartSlideshow={() => setPreview(true)}
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
