import {
  ArrowLeft,
  Download,
  Eye,
  PanelLeftOpen,
  Pencil,
  Sparkles,
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
import {
  downloadArtifact,
  downloadArtifactPdf,
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
import { useDialogFocusScope } from '../../src/internal/design-system/primitives/overlay/dialog-focus-scope';
import { FileKindIcon, fileKindExtension, fileKindLabel } from './file-kind';
import type { NoticeTone } from './playground-types';

const assistantMinimumWidth = 340;
const assistantMaximumWidth = 680;

export function EditorWorkspace({
  artifact,
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
          {artifact.content.type === 'document' && (
            <DocumentEditor
              artifactId={artifact.id}
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
            modal={assistantModal}
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

export function AssistantPanel({
  artifact,
  lastRequest,
  modal,
  width,
  onClose,
  onWidthChange,
}: {
  artifact: OfficeArtifact;
  lastRequest: EditorAgentRequest | null;
  modal: boolean;
  width: number;
  onClose: () => void;
  onWidthChange: (width: number) => void;
}) {
  const closeButtonRef = useRef<HTMLButtonElement>(null);
  const focusScope = useDialogFocusScope<HTMLElement>({
    active: modal,
    onEscape: onClose,
    initialFocus: () => closeButtonRef.current,
  });
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
        {lastRequest ? (
          <section className="playground-agent-request">
            <span className="playground-agent-request-icon">
              <Sparkles size={18} />
            </span>
            <div>
              <small>AI 请求</small>
              <h2>{lastRequest.instruction}</h2>
            </div>
            {lastRequest.selection && (
              <blockquote>{lastRequest.selection}</blockquote>
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
