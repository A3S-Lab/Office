import type {
  EditorAgentRequest,
  OfficeArtifact,
  OfficeArtifactKind,
  OfficeArtifactContent,
  OfficeFileImportProgress,
} from '@a3s-lab/office/core';
import { AlertCircle, CheckCircle2, Info } from 'lucide-react';
import {
  lazy,
  StrictMode,
  Suspense,
  useCallback,
  useEffect,
  useRef,
  useState,
} from 'react';
import { createRoot } from 'react-dom/client';
import '@a3s-lab/office/styles.css';
import { serializeDocumentParagraphFormatting } from '../../src/internal/features/work/work-document-paragraph-format-changes';
import { WorkEditorLoadingState } from '../../src/internal/features/work/components/work-editor-loading-state';
import { WORK_IMPORT_ACCEPT as OFFICE_FILE_ACCEPT } from '../../src/internal/features/work/work-file-contract';
import { workKindForFile } from '../../src/internal/features/work/work-file-kind';
import { createWorkArtifact as createArtifact } from '../../src/internal/features/work/work-templates';
import {
  loadPlaygroundEditorWorkspace,
  preloadPlaygroundEditor,
} from './editor-preload';
import {
  createMaximumSparseSpreadsheetArtifact,
  MAXIMUM_SPARSE_SPREADSHEET_ARTIFACT_ID,
  MAXIMUM_SPARSE_SPREADSHEET_FIXTURE,
} from './maximum-sparse-spreadsheet-fixture';
import {
  PlaygroundImportProgress,
  type PlaygroundImportState,
} from './playground-import-progress';
import type { NoticeTone, PlaygroundNotice } from './playground-types';
import {
  createSpreadsheetCopyFromAboveArtifact,
  SPREADSHEET_COPY_FROM_ABOVE_ARTIFACT_ID,
  SPREADSHEET_COPY_FROM_ABOVE_FIXTURE,
} from './spreadsheet-copy-from-above-fixture';
import {
  createSpreadsheetDateTimeArtifact,
  SPREADSHEET_DATE_TIME_ARTIFACT_ID,
  SPREADSHEET_DATE_TIME_FIXTURE,
} from './spreadsheet-date-time-fixture';
import {
  createSpreadsheetDataValidationArtifact,
  SPREADSHEET_DATA_VALIDATION_ARTIFACT_ID,
  SPREADSHEET_DATA_VALIDATION_FIXTURE,
} from './spreadsheet-data-validation-fixture';
import {
  createSpreadsheetGoToArtifact,
  SPREADSHEET_GO_TO_ARTIFACT_ID,
  SPREADSHEET_GO_TO_FIXTURE,
} from './spreadsheet-go-to-fixture';
import {
  createSpreadsheetGradientFillArtifact,
  SPREADSHEET_GRADIENT_FILL_ARTIFACT_ID,
  SPREADSHEET_GRADIENT_FILL_FIXTURE,
} from './spreadsheet-gradient-fill-fixture';
import {
  createSpreadsheetHyperlinkArtifact,
  SPREADSHEET_HYPERLINK_ARTIFACT_ID,
  SPREADSHEET_HYPERLINK_FIXTURE,
} from './spreadsheet-hyperlink-fixture';
import {
  createSpreadsheetPasteSpecialArtifact,
  SPREADSHEET_PASTE_SPECIAL_ARTIFACT_ID,
  SPREADSHEET_PASTE_SPECIAL_FIXTURE,
} from './spreadsheet-paste-special-fixture';
import {
  createSpreadsheetPatternFillArtifact,
  SPREADSHEET_PATTERN_FILL_ARTIFACT_ID,
  SPREADSHEET_PATTERN_FILL_FIXTURE,
} from './spreadsheet-pattern-fill-fixture';
import {
  createSpreadsheetRichTextArtifact,
  SPREADSHEET_RICH_TEXT_ARTIFACT_ID,
  SPREADSHEET_RICH_TEXT_FIXTURE,
} from './spreadsheet-rich-text-fixture';
import {
  collaborationServerDocumentationUrl,
  documentationEntryUrl,
  legacyDocsPath,
} from './site-routes';
import { SiteSidebar } from './site-sidebar';
import { useMediaQuery } from './use-media-query';
import { usePlaygroundSidebarState } from './use-playground-sidebar-state';
import { WorkspaceHome } from './workspace-home';
import './playground.css';
import './workspace.css';

const EditorWorkspace = lazy(async () => ({
  default: (await loadPlaygroundEditorWorkspace()).EditorWorkspace,
}));

function Playground() {
  const [e2eFixture] = useState(readPlaygroundE2eFixture);
  const sidebarModal = useMediaQuery('(max-width: 839px)');
  const {
    closeAll: closeSidebarForEditor,
    open: sidebarOpen,
    setOpen: setSidebarOpen,
  } = usePlaygroundSidebarState(sidebarModal);
  const [artifacts, setArtifacts] = useState<OfficeArtifact[]>(() =>
    createInitialArtifacts(e2eFixture),
  );
  const [activeArtifactId, setActiveArtifactId] = useState<string | null>(
    e2eFixture === MAXIMUM_SPARSE_SPREADSHEET_FIXTURE
      ? MAXIMUM_SPARSE_SPREADSHEET_ARTIFACT_ID
      : e2eFixture === SPREADSHEET_GO_TO_FIXTURE
        ? SPREADSHEET_GO_TO_ARTIFACT_ID
        : e2eFixture === SPREADSHEET_DATA_VALIDATION_FIXTURE
          ? SPREADSHEET_DATA_VALIDATION_ARTIFACT_ID
          : e2eFixture === SPREADSHEET_HYPERLINK_FIXTURE
            ? SPREADSHEET_HYPERLINK_ARTIFACT_ID
            : e2eFixture === SPREADSHEET_PASTE_SPECIAL_FIXTURE
              ? SPREADSHEET_PASTE_SPECIAL_ARTIFACT_ID
              : e2eFixture === SPREADSHEET_COPY_FROM_ABOVE_FIXTURE
                ? SPREADSHEET_COPY_FROM_ABOVE_ARTIFACT_ID
                : e2eFixture === SPREADSHEET_DATE_TIME_FIXTURE
                  ? SPREADSHEET_DATE_TIME_ARTIFACT_ID
                  : e2eFixture === SPREADSHEET_GRADIENT_FILL_FIXTURE
                    ? SPREADSHEET_GRADIENT_FILL_ARTIFACT_ID
                    : e2eFixture === SPREADSHEET_PATTERN_FILL_FIXTURE
                      ? SPREADSHEET_PATTERN_FILL_ARTIFACT_ID
                      : e2eFixture === SPREADSHEET_RICH_TEXT_FIXTURE
                        ? SPREADSHEET_RICH_TEXT_ARTIFACT_ID
                        : null,
  );
  const [collaborationDemoArtifactId, setCollaborationDemoArtifactId] =
    useState<string | null>(null);
  const [suggestionDemoArtifactId, setSuggestionDemoArtifactId] = useState<
    string | null
  >(null);
  const [assistantOpen, setAssistantOpen] = useState(false);
  const [assistantWidth, setAssistantWidth] = useState(readAssistantWidth);
  const [lastAgentRequest, setLastAgentRequest] =
    useState<EditorAgentRequest | null>(null);
  const [notice, setNotice] = useState<PlaygroundNotice | null>(null);
  const [activeImport, setActiveImport] =
    useState<PlaygroundImportState | null>(null);
  const assistantModal = useMediaQuery('(max-width: 1040px)');
  const fileInput = useRef<HTMLInputElement>(null);
  const pdfInput = useRef<HTMLInputElement>(null);
  const importCounter = useRef(0);
  const importController = useRef<{
    controller: AbortController;
    id: number;
    placeholderId?: string;
    restoreArtifactId: string | null;
  } | null>(null);
  const activeArtifact =
    artifacts.find((artifact) => artifact.id === activeArtifactId) ?? null;
  const docsUrl = documentationEntryUrl(document.baseURI);
  const collaborationDocsUrl = collaborationServerDocumentationUrl(
    document.baseURI,
  );

  useEffect(() => {
    const redirectLegacyDocsRoute = () => {
      const path = legacyDocsPath(window.location.hash);
      if (path) window.location.replace(new URL(path, document.baseURI));
    };
    window.addEventListener('hashchange', redirectLegacyDocsRoute);
    redirectLegacyDocsRoute();
    return () =>
      window.removeEventListener('hashchange', redirectLegacyDocsRoute);
  }, []);

  useEffect(
    () => () => {
      const pending = importController.current;
      importController.current = null;
      pending?.controller.abort();
    },
    [],
  );

  useEffect(() => {
    if (!notice) return;
    const timeout = window.setTimeout(() => setNotice(null), 3200);
    return () => window.clearTimeout(timeout);
  }, [notice]);

  const showNotice = useCallback(
    (message: string, tone: NoticeTone = 'neutral') => {
      setNotice({ id: Date.now(), message, tone });
    },
    [],
  );

  const openArtifact = (artifactId: string) => {
    setCollaborationDemoArtifactId(null);
    setSuggestionDemoArtifactId(null);
    activateArtifact(artifactId);
  };

  const activateArtifact = (artifactId: string) => {
    const now = Date.now();
    setArtifacts((current) =>
      current.map((artifact) =>
        artifact.id === artifactId
          ? { ...artifact, lastOpenedAt: now }
          : artifact,
      ),
    );
    setActiveArtifactId(artifactId);
    setLastAgentRequest(null);
    setAssistantOpen(false);
    closeSidebarForEditor();
  };

  const openCollaborationDemo = () => {
    const artifact = artifacts.find(({ kind }) => kind === 'document');
    if (!artifact) {
      showNotice('请先新建一个文字文档，再打开协作演示。', 'danger');
      return;
    }
    setCollaborationDemoArtifactId(artifact.id);
    setSuggestionDemoArtifactId(null);
    activateArtifact(artifact.id);
    showNotice('已以评论权限加入多人实时协作', 'success');
  };

  const openSuggestionDemo = () => {
    const artifact = artifacts.find(({ kind }) => kind === 'document');
    if (!artifact) {
      showNotice('请先新建一个文字文档，再打开建议协作演示。', 'danger');
      return;
    }
    setCollaborationDemoArtifactId(null);
    setSuggestionDemoArtifactId(artifact.id);
    activateArtifact(artifact.id);
    showNotice('已打开建议者与编辑者的实时协作', 'success');
  };

  const openFormattingReviewDemo = () => {
    const artifact = artifacts.find(({ kind }) => kind === 'document');
    if (!artifact || artifact.content.type !== 'document') {
      showNotice('请先新建一个文字文档，再打开格式修订演示。', 'danger');
      return;
    }
    const paragraphFormattingBefore = escapeHtmlAttribute(
      serializeDocumentParagraphFormatting({}),
    );
    const content: OfficeArtifactContent = {
      ...artifact.content,
      html: [
        '<section data-document-section="true">',
        '<h1>字符与段落格式修订</h1>',
        '<p>审阅者可以分别接受或拒绝字符格式与段落布局，正文内容始终独立保留。</p>',
        '<p><span data-testid="character-formatting-revision-demo" data-document-change="true" data-change-kind="formatting" data-change-id="playground-formatting-review" data-change-author="Ada Reviewer" data-change-date="2026-08-17T14:30:00.000Z" data-change-before="[]"><strong>这段文字新增了粗体格式</strong></span>，正文内容本身没有变化。</p>',
        `<p data-testid="paragraph-formatting-revision-demo" data-document-change="true" data-change-kind="paragraph-formatting" data-change-id="playground-paragraph-formatting-review" data-change-author="Lin Reviewer" data-change-date="2026-08-18T09:15:00.000Z" data-change-before="${paragraphFormattingBefore}" data-office-indent-level="2" data-office-space-after="18" data-office-line-rule="auto" data-office-auto-line-height="1.5" style="text-align: right; margin-left: 48px; margin-bottom: 18pt; line-height: 1.5">这段文字调整为右对齐、两级缩进和 1.5 倍行距，正文内容本身没有变化。</p>`,
        '</section>',
      ].join(''),
      model: undefined,
      trackChanges: true,
    };
    setArtifacts((current) =>
      current.map((candidate) =>
        candidate.id === artifact.id
          ? {
              ...candidate,
              content,
              revision: candidate.revision + 1,
              updatedAt: Date.now(),
            }
          : candidate,
      ),
    );
    setCollaborationDemoArtifactId(null);
    setSuggestionDemoArtifactId(null);
    activateArtifact(artifact.id);
    showNotice('已打开可接受或拒绝的字符与段落格式修订', 'success');
  };

  const newArtifact = (templateId: string) => {
    const artifact = createArtifact(templateId);
    setCollaborationDemoArtifactId(null);
    setSuggestionDemoArtifactId(null);
    setArtifacts((current) => [artifact, ...current]);
    setActiveArtifactId(artifact.id);
    setLastAgentRequest(null);
    setAssistantOpen(false);
    closeSidebarForEditor();
    showNotice(`${artifact.title} 已创建`, 'success');
  };

  const updateActiveArtifact = useCallback(
    (update: (artifact: OfficeArtifact) => OfficeArtifact) => {
      if (!activeArtifactId) return;
      setArtifacts((current) =>
        current.map((artifact) =>
          artifact.id === activeArtifactId ? update(artifact) : artifact,
        ),
      );
    },
    [activeArtifactId],
  );

  const importFile = async (file: File) => {
    const editorKind = workKindForFile(file);
    const editorPreload = editorKind
      ? preloadPlaygroundEditor(editorKind, {
          preloadRuntimeAssets: editorKind === 'pdf',
        }).catch(() => undefined)
      : Promise.resolve();
    const previousImport = importController.current;
    previousImport?.controller.abort();
    if (previousImport?.placeholderId) {
      setArtifacts((current) =>
        current.filter(
          (artifact) => artifact.id !== previousImport.placeholderId,
        ),
      );
      setActiveArtifactId((current) =>
        current === previousImport.placeholderId
          ? previousImport.restoreArtifactId
          : current,
      );
    }
    const id = importCounter.current + 1;
    importCounter.current = id;
    const controller = new AbortController();
    const restoreArtifactId = previousImport?.placeholderId
      ? previousImport.restoreArtifactId
      : activeArtifactId;
    const placeholderTemplate = importPlaceholderTemplate(editorKind);
    const placeholder = placeholderTemplate
      ? createArtifact(placeholderTemplate)
      : null;
    importController.current = {
      controller,
      id,
      ...(placeholder ? { placeholderId: placeholder.id } : {}),
      restoreArtifactId,
    };
    if (placeholder) {
      placeholder.title =
        file.name.replace(/\.[^.]+$/i, '') || placeholder.title;
      setArtifacts((current) => [
        placeholder,
        ...current.filter((artifact) => artifact.id !== placeholder.id),
      ]);
      setActiveArtifactId(placeholder.id);
      setLastAgentRequest(null);
      setAssistantOpen(false);
      closeSidebarForEditor();
    }
    const initialProgress: OfficeFileImportProgress = {
      stage: 'reading',
      stageProgress: 0,
      progress: 0,
      bytesRead: 0,
      totalBytes: file.size,
    };
    setActiveImport({ fileName: file.name, id, progress: initialProgress });
    try {
      const { importOfficeFile } = await import('@a3s-lab/office/core');
      const imported = await importOfficeFile(file, {
        ...(placeholder ? { artifactId: placeholder.id } : {}),
        ...(placeholder?.content.type === 'spreadsheet'
          ? {
              spreadsheetSheetIds: placeholder.content.sheets.flatMap(
                (sheet) => (sheet.id ? [sheet.id] : []),
              ),
            }
          : {}),
        signal: controller.signal,
        onProgress: (progress) => {
          if (importController.current?.id !== id) return;
          setActiveImport({ fileName: file.name, id, progress });
        },
      });
      if (editorKind === 'pdf') await editorPreload;
      if (importController.current?.id !== id) return;
      const opened = { ...imported, lastOpenedAt: Date.now() };
      setCollaborationDemoArtifactId(null);
      setSuggestionDemoArtifactId(null);
      setArtifacts((current) => [
        opened,
        ...current.filter((artifact) => artifact.id !== opened.id),
      ]);
      setActiveArtifactId(opened.id);
      setLastAgentRequest(null);
      setAssistantOpen(false);
      closeSidebarForEditor();
      showNotice(`${file.name} 已打开`, 'success');
    } catch (error) {
      if (importController.current?.id !== id) return;
      if (placeholder) {
        setArtifacts((current) =>
          current.filter((artifact) => artifact.id !== placeholder.id),
        );
        setActiveArtifactId((current) =>
          current === placeholder.id ? restoreArtifactId : current,
        );
      }
      if (error instanceof Error && error.name === 'AbortError') {
        showNotice(`已取消导入 ${file.name}`, 'neutral');
        return;
      }
      showNotice(
        error instanceof Error ? error.message : '无法打开这个文件',
        'danger',
      );
    } finally {
      if (importController.current?.id === id) {
        importController.current = null;
        setActiveImport(null);
      }
    }
  };

  return (
    <main
      className={`playground-site ${sidebarOpen ? 'sidebar-visible' : ''} ${
        activeArtifact ? 'editor-open' : ''
      }`}
    >
      <input
        ref={fileInput}
        className="playground-file-input"
        type="file"
        accept={OFFICE_FILE_ACCEPT}
        aria-label="打开 Office 或 PDF 文件"
        onChange={(event) => {
          const input = event.currentTarget;
          const file = input.files?.[0];
          const pendingImport = file ? importFile(file) : null;
          input.value = '';
          if (pendingImport) void pendingImport;
        }}
      />
      <input
        ref={pdfInput}
        className="playground-file-input"
        type="file"
        accept=".pdf,application/pdf"
        aria-label="打开 PDF 文件"
        onChange={(event) => {
          const input = event.currentTarget;
          const file = input.files?.[0];
          const pendingImport = file ? importFile(file) : null;
          input.value = '';
          if (pendingImport) void pendingImport;
        }}
      />

      {sidebarOpen && (
        <SiteSidebar
          docsUrl={docsUrl}
          modal={sidebarModal}
          onCollapse={() => setSidebarOpen(false)}
          onCreate={newArtifact}
          onHome={() => {
            setActiveArtifactId(null);
            setAssistantOpen(false);
            setLastAgentRequest(null);
            if (window.innerWidth < 840) setSidebarOpen(false);
          }}
          onOpenFile={() => fileInput.current?.click()}
          onOpenPdf={() => pdfInput.current?.click()}
        />
      )}
      {sidebarOpen && (
        <button
          type="button"
          className="playground-sidebar-scrim"
          aria-label="关闭办公侧边栏"
          aria-hidden="true"
          tabIndex={-1}
          onClick={() => setSidebarOpen(false)}
        />
      )}

      <section className="playground-main-pane">
        {activeArtifact ? (
          <Suspense
            fallback={<WorkEditorLoadingState title="正在加载编辑器" />}
          >
            <EditorWorkspace
              key={activeArtifact.id}
              artifact={activeArtifact}
              collaborationDemo={
                collaborationDemoArtifactId === activeArtifact.id
              }
              suggestionDemo={suggestionDemoArtifactId === activeArtifact.id}
              assistantModal={assistantModal}
              assistantOpen={assistantOpen}
              assistantWidth={assistantWidth}
              lastAgentRequest={lastAgentRequest}
              sidebarOpen={sidebarOpen}
              onAgentRequest={(request) => {
                setLastAgentRequest(request);
                setAssistantOpen(true);
              }}
              onAssistantWidthChange={(width) => {
                setAssistantWidth(width);
                persistAssistantWidth(width);
              }}
              onBack={() => {
                setActiveArtifactId(null);
                setCollaborationDemoArtifactId(null);
                setSuggestionDemoArtifactId(null);
                setAssistantOpen(false);
                setLastAgentRequest(null);
                if (window.innerWidth >= 840) setSidebarOpen(true);
              }}
              onChange={(content: OfficeArtifactContent) =>
                updateActiveArtifact((artifact) => ({
                  ...artifact,
                  content,
                  kind: content.type,
                  revision: artifact.revision + 1,
                  updatedAt: Date.now(),
                }))
              }
              onNotice={showNotice}
              onOpenSidebar={() => setSidebarOpen(true)}
              onRename={(title) =>
                updateActiveArtifact((artifact) => ({
                  ...artifact,
                  title,
                  revision: artifact.revision + 1,
                  updatedAt: Date.now(),
                }))
              }
              onToggleAssistant={() => setAssistantOpen((current) => !current)}
              onTouch={() =>
                updateActiveArtifact((artifact) => ({
                  ...artifact,
                  revision: artifact.revision + 1,
                  updatedAt: Date.now(),
                }))
              }
            />
          </Suspense>
        ) : (
          <WorkspaceHome
            artifacts={artifacts}
            collaborationDocsUrl={collaborationDocsUrl}
            sidebarOpen={sidebarOpen}
            onCreate={newArtifact}
            onImport={() => fileInput.current?.click()}
            onOpen={openArtifact}
            onOpenCollaborationDemo={openCollaborationDemo}
            onOpenFormattingReviewDemo={openFormattingReviewDemo}
            onOpenSuggestionDemo={openSuggestionDemo}
            onOpenPdf={() => pdfInput.current?.click()}
            onOpenSidebar={() => setSidebarOpen(true)}
          />
        )}
      </section>

      {activeImport && (
        <PlaygroundImportProgress
          state={activeImport}
          onCancel={() => importController.current?.controller.abort()}
        />
      )}

      {notice && <PlaygroundToast key={notice.id} notice={notice} />}
    </main>
  );
}

function importPlaceholderTemplate(
  kind: OfficeArtifactKind | null,
): 'blank-document' | 'blank-spreadsheet' | null {
  if (kind === 'document') return 'blank-document';
  if (kind === 'spreadsheet') return 'blank-spreadsheet';
  return null;
}

function PlaygroundToast({ notice }: { notice: PlaygroundNotice }) {
  return (
    <output className={`playground-toast ${notice.tone}`} aria-live="polite">
      {notice.tone === 'success' ? (
        <CheckCircle2 size={16} />
      ) : notice.tone === 'danger' ? (
        <AlertCircle size={16} />
      ) : (
        <Info size={16} />
      )}
      <span>{notice.message}</span>
    </output>
  );
}

function createInitialArtifacts(e2eFixture: string | null): OfficeArtifact[] {
  const project = createArtifact('project-brief');
  const plan = createArtifact('quarterly-plan');
  const deck = createArtifact('strategy-deck');
  const markdown = createArtifact('blank-markdown');
  markdown.title = '产品说明';
  markdown.content = {
    type: 'markdown',
    markdown: [
      '# A3S Office',
      '',
      '一套可嵌入 React、Vue 和 Web Component 的在线 Office 编辑器。',
      '',
      '## 接入检查',
      '',
      '- [x] 安装 `@a3s-lab/office`',
      '- [x] 引入全局样式',
      '- [ ] 接入宿主应用的持久化',
      '',
      '## 编辑器',
      '',
      '| 类型 | 组件 |',
      '| --- | --- |',
      '| 文字 | `DocumentEditor` |',
      '| Markdown | `MarkdownEditor` |',
      '',
    ].join('\n'),
  };
  const now = Date.now();
  const artifacts = [
    { ...project, lastOpenedAt: now - 1_000 },
    { ...plan, lastOpenedAt: now - 2_000 },
    { ...deck, lastOpenedAt: now - 3_000 },
    { ...markdown, lastOpenedAt: now - 4_000 },
  ];
  if (e2eFixture === MAXIMUM_SPARSE_SPREADSHEET_FIXTURE) {
    return [createMaximumSparseSpreadsheetArtifact(), ...artifacts];
  }
  if (e2eFixture === SPREADSHEET_GO_TO_FIXTURE) {
    return [createSpreadsheetGoToArtifact(), ...artifacts];
  }
  if (e2eFixture === SPREADSHEET_DATA_VALIDATION_FIXTURE) {
    return [createSpreadsheetDataValidationArtifact(), ...artifacts];
  }
  if (e2eFixture === SPREADSHEET_HYPERLINK_FIXTURE) {
    return [createSpreadsheetHyperlinkArtifact(), ...artifacts];
  }
  if (e2eFixture === SPREADSHEET_PASTE_SPECIAL_FIXTURE) {
    return [createSpreadsheetPasteSpecialArtifact(), ...artifacts];
  }
  if (e2eFixture === SPREADSHEET_COPY_FROM_ABOVE_FIXTURE) {
    return [createSpreadsheetCopyFromAboveArtifact(), ...artifacts];
  }
  if (e2eFixture === SPREADSHEET_DATE_TIME_FIXTURE) {
    return [createSpreadsheetDateTimeArtifact(), ...artifacts];
  }
  if (e2eFixture === SPREADSHEET_GRADIENT_FILL_FIXTURE) {
    return [createSpreadsheetGradientFillArtifact(), ...artifacts];
  }
  if (e2eFixture === SPREADSHEET_PATTERN_FILL_FIXTURE) {
    return [createSpreadsheetPatternFillArtifact(), ...artifacts];
  }
  if (e2eFixture === SPREADSHEET_RICH_TEXT_FIXTURE) {
    return [createSpreadsheetRichTextArtifact(), ...artifacts];
  }
  return artifacts;
}

function readPlaygroundE2eFixture(): string | null {
  return typeof window === 'undefined'
    ? null
    : new URLSearchParams(window.location.search).get('e2e');
}

function escapeHtmlAttribute(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}

function readAssistantWidth(): number {
  try {
    const value = Number(localStorage.getItem('a3s-office.assistant-width'));
    return Number.isFinite(value) && value >= 340 ? value : 460;
  } catch {
    return 460;
  }
}

function persistAssistantWidth(width: number): void {
  try {
    localStorage.setItem('a3s-office.assistant-width', String(width));
  } catch {
    // Resizing still works for the current session when storage is unavailable.
  }
}

const root = document.getElementById('root');
if (!root) throw new Error('Playground root element is missing.');

createRoot(root).render(
  <StrictMode>
    <Playground />
  </StrictMode>,
);
