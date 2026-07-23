import { AlertCircle, CheckCircle2, Info } from 'lucide-react';
import { StrictMode, useCallback, useEffect, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import {
  createArtifact,
  importOfficeFile,
  OFFICE_FILE_ACCEPT,
  type EditorAgentRequest,
  type OfficeArtifact,
  type OfficeArtifactContent,
} from '@a3s-lab/office/core';
import '@a3s-lab/office/styles.css';
import { CliDocsPage } from './cli-docs-page';
import { EditorWorkspace } from './editor-workspace';
import type {
  NoticeTone,
  PlaygroundNotice,
  SiteRoute,
} from './playground-types';
import { SiteSidebar } from './site-sidebar';
import { SkillDownloadPage } from './skill-download-page';
import { WorkspaceHome } from './workspace-home';
import './playground.css';
import './workspace.css';
import './docs-pages.css';

function Playground() {
  const [route, setRoute] = useState<SiteRoute>(readRoute);
  const [sidebarOpen, setSidebarOpen] = useState(
    () => window.innerWidth >= 840,
  );
  const [artifacts, setArtifacts] = useState<OfficeArtifact[]>(
    createInitialArtifacts,
  );
  const [activeArtifactId, setActiveArtifactId] = useState<string | null>(null);
  const [assistantOpen, setAssistantOpen] = useState(false);
  const [assistantWidth, setAssistantWidth] = useState(readAssistantWidth);
  const [lastAgentRequest, setLastAgentRequest] =
    useState<EditorAgentRequest | null>(null);
  const [notice, setNotice] = useState<PlaygroundNotice | null>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  const pdfInput = useRef<HTMLInputElement>(null);
  const activeArtifact =
    artifacts.find((artifact) => artifact.id === activeArtifactId) ?? null;

  useEffect(() => {
    const onHashChange = () => setRoute(readRoute());
    window.addEventListener('hashchange', onHashChange);
    return () => window.removeEventListener('hashchange', onHashChange);
  }, []);

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

  const navigate = (nextRoute: SiteRoute) => {
    setRoute(nextRoute);
    const nextHash = `#${nextRoute}`;
    if (window.location.hash !== nextHash) {
      window.history.pushState(null, '', nextHash);
    }
    if (window.innerWidth < 840) setSidebarOpen(false);
  };

  const openArtifact = (artifactId: string) => {
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
    setSidebarOpen(false);
    navigate('office');
  };

  const newArtifact = (templateId: string) => {
    const artifact = createArtifact(templateId);
    setArtifacts((current) => [artifact, ...current]);
    setActiveArtifactId(artifact.id);
    setLastAgentRequest(null);
    setAssistantOpen(false);
    setSidebarOpen(false);
    navigate('office');
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
    try {
      const imported = await importOfficeFile(file);
      const opened = { ...imported, lastOpenedAt: Date.now() };
      setArtifacts((current) => [
        opened,
        ...current.filter((artifact) => artifact.id !== opened.id),
      ]);
      setActiveArtifactId(opened.id);
      setLastAgentRequest(null);
      setAssistantOpen(false);
      setSidebarOpen(false);
      navigate('office');
      showNotice(`${file.name} 已打开`, 'success');
    } catch (error) {
      showNotice(
        error instanceof Error ? error.message : '无法打开这个文件',
        'danger',
      );
    }
  };

  const assetUrl = useCallback(
    (fileName: string) => new URL(fileName, document.baseURI).href,
    [],
  );

  return (
    <main className={`playground-site ${sidebarOpen ? 'sidebar-visible' : ''}`}>
      <input
        ref={fileInput}
        className="playground-file-input"
        type="file"
        accept={OFFICE_FILE_ACCEPT}
        aria-label="打开 Office 或 PDF 文件"
        onChange={(event) => {
          const file = event.target.files?.[0];
          event.target.value = '';
          if (file) void importFile(file);
        }}
      />
      <input
        ref={pdfInput}
        className="playground-file-input"
        type="file"
        accept=".pdf,application/pdf"
        aria-label="打开 PDF 文件"
        onChange={(event) => {
          const file = event.target.files?.[0];
          event.target.value = '';
          if (file) void importFile(file);
        }}
      />

      {sidebarOpen && (
        <SiteSidebar
          route={route}
          onCollapse={() => setSidebarOpen(false)}
          onCreate={newArtifact}
          onNavigate={navigate}
          onOpenFile={() => fileInput.current?.click()}
          onOpenPdf={() => pdfInput.current?.click()}
        />
      )}
      {sidebarOpen && (
        <button
          type="button"
          className="playground-sidebar-scrim"
          aria-label="关闭办公侧边栏"
          onClick={() => setSidebarOpen(false)}
        />
      )}

      <section className="playground-main-pane">
        {route === 'office' &&
          (activeArtifact ? (
            <EditorWorkspace
              key={activeArtifact.id}
              artifact={activeArtifact}
              assetUrl={assetUrl}
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
          ) : (
            <WorkspaceHome
              artifacts={artifacts}
              sidebarOpen={sidebarOpen}
              onCreate={newArtifact}
              onImport={() => fileInput.current?.click()}
              onOpen={openArtifact}
              onOpenPdf={() => pdfInput.current?.click()}
              onOpenSidebar={() => setSidebarOpen(true)}
            />
          ))}
        {route === 'cli' && (
          <CliDocsPage
            sidebarOpen={sidebarOpen}
            onOpenSidebar={() => setSidebarOpen(true)}
            onOpenSkill={() => navigate('skill')}
          />
        )}
        {route === 'skill' && (
          <SkillDownloadPage
            rawSkillUrl={assetUrl('downloads/a3s-office-skill/SKILL.md')}
            sidebarOpen={sidebarOpen}
            skillDownloadUrl={assetUrl('downloads/a3s-office-skill.tar.gz')}
            onOpenCli={() => navigate('cli')}
            onOpenSidebar={() => setSidebarOpen(true)}
          />
        )}
      </section>

      {notice && <PlaygroundToast key={notice.id} notice={notice} />}
    </main>
  );
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

function readRoute(): SiteRoute {
  const route = window.location.hash.slice(1).split('/')[0];
  if (route === 'cli' || route === 'skill') return route;
  return 'office';
}

function createInitialArtifacts(): OfficeArtifact[] {
  const project = createArtifact('project-brief');
  const plan = createArtifact('quarterly-plan');
  const deck = createArtifact('strategy-deck');
  const markdown = createArtifact('blank-markdown');
  markdown.title = '产品说明';
  markdown.content = {
    type: 'markdown',
    markdown:
      '# A3S Office\n\n一套可嵌入 React、Vue 和 Web Component 的在线 Office 编辑器。\n',
  };
  const now = Date.now();
  return [
    { ...project, lastOpenedAt: now - 1_000 },
    { ...plan, lastOpenedAt: now - 2_000 },
    { ...deck, lastOpenedAt: now - 3_000 },
    { ...markdown, lastOpenedAt: now - 4_000 },
  ];
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
