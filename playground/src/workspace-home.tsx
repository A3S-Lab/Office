import type {
  OfficeArtifact,
  OfficeArtifactKind,
  OfficeTemplate,
} from '@a3s-lab/office/core';
import {
  ArrowRight,
  FileDiff,
  FilePlus2,
  Languages,
  LayoutGrid,
  ListChecks,
  ListOrdered,
  ListTree,
  PanelLeftOpen,
  Pencil,
  Play,
  Search,
  ServerCog,
  Sparkles,
  Upload,
  UsersRound,
} from 'lucide-react';
import { useMemo, useState } from 'react';
import { WORK_TEMPLATES as officeTemplates } from '../../src/internal/features/work/work-templates';
import { warmPlaygroundEditor } from './editor-preload';
import { FileKindIcon, fileKindExtension, fileKindLabel } from './file-kind';
import { LATEST_CAPABILITIES } from './latest-capabilities';

const templateCellIds = Array.from(
  { length: 20 },
  (_, index) => `cell-${index + 1}`,
);
type LatestCapabilityFilter = 'all' | OfficeArtifactKind;

const latestCapabilityKinds: OfficeArtifactKind[] = [
  'document',
  'spreadsheet',
  'presentation',
  'pdf',
  'markdown',
];

export function WorkspaceHome({
  artifacts,
  collaborationDocsUrl,
  sidebarOpen,
  onOpenSidebar,
  onCreate,
  onOpen,
  onImport,
  onOpenCollaborationDemo,
  onOpenSuggestionDemo,
  onOpenFormattingReviewDemo,
  onOpenPdf,
}: {
  artifacts: OfficeArtifact[];
  collaborationDocsUrl: string;
  sidebarOpen: boolean;
  onOpenSidebar: () => void;
  onCreate: (templateId: string) => void;
  onOpen: (artifactId: string) => void;
  onImport: () => void;
  onOpenCollaborationDemo: () => void;
  onOpenSuggestionDemo: () => void;
  onOpenFormattingReviewDemo: () => void;
  onOpenPdf: () => void;
}) {
  const [query, setQuery] = useState('');
  const [latestFilter, setLatestFilter] =
    useState<LatestCapabilityFilter>('all');
  const visibleArtifacts = useMemo(() => {
    const normalized = query.trim().toLocaleLowerCase();
    return [...artifacts]
      .filter(
        (artifact) =>
          !normalized ||
          artifact.title.toLocaleLowerCase().includes(normalized) ||
          fileKindLabel(artifact.kind).toLocaleLowerCase().includes(normalized),
      )
      .sort((left, right) => right.lastOpenedAt - left.lastOpenedAt);
  }, [artifacts, query]);
  const latestFilters = useMemo(
    () => [
      { count: LATEST_CAPABILITIES.length, id: 'all' as const },
      ...latestCapabilityKinds
        .map((kind) => ({
          count: LATEST_CAPABILITIES.filter(
            (capability) => capability.kind === kind,
          ).length,
          id: kind,
        }))
        .filter(({ count }) => count > 0),
    ],
    [],
  );
  const visibleLatestTemplates = useMemo(
    () =>
      latestFilter === 'all'
        ? LATEST_CAPABILITIES
        : LATEST_CAPABILITIES.filter(
            (capability) => capability.kind === latestFilter,
          ),
    [latestFilter],
  );

  return (
    <section className="playground-home">
      <header className="playground-home-header">
        <div className="playground-home-title">
          {!sidebarOpen && (
            <button
              type="button"
              className="playground-icon-button sidebar-open"
              aria-label="展开办公侧边栏"
              title="展开侧边栏"
              onClick={onOpenSidebar}
            >
              <PanelLeftOpen size={17} />
            </button>
          )}
          <div>
            <span>A3S Office</span>
            <h1>我的文档</h1>
          </div>
        </div>
        <div className="playground-home-actions">
          <label className="playground-search">
            <Search size={15} />
            <span className="sr-only">搜索文件</span>
            <input
              value={query}
              placeholder="搜索文件"
              onChange={(event) => setQuery(event.target.value)}
            />
          </label>
          <button
            type="button"
            className="playground-secondary-button"
            onClick={onImport}
          >
            <Upload size={15} />
            打开文件
          </button>
        </div>
      </header>

      <section
        className="playground-latest-capabilities"
        aria-labelledby="playground-latest-capabilities-title"
      >
        <div className="playground-latest-capabilities-heading">
          <div>
            <div className="playground-latest-capabilities-title-row">
              <h2 id="playground-latest-capabilities-title">最新能力</h2>
              <span>
                <Sparkles size={13} aria-hidden="true" />
                main 已部署
              </span>
            </div>
            <p>按编辑器筛选近期发布的完整可编辑工作流。</p>
          </div>
          <output aria-live="polite">
            {visibleLatestTemplates.length} / {LATEST_CAPABILITIES.length} 项
          </output>
        </div>
        <fieldset className="playground-latest-capability-filters">
          <legend className="sr-only">按编辑器筛选最新能力</legend>
          {latestFilters.map(({ count, id }) => (
            <button
              type="button"
              aria-controls="playground-latest-capability-list"
              aria-pressed={latestFilter === id}
              key={id}
              onClick={() => setLatestFilter(id)}
            >
              <span>{latestCapabilityFilterLabel(id)}</span>
              <small>{count}</small>
            </button>
          ))}
        </fieldset>
        <div
          className="playground-latest-capability-list"
          id="playground-latest-capability-list"
        >
          {visibleLatestTemplates.map((capability) => (
            <button
              type="button"
              className={`playground-latest-capability ${capability.kind}`}
              aria-label={`打开最新能力：${capability.name}`}
              key={capability.id}
              onFocus={() => warmPlaygroundEditor(capability.kind)}
              onClick={() => {
                if (capability.launch.type === 'pdf-page-organization') {
                  onOpenPdf();
                  return;
                }
                onCreate(capability.launch.templateId);
              }}
              onPointerEnter={() => warmPlaygroundEditor(capability.kind)}
            >
              <span className="playground-latest-capability-icon">
                <LatestCapabilityIcon templateId={capability.id} />
              </span>
              <span className="playground-latest-capability-copy">
                <span className="playground-latest-capability-meta">
                  <small>{latestCapabilityEditorLabel(capability.kind)}</small>
                  <small>v{capability.release}</small>
                </span>
                <strong>{capability.name}</strong>
                <span className="playground-latest-capability-description">
                  {capability.description}
                </span>
              </span>
              <ArrowRight size={14} aria-hidden="true" />
            </button>
          ))}
        </div>
      </section>

      <section
        className="playground-collaboration-overview"
        aria-labelledby="playground-collaboration-title"
      >
        <div
          className="playground-collaboration-participants"
          aria-hidden="true"
        >
          <span className="human">林</span>
          <span className="agent">A3S</span>
          <i />
        </div>
        <div className="playground-collaboration-copy">
          <h2 id="playground-collaboration-title">多人实时协作</h2>
          <p>
            体验评论线程、文字建议、字符与段落格式修订；在实时视图中提交、接受、
            拒绝和审计改动，同时查看参与者、远端选区和在线状态。
          </p>
          <span className="playground-collaboration-backend-note">
            生产后端示例基于 A3S Boot，并在持久化前校验评论与建议权限。
          </span>
        </div>
        <div className="playground-collaboration-actions">
          <button
            type="button"
            className="playground-primary-button"
            onClick={onOpenCollaborationDemo}
          >
            <UsersRound size={15} />
            体验实时评论
          </button>
          <button
            type="button"
            className="playground-secondary-button"
            onClick={onOpenSuggestionDemo}
          >
            <Pencil size={15} />
            体验建议协作
          </button>
          <button
            type="button"
            className="playground-secondary-button"
            onClick={onOpenFormattingReviewDemo}
          >
            <FileDiff size={15} />
            体验格式修订
          </button>
          <a
            className="playground-secondary-button"
            href={collaborationDocsUrl}
          >
            <ServerCog size={15} />
            查看 A3S Boot 后端
          </a>
        </div>
      </section>

      <section
        className="playground-template-section"
        aria-labelledby="playground-create-title"
      >
        <div className="playground-section-heading">
          <div>
            <h2 id="playground-create-title">新建</h2>
            <span>选择空白文件或模板</span>
          </div>
        </div>
        <div className="playground-template-grid">
          {officeTemplates.map((template) => (
            <TemplateCard
              key={template.id}
              template={template}
              onCreate={() => onCreate(template.id)}
            />
          ))}
          <button
            type="button"
            className="playground-template-card pdf"
            onFocus={() => warmPlaygroundEditor('pdf')}
            onClick={onOpenPdf}
            onPointerEnter={() => warmPlaygroundEditor('pdf')}
          >
            <span className="playground-template-preview pdf">
              <span className="template-document-sheet">
                <b className="template-pdf-mark">PDF</b>
              </span>
              <span className="template-open-badge">打开</span>
            </span>
            <span className="playground-template-copy">
              <strong>PDF 编辑器</strong>
              <small>查看、批注并保存 PDF</small>
            </span>
          </button>
        </div>
      </section>

      <section
        className="playground-recent-section"
        aria-labelledby="playground-recent-title"
      >
        <div className="playground-section-heading">
          <div>
            <h2 id="playground-recent-title">最近文件</h2>
            <span>{visibleArtifacts.length} 个文件</span>
          </div>
          {artifacts.length > 0 && (
            <button
              type="button"
              className="playground-text-button"
              onClick={() => setQuery('')}
            >
              查看全部
              <ArrowRight size={13} />
            </button>
          )}
        </div>

        {visibleArtifacts.length > 0 ? (
          <div className="playground-artifact-grid">
            {visibleArtifacts.map((artifact) => (
              <button
                type="button"
                className={`playground-artifact-card ${artifact.kind}`}
                key={artifact.id}
                onFocus={() => warmPlaygroundEditor(artifact.kind)}
                onClick={() => onOpen(artifact.id)}
                onPointerEnter={() => warmPlaygroundEditor(artifact.kind)}
              >
                <ArtifactPreview kind={artifact.kind} />
                <span className="playground-artifact-copy">
                  <strong>{artifact.title}</strong>
                  <small>{fileKindExtension(artifact.kind)} · 本次会话</small>
                </span>
                <span className={`playground-artifact-kind ${artifact.kind}`}>
                  <FileKindIcon kind={artifact.kind} size={14} />
                </span>
              </button>
            ))}
          </div>
        ) : (
          <div className="playground-empty-state">
            <span>
              <FilePlus2 size={21} />
            </span>
            <strong>{query ? '没有匹配的文件' : '还没有打开文件'}</strong>
            <p>
              {query
                ? '换一个名称或文件类型试试。'
                : '从上方模板开始，或打开已有的 Office 与 PDF 文件。'}
            </p>
            <button
              type="button"
              className="playground-primary-button"
              onClick={query ? () => setQuery('') : onImport}
            >
              {query ? '清除搜索' : '打开文件'}
            </button>
          </div>
        )}
      </section>
    </section>
  );
}

function latestCapabilityFilterLabel(filter: LatestCapabilityFilter): string {
  if (filter === 'all') return '全部';
  if (filter === 'document') return '文字';
  if (filter === 'spreadsheet') return '表格';
  if (filter === 'presentation') return '演示';
  if (filter === 'pdf') return 'PDF';
  return 'Markdown';
}

function latestCapabilityEditorLabel(kind: OfficeArtifactKind): string {
  if (kind === 'document') return 'Writer';
  if (kind === 'spreadsheet') return 'Spreadsheet';
  if (kind === 'presentation') return 'Presentation';
  if (kind === 'pdf') return 'PDF';
  return 'Markdown';
}

function LatestCapabilityIcon({ templateId }: { templateId: string }) {
  if (templateId === 'pdf-page-organization') {
    return <LayoutGrid size={18} aria-hidden="true" />;
  }
  if (templateId === 'document-comparison') {
    return <FileDiff size={18} aria-hidden="true" />;
  }
  if (templateId === 'table-of-contents') {
    return <ListTree size={19} aria-hidden="true" />;
  }
  if (templateId === 'document-index') {
    return <ListOrdered size={18} aria-hidden="true" />;
  }
  if (templateId === 'proofing-languages') {
    return <Languages size={16} aria-hidden="true" />;
  }
  if (templateId === 'data-validation') {
    return <ListChecks size={16} aria-hidden="true" />;
  }
  if (templateId === 'animated-deck') {
    return <Play size={17} aria-hidden="true" />;
  }
  return <Sparkles size={16} aria-hidden="true" />;
}

function TemplateCard({
  template,
  onCreate,
}: {
  template: OfficeTemplate;
  onCreate: () => void;
}) {
  return (
    <button
      type="button"
      className={`playground-template-card ${template.kind}`}
      style={
        {
          '--template-accent': template.accent,
        } as React.CSSProperties
      }
      onFocus={() => warmPlaygroundEditor(template.kind)}
      onClick={onCreate}
      onPointerEnter={() => warmPlaygroundEditor(template.kind)}
    >
      <span className={`playground-template-preview ${template.kind}`}>
        <TemplateArtwork kind={template.kind} />
      </span>
      <span className="playground-template-copy">
        <strong>{template.name}</strong>
        <small>{template.description}</small>
      </span>
    </button>
  );
}

function TemplateArtwork({ kind }: { kind: OfficeArtifactKind }) {
  if (kind === 'spreadsheet') {
    return (
      <span className="template-sheet-grid" aria-hidden="true">
        {templateCellIds.map((cellId) => (
          <i key={`${kind}-${cellId}`} />
        ))}
      </span>
    );
  }
  if (kind === 'presentation') {
    return (
      <span className="template-slide" aria-hidden="true">
        <b />
        <i />
        <i />
      </span>
    );
  }
  if (kind === 'markdown') {
    return (
      <span className="template-document-sheet markdown" aria-hidden="true">
        <b>#</b>
        <i />
        <i />
        <i />
      </span>
    );
  }
  return (
    <span className="template-document-sheet" aria-hidden="true">
      <b />
      <i />
      <i />
      <i />
      <i />
    </span>
  );
}

function ArtifactPreview({ kind }: { kind: OfficeArtifactKind }) {
  if (kind === 'spreadsheet') {
    return (
      <span className="artifact-preview">
        <span className="template-sheet-grid">
          {templateCellIds.map((cellId) => (
            <i key={`${kind}-recent-${cellId}`} />
          ))}
        </span>
      </span>
    );
  }
  if (kind === 'presentation') {
    return (
      <span className="artifact-preview">
        <span className="template-slide">
          <b />
          <i />
          <i />
        </span>
      </span>
    );
  }
  if (kind === 'pdf') {
    return (
      <span className="artifact-preview">
        <span className="template-document-sheet pdf">
          <b className="template-pdf-mark">PDF</b>
        </span>
      </span>
    );
  }
  if (kind === 'markdown') {
    return (
      <span className="artifact-preview">
        <span className="template-document-sheet markdown">
          <b>#</b>
          <i />
          <i />
          <i />
        </span>
      </span>
    );
  }
  return (
    <span className="artifact-preview">
      <span className="template-document-sheet">
        <b />
        <i />
        <i />
        <i />
        <i />
      </span>
    </span>
  );
}
