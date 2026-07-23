import {
  ArrowRight,
  FilePlus2,
  FileType2,
  PanelLeftOpen,
  Search,
  Upload,
} from 'lucide-react';
import { useMemo, useState } from 'react';
import {
  officeTemplates,
  type OfficeArtifact,
  type OfficeArtifactKind,
  type OfficeTemplate,
} from '@a3s-lab/office/core';
import { preloadOfficeEditor } from '@a3s-lab/office/react';
import { FileKindIcon, fileKindExtension, fileKindLabel } from './file-kind';

const templateCellIds = Array.from(
  { length: 20 },
  (_, index) => `cell-${index + 1}`,
);

function warmOfficeEditor(kind: OfficeArtifactKind): void {
  void preloadOfficeEditor(kind).catch(() => undefined);
}

export function WorkspaceHome({
  artifacts,
  sidebarOpen,
  onOpenSidebar,
  onCreate,
  onOpen,
  onImport,
  onOpenPdf,
}: {
  artifacts: OfficeArtifact[];
  sidebarOpen: boolean;
  onOpenSidebar: () => void;
  onCreate: (templateId: string) => void;
  onOpen: (artifactId: string) => void;
  onImport: () => void;
  onOpenPdf: () => void;
}) {
  const [query, setQuery] = useState('');
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
            onFocus={() => warmOfficeEditor('pdf')}
            onClick={onOpenPdf}
            onPointerEnter={() => warmOfficeEditor('pdf')}
          >
            <span className="playground-template-preview pdf">
              <span className="template-document-sheet">
                <FileType2 size={28} />
                <b>PDF</b>
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
                onFocus={() => warmOfficeEditor(artifact.kind)}
                onClick={() => onOpen(artifact.id)}
                onPointerEnter={() => warmOfficeEditor(artifact.kind)}
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
      onFocus={() => warmOfficeEditor(template.kind)}
      onClick={onCreate}
      onPointerEnter={() => warmOfficeEditor(template.kind)}
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
          <FileType2 size={25} />
          <b>PDF</b>
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
