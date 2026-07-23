import {
  BookOpen,
  ExternalLink,
  FileCode2,
  FileSpreadsheet,
  FileText,
  FileType2,
  FolderOpen,
  Github,
  HardDrive,
  Package,
  PanelLeftClose,
  Presentation,
} from 'lucide-react';
import {
  type OfficeEditorKind,
  preloadOfficeEditor,
} from '@a3s-lab/office/react';
import type { SiteRoute } from './playground-types';

function warmOfficeEditor(kind: OfficeEditorKind): void {
  void preloadOfficeEditor(kind).catch(() => undefined);
}

export function SiteSidebar({
  route,
  onCollapse,
  onNavigate,
  onCreate,
  onOpenFile,
  onOpenPdf,
}: {
  route: SiteRoute;
  onCollapse: () => void;
  onNavigate: (route: SiteRoute) => void;
  onCreate: (templateId: string) => void;
  onOpenFile: () => void;
  onOpenPdf: () => void;
}) {
  return (
    <aside className="playground-sidebar" aria-label="A3S Office 导航">
      <header className="sidebar-product-header">
        <strong>办公</strong>
        <button
          type="button"
          className="playground-icon-button"
          aria-label="收起办公侧边栏"
          title="收起侧边栏"
          onClick={onCollapse}
        >
          <PanelLeftClose size={16} />
        </button>
      </header>

      <section className="playground-workspace-card" aria-label="当前工作区">
        <span className="playground-workspace-label">工作区</span>
        <div>
          <span className="playground-workspace-icon">
            <HardDrive size={17} />
          </span>
          <span>
            <strong>在线 Playground</strong>
            <small>浏览器本地会话</small>
          </span>
        </div>
      </section>

      <nav className="playground-primary-nav" aria-label="产品页面">
        <span className="playground-sidebar-label">产品</span>
        <button
          type="button"
          className={route === 'office' ? 'active' : ''}
          aria-current={route === 'office' ? 'page' : undefined}
          onClick={() => onNavigate('office')}
        >
          <FolderOpen size={16} />
          <span>编辑器</span>
        </button>
        <button
          type="button"
          className={route === 'cli' ? 'active' : ''}
          aria-current={route === 'cli' ? 'page' : undefined}
          onClick={() => onNavigate('cli')}
        >
          <BookOpen size={16} />
          <span>Office CLI</span>
        </button>
        <button
          type="button"
          className={route === 'skill' ? 'active' : ''}
          aria-current={route === 'skill' ? 'page' : undefined}
          onClick={() => onNavigate('skill')}
        >
          <Package size={16} />
          <span>CLI Skill</span>
        </button>
      </nav>

      <section className="playground-quick-create" aria-label="快速新建">
        <span className="playground-sidebar-label">快速新建</span>
        <button
          type="button"
          onFocus={() => warmOfficeEditor('document')}
          onClick={() => onCreate('blank-document')}
          onPointerEnter={() => warmOfficeEditor('document')}
        >
          <span className="quick-create-icon document">
            <FileText size={15} />
          </span>
          <span>文字</span>
        </button>
        <button
          type="button"
          onFocus={() => warmOfficeEditor('spreadsheet')}
          onClick={() => onCreate('blank-spreadsheet')}
          onPointerEnter={() => warmOfficeEditor('spreadsheet')}
        >
          <span className="quick-create-icon spreadsheet">
            <FileSpreadsheet size={15} />
          </span>
          <span>表格</span>
        </button>
        <button
          type="button"
          onFocus={() => warmOfficeEditor('presentation')}
          onClick={() => onCreate('blank-presentation')}
          onPointerEnter={() => warmOfficeEditor('presentation')}
        >
          <span className="quick-create-icon presentation">
            <Presentation size={15} />
          </span>
          <span>演示</span>
        </button>
        <button
          type="button"
          onFocus={() => warmOfficeEditor('pdf')}
          onClick={onOpenPdf}
          onPointerEnter={() => warmOfficeEditor('pdf')}
        >
          <span className="quick-create-icon pdf">
            <FileType2 size={15} />
          </span>
          <span>PDF</span>
          <small>打开</small>
        </button>
        <button
          type="button"
          onFocus={() => warmOfficeEditor('markdown')}
          onClick={() => onCreate('blank-markdown')}
          onPointerEnter={() => warmOfficeEditor('markdown')}
        >
          <span className="quick-create-icon markdown">
            <FileCode2 size={15} />
          </span>
          <span>Markdown</span>
        </button>
      </section>

      <div className="playground-sidebar-footer">
        <button
          type="button"
          className="playground-open-file"
          onClick={onOpenFile}
        >
          <FolderOpen size={15} />
          打开文件
        </button>
        <a
          href="https://github.com/A3S-Lab/Office"
          target="_blank"
          rel="noreferrer"
        >
          <Github size={15} />
          <span>GitHub</span>
          <ExternalLink size={12} />
        </a>
      </div>
    </aside>
  );
}
