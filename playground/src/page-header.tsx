import { PanelLeftOpen } from 'lucide-react';
import type { ReactNode } from 'react';

export function PageHeader({
  eyebrow,
  title,
  description,
  sidebarOpen,
  actions,
  onOpenSidebar,
}: {
  eyebrow: string;
  title: string;
  description?: string;
  sidebarOpen: boolean;
  actions?: ReactNode;
  onOpenSidebar: () => void;
}) {
  return (
    <header className="playground-page-header">
      <div className="playground-page-heading">
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
          <span>{eyebrow}</span>
          <h1>{title}</h1>
          {description && <p>{description}</p>}
        </div>
      </div>
      {actions && <div className="playground-page-actions">{actions}</div>}
    </header>
  );
}
