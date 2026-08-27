import { Fragment } from 'react';
import type { ReactNode } from 'react';

export type HomeLanguage = 'zh' | 'en';
export type ChapterKind =
  | 'document'
  | 'markdown'
  | 'spreadsheet'
  | 'presentation'
  | 'pdf';

export const chapterOrder: readonly ChapterKind[] = [
  'document',
  'markdown',
  'spreadsheet',
  'presentation',
  'pdf',
];

export function MotionArrow() {
  return (
    <svg
      aria-hidden="true"
      className="office-editor-chapters__arrow"
      viewBox="0 0 16 16"
    >
      <path d="M3 8h10" />
      <path d="m9 4 4 4-4 4" />
    </svg>
  );
}

export function SyncIcon() {
  return (
    <svg aria-hidden="true" viewBox="0 0 20 16">
      <path d="M2 5h13l-3-3" />
      <path d="m15 5-3 3" />
      <path d="M18 11H5l3 3" />
      <path d="m5 11 3-3" />
    </svg>
  );
}

type MotionCommandKind =
  | 'undo'
  | 'redo'
  | 'paste'
  | 'paint'
  | 'font'
  | 'align'
  | 'grid'
  | 'filter'
  | 'source'
  | 'preview'
  | 'link'
  | 'slide'
  | 'arrange'
  | 'play'
  | 'rotate'
  | 'comment'
  | 'pages'
  | 'zoom'
  | 'share'
  | 'search'
  | 'save'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'table'
  | 'image'
  | 'more';

interface MotionCommand {
  label: string;
  icon: MotionCommandKind;
  active?: boolean;
}

function MotionCommandGlyph({ kind }: { kind: MotionCommandKind }) {
  const paths: Record<MotionCommandKind, ReactNode> = {
    undo: <path d="M7 5 3 8l4 3M3 8h7a4 4 0 0 1 4 4" />,
    redo: <path d="m9 5 4 3-4 3m4-3H6a4 4 0 0 0-4 4" />,
    paste: <path d="M6 4h8v10H6zM4 2h6v2H4zM2 6h2v8h2" />,
    paint: <path d="m4 3 8 8M3 12h5M10 2l3 3-5 5-3-3z" />,
    font: <path d="M4 13 8 3l4 10M5.5 9h5M13 4v9M11 6h4" />,
    align: <path d="M3 4h10M3 8h7M3 12h10" />,
    grid: <path d="M3 3h10v10H3zM3 7h10M7 3v10" />,
    filter: <path d="M3 4h10l-4 4v4l-2 1V8z" />,
    source: <path d="m6 4-3 4 3 4M10 4l3 4-3 4M9 3 7 13" />,
    preview: (
      <path d="M2 8s2-4 6-4 6 4 6 4-2 4-6 4-6-4-6-4zM8 6a2 2 0 1 0 0 4 2 2 0 0 0-4" />
    ),
    link: (
      <path d="M6 9 4 11a2 2 0 0 1-3-3l2-2a2 2 0 0 1 3 0M8 7l2-2a2 2 0 0 1 3 3l-2 2a2 2 0 0 1-3 0M5 8l3-3" />
    ),
    slide: <path d="M2 3h12v9H2zM5 14h6M8 12v2" />,
    arrange: <path d="M3 4h5v4H3zM8 8h5v4H8zM10 4h3" />,
    play: <path d="m5 3 7 5-7 5z" />,
    rotate: (
      <path d="M4 6a5 5 0 0 1 8 1M12 3v4H8M12 10a5 5 0 0 1-8-1M4 13V9h4" />
    ),
    comment: <path d="M2 3h12v8H7l-3 3v-3H2z" />,
    pages: <path d="M4 2h8v11H4zM2 5v9h8" />,
    zoom: (
      <path d="M6.5 3a4 4 0 1 0 0 8 4 4 0 0 0 0-8zM10 10l3 3M6.5 5v4M4.5 7h4" />
    ),
    share: <path d="M8 9V2m0 0L5 5m3-3 3 3M3 8v5h10V8" />,
    search: <path d="M6.5 3a3.5 3.5 0 1 0 0 7 3.5 3.5 0 0 0 0-7zM9 9.5l3 3" />,
    save: <path d="M3 2h8l2 2v10H3zM5 2v4h5V2M5 14v-5h6v5" />,
    bold: <path d="M5 3h4a2 2 0 0 1 0 4H5zm0 4h5a2.5 2.5 0 0 1 0 5H5z" />,
    italic: <path d="M9 3H6M10 13H7M8 3 6 13" />,
    underline: <path d="M4 3v4a4 4 0 0 0 8 0V3M3 14h10" />,
    table: <path d="M2 3h12v10H2zM2 7h12M2 10h12M6 3v10M10 3v10" />,
    image: <path d="M2 3h12v10H2zM4 10l2-2 2 2 2-3 2 3M5 6h.01" />,
    more: <path d="M3 8h.01M8 8h.01M13 8h.01" />,
  };
  return (
    <svg aria-hidden="true" viewBox="0 0 16 16">
      {paths[kind]}
    </svg>
  );
}

export function MotionRibbon({
  tabs,
  active,
  commands,
  accent = 'blue',
}: {
  tabs: string[];
  active: string;
  commands: MotionCommand[];
  accent?: 'blue' | 'green' | 'orange' | 'violet';
}) {
  return (
    <div
      className="office-motion-ribbon"
      data-accent={accent}
      aria-hidden="true"
    >
      <div className="office-motion-ribbon__tabs">
        {tabs.map((tab) => (
          <span className={tab === active ? 'is-active' : undefined} key={tab}>
            {tab}
          </span>
        ))}
        <i className="office-motion-ribbon__tab-spacer" />
        <span className="office-motion-ribbon__window-actions">
          <i />
          <i />
          <i />
        </span>
      </div>
      <div className="office-motion-ribbon__commands">
        {commands.map((command, index) => (
          <Fragment key={`${command.icon}-${command.label}`}>
            {index > 0 && <i className="office-motion-ribbon__separator" />}
            <span
              className={`office-motion-command${command.active ? ' is-active' : ''}`}
              data-command={command.icon}
            >
              <MotionCommandGlyph kind={command.icon} />
              <small>{command.label}</small>
            </span>
          </Fragment>
        ))}
      </div>
    </div>
  );
}

export function WindowChrome({
  file,
  status,
  children,
}: {
  file: string;
  status: string;
  children: ReactNode;
}) {
  return (
    <div className="office-motion-window">
      <header className="office-motion-window__bar">
        <span className="office-motion-window__dots" aria-hidden="true">
          <i />
          <i />
          <i />
        </span>
        <code>{file}</code>
        <span className="office-motion-window__status">
          <i aria-hidden="true" />
          {status}
        </span>
      </header>
      {children}
    </div>
  );
}
