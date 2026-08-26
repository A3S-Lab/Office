import type { ReactNode } from 'react';
import {
  Component,
  lazy,
  Suspense,
  useCallback,
  useEffect,
  useMemo,
  useState,
} from 'react';
import type {
  DocumentContent,
  PresentationContent,
  SpreadsheetContent,
} from '../../../src/core';

type HomeLanguage = 'zh' | 'en';
type DemoKind = 'document' | 'spreadsheet' | 'presentation';

interface HomeEditorDemoProps {
  language: HomeLanguage;
}

interface EditorErrorBoundaryProps {
  children: ReactNode;
  onRetry: () => void;
  retryLabel: string;
  title: string;
}

interface EditorErrorBoundaryState {
  error: Error | null;
}

const LazyDocumentEditor = lazy(async () => {
  const module = await import('../../../src/react');
  return { default: module.DocumentEditor };
});

const LazySpreadsheetEditor = lazy(async () => {
  const module = await import('../../../src/react');
  return { default: module.SpreadsheetEditor };
});

const LazyPresentationEditor = lazy(async () => {
  const module = await import('../../../src/react');
  return { default: module.PresentationEditor };
});

const stageOrder: readonly DemoKind[] = [
  'document',
  'spreadsheet',
  'presentation',
];

const stageCopy: Record<
  HomeLanguage,
  Record<DemoKind, { label: string; file: string; detail: string }>
> = {
  zh: {
    document: {
      label: '文档',
      file: 'launch-brief.docx',
      detail: '正文、审阅与多人状态',
    },
    spreadsheet: {
      label: '表格',
      file: 'launch-plan.xlsx',
      detail: '公式、表格与数据状态',
    },
    presentation: {
      label: '演示文稿',
      file: 'launch-story.pptx',
      detail: '场景、对象与入场动画',
    },
  },
  en: {
    document: {
      label: 'Document',
      file: 'launch-brief.docx',
      detail: 'Writing, review, and presence',
    },
    spreadsheet: {
      label: 'Spreadsheet',
      file: 'launch-plan.xlsx',
      detail: 'Formulas, tables, and data state',
    },
    presentation: {
      label: 'Presentation',
      file: 'launch-story.pptx',
      detail: 'Scenes, objects, and motion',
    },
  },
};

const staticDocumentContent: DocumentContent = {
  type: 'document',
  pageSize: 'a4',
  trackChanges: true,
  html: [
    '<h1>Launch brief</h1>',
    '<p><strong>Product team</strong> · Shared working document</p>',
    '<p>One source of truth for the launch plan, with review history and live presence in the same surface.</p>',
    '<h2>What changed this week</h2>',
    '<p><span data-comment-id="home-comment" data-document-comment="true">The review group aligned on the first release window.</span> <ins data-document-change="true" data-change-kind="insertion" data-change-id="home-change" data-change-author="Mina Chen" data-change-date="2026-08-20T09:30:00.000Z">The release checklist is now ready for sign-off.</ins></p>',
    '<h2>Next steps</h2>',
    '<ol><li><p>Confirm owners and acceptance criteria</p></li><li><p>Publish the integration guide</p></li><li><p>Open the collaboration room</p></li></ol>',
  ].join(''),
  comments: [
    {
      id: 'home-comment',
      author: 'Mina Chen',
      date: '2026-08-20T09:30:00.000Z',
      text: 'Can we keep this decision visible to the whole launch group?',
      resolved: false,
    },
  ],
};

const staticSpreadsheetContent: SpreadsheetContent = {
  type: 'spreadsheet',
  sheets: [
    {
      id: 'home-plan',
      name: 'Launch plan',
      status: 1,
      order: 0,
      row: 18,
      column: 8,
      data: [
        [
          cell('Workstream', 'header'),
          cell('Owner', 'header'),
          cell('W1', 'header'),
          cell('W2', 'header'),
          cell('W3', 'header'),
          cell('Progress', 'header'),
        ],
        [
          cell('Docs and examples'),
          cell('Lin'),
          cell('Done'),
          cell('Done'),
          cell('In review'),
          cell('83%'),
        ],
        [
          cell('Realtime room'),
          cell('Mina'),
          cell('Done'),
          cell('In progress'),
          cell('Next'),
          cell('58%'),
        ],
        [
          cell('Release checks'),
          cell('Noah'),
          cell('Done'),
          cell('Done'),
          cell('Done'),
          cell('100%'),
        ],
        [
          cell('Customer preview'),
          cell('Avery'),
          cell('In progress'),
          cell('Next'),
          cell('Next'),
          cell('32%'),
        ],
      ],
      config: {
        columnlen: { 0: 170, 1: 90, 2: 84, 3: 84, 4: 84, 5: 96 },
        rowlen: { 0: 30 },
      },
      luckysheet_select_save: [
        {
          row: [2, 2],
          column: [5, 5],
          row_focus: 2,
          column_focus: 5,
        },
      ],
    },
  ],
};

const staticPresentationContent: PresentationContent = {
  type: 'presentation',
  width: 13.333,
  height: 7.5,
  slides: [
    {
      id: 'home-slide',
      name: 'Launch story',
      background: '#f8fbff',
      elements: [
        {
          id: 'home-slide-kicker',
          type: 'text',
          x: 8,
          y: 11,
          width: 78,
          height: 8,
          text: 'A3S OFFICE · SHARED WORKSPACE',
          fontSize: 13,
          color: '#2864e8',
          fill: 'transparent',
          bold: true,
          align: 'left',
        },
        {
          id: 'home-slide-title',
          type: 'text',
          x: 8,
          y: 27,
          width: 78,
          height: 20,
          text: 'Move from draft to decision.',
          fontSize: 33,
          color: '#172033',
          fill: 'transparent',
          bold: true,
          align: 'left',
        },
        {
          id: 'home-slide-body',
          type: 'shape',
          shapeType: 'roundRect',
          x: 8,
          y: 58,
          width: 45,
          height: 18,
          text: 'One shared state · three surfaces',
          fontSize: 16,
          color: '#ffffff',
          fill: '#1264ff',
          bold: true,
          align: 'center',
          radius: 3,
        },
      ],
      animations: [
        {
          id: 'home-animation',
          elementId: 'home-slide-title',
          effect: 'fade',
          trigger: 'with-previous',
          durationMs: 420,
          delayMs: 0,
        },
      ],
    },
  ],
};

function cell(value: string, variant?: 'header') {
  return {
    v: value,
    m: value,
    ...(variant === 'header' ? { bl: 1, bg: '#eaf2ff', fc: '#174ea6' } : {}),
  };
}

class EditorErrorBoundary extends Component<
  EditorErrorBoundaryProps,
  EditorErrorBoundaryState
> {
  state: EditorErrorBoundaryState = { error: null };

  static getDerivedStateFromError(error: Error): EditorErrorBoundaryState {
    return { error };
  }

  render() {
    if (this.state.error) {
      return (
        <div className="office-editor-demo__error" role="alert">
          <strong>{this.props.title}</strong>
          <span>
            {this.state.error.message || 'The editor could not start.'}
          </span>
          <button
            type="button"
            onClick={() => {
              this.setState({ error: null });
              this.props.onRetry();
            }}
          >
            {this.props.retryLabel}
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

function EditorLoading({ language }: { language: HomeLanguage }) {
  return (
    <div
      className="office-editor-demo__loading"
      role="status"
      aria-live="polite"
    >
      <span className="office-editor-demo__loading-page" aria-hidden="true">
        <i />
        <i />
        <i />
      </span>
      <strong>
        {language === 'zh' ? '正在打开真实编辑器' : 'Opening the live editor'}
      </strong>
      <small>
        {language === 'zh'
          ? '编辑器资源按需加载，内容会保留在当前演示状态。'
          : 'Editor resources load on demand while this demo keeps its state.'}
      </small>
    </div>
  );
}

function StaticEditorFallback({ language }: { language: HomeLanguage }) {
  return (
    <div
      className="office-editor-demo__fallback"
      data-editor-runtime="fallback"
    >
      <div className="office-editor-demo__fallback-toolbar" aria-hidden="true">
        <span />
        <span />
        <span />
        <span />
        <span />
      </div>
      <article>
        <small>
          {language === 'zh'
            ? '交互式编辑器预览'
            : 'Interactive editor preview'}
        </small>
        <h3>
          {language === 'zh' ? '发布协作简报' : 'Launch collaboration brief'}
        </h3>
        <p>
          {language === 'zh'
            ? '浏览器完成加载后，这里会显示完整 Ribbon、页面、审阅和协作状态。'
            : 'The full Ribbon, page, review, and collaboration state appears after the browser bundle is ready.'}
        </p>
        <div className="office-editor-demo__fallback-lines" aria-hidden="true">
          <i />
          <i />
          <i />
          <i className="is-short" />
        </div>
      </article>
    </div>
  );
}

function renderEditor(
  kind: DemoKind,
  content: DocumentContent | SpreadsheetContent | PresentationContent,
  onChange: (
    next: DocumentContent | SpreadsheetContent | PresentationContent,
  ) => void,
) {
  if (kind === 'document') {
    return (
      <LazyDocumentEditor
        artifactId="home-document"
        autoFocus={false}
        content={content as DocumentContent}
        defaultRibbonCollapsed={false}
        onChange={onChange}
        preview={false}
        saveStatus="已同步 · 3 位协作者"
        theme="light"
      />
    );
  }
  if (kind === 'spreadsheet') {
    return (
      <LazySpreadsheetEditor
        autoFocus={false}
        content={content as SpreadsheetContent}
        onChange={onChange}
        preview={false}
        saveStatus="已同步 · 3 位协作者"
        theme="light"
      />
    );
  }
  return (
    <LazyPresentationEditor
      autoFocus={false}
      content={content as PresentationContent}
      onChange={onChange}
      preview={false}
      saveStatus="已同步 · 3 位协作者"
      theme="light"
    />
  );
}

export function HomeEditorDemo({ language }: HomeEditorDemoProps) {
  const zh = language === 'zh';
  const [mounted, setMounted] = useState(false);
  const [stageIndex, setStageIndex] = useState(0);
  const [paused, setPaused] = useState(false);
  const [reducedMotion, setReducedMotion] = useState(false);
  const [retryKey, setRetryKey] = useState(0);
  const [documentContent, setDocumentContent] = useState<DocumentContent>(
    staticDocumentContent,
  );
  const [spreadsheetContent, setSpreadsheetContent] =
    useState<SpreadsheetContent>(staticSpreadsheetContent);
  const [presentationContent, setPresentationContent] =
    useState<PresentationContent>(staticPresentationContent);
  const kind = stageOrder[stageIndex] ?? 'document';
  const copy = stageCopy[language][kind];

  useEffect(() => {
    setMounted(true);
    const media = window.matchMedia('(prefers-reduced-motion: reduce)');
    const onChange = () => setReducedMotion(media.matches);
    onChange();
    media.addEventListener?.('change', onChange);
    return () => media.removeEventListener?.('change', onChange);
  }, []);

  useEffect(() => {
    if (!mounted || paused || reducedMotion) return;
    const timer = window.setInterval(() => {
      setStageIndex((current) => (current + 1) % stageOrder.length);
      setRetryKey((current) => current + 1);
    }, 8_500);
    return () => window.clearInterval(timer);
  }, [mounted, paused, reducedMotion]);

  useEffect(() => {
    if (!mounted || reducedMotion) return;
    const nextKind = stageOrder[(stageIndex + 1) % stageOrder.length];
    void import('../../../src/react')
      .then(({ preloadOfficeEditor }) => preloadOfficeEditor(nextKind))
      .catch(() => undefined);
  }, [mounted, reducedMotion, stageIndex]);

  const content = useMemo(() => {
    if (kind === 'document') return documentContent;
    if (kind === 'spreadsheet') return spreadsheetContent;
    return presentationContent;
  }, [documentContent, kind, presentationContent, spreadsheetContent]);

  const onContentChange = useCallback(
    (next: DocumentContent | SpreadsheetContent | PresentationContent) => {
      if (kind === 'document') setDocumentContent(next as DocumentContent);
      if (kind === 'spreadsheet')
        setSpreadsheetContent(next as SpreadsheetContent);
      if (kind === 'presentation')
        setPresentationContent(next as PresentationContent);
    },
    [kind],
  );

  const setStage = (next: number) => {
    setStageIndex(next);
    setRetryKey((current) => current + 1);
  };

  return (
    <figure
      className="office-editor-demo"
      data-editor-demo-stage={kind}
      data-editor-runtime={mounted ? 'live' : 'fallback'}
      onFocus={() => setPaused(true)}
      onBlur={(event) => {
        if (!event.currentTarget.contains(event.relatedTarget))
          setPaused(false);
      }}
      onMouseEnter={() => setPaused(true)}
      onMouseLeave={() => setPaused(false)}
      aria-label={
        zh
          ? 'A3S Office 真实编辑器演示，可切换文档、表格和演示文稿'
          : 'A3S Office live editor demo with document, spreadsheet, and presentation surfaces'
      }
    >
      <div className="office-editor-demo__window">
        <header className="office-editor-demo__window-bar">
          <span className="office-editor-demo__window-dots" aria-hidden="true">
            <i />
            <i />
            <i />
          </span>
          <strong>{copy.file}</strong>
          <span className="office-editor-demo__sync-state">
            <i /> {zh ? '实时同步' : 'Live sync'}
          </span>
        </header>

        <div
          className="office-editor-demo__surface"
          id="office-editor-demo-surface"
          key={`${kind}-${retryKey}`}
          role="tabpanel"
          aria-labelledby={`office-editor-demo-tab-${kind}`}
        >
          {!mounted ? (
            <StaticEditorFallback language={language} />
          ) : (
            <EditorErrorBoundary
              key={`${kind}-${retryKey}`}
              onRetry={() => setRetryKey((current) => current + 1)}
              retryLabel={zh ? '重试编辑器' : 'Retry editor'}
              title={zh ? '编辑器加载失败' : 'Editor failed to load'}
            >
              <Suspense fallback={<EditorLoading language={language} />}>
                {renderEditor(kind, content, onContentChange)}
              </Suspense>
            </EditorErrorBoundary>
          )}
        </div>

        <footer className="office-editor-demo__footer">
          <div
            className="office-editor-demo__stage-tabs"
            role="tablist"
            aria-label={zh ? '演示编辑器类型' : 'Demo editor surfaces'}
          >
            {stageOrder.map((stage, index) => {
              const active = index === stageIndex;
              return (
                <button
                  key={stage}
                  id={`office-editor-demo-tab-${stage}`}
                  type="button"
                  role="tab"
                  aria-selected={active}
                  aria-controls="office-editor-demo-surface"
                  aria-label={stageCopy[language][stage].label}
                  className={active ? 'is-active' : undefined}
                  onClick={() => setStage(index)}
                >
                  <span>{stageCopy[language][stage].label}</span>
                  <small>{stageCopy[language][stage].detail}</small>
                </button>
              );
            })}
          </div>
          <span className="office-editor-demo__footer-note">
            <b>{zh ? '真实组件' : 'Live component'}</b>
            {zh
              ? ' · 内容与状态由 A3S Office 驱动'
              : ' · Powered by A3S Office state'}
          </span>
        </footer>
      </div>
      <figcaption className="office-editor-demo__caption">
        <span>
          <i /> {copy.label}
        </span>
        <strong>{copy.detail}</strong>
        <small>
          {paused
            ? zh
              ? '已暂停，点击底部标签继续'
              : 'Paused · use the tabs below to continue'
            : zh
              ? '自动切换 · 悬停或聚焦可暂停'
              : 'Auto-rotating · hover or focus to pause'}
        </small>
      </figcaption>
    </figure>
  );
}
