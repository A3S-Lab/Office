import { useLang, withBase } from '@rspress/core/runtime';
import { useEffect } from 'react';
import { HomeEditorDemo } from './HomeEditorDemo';

type Language = 'zh' | 'en';

type Surface = {
  name: string;
  detail: string;
  component: string;
};

type WorkflowStage = {
  index: string;
  name: string;
  detail: string;
  links: Array<{ label: string; href: string }>;
  accent?: boolean;
};

const surfaces: Record<Language, Surface[]> = {
  zh: [
    {
      name: '文档',
      detail: 'TipTap / ProseMirror 与 A3S Office 排版内核',
      component: 'DocumentEditor',
    },
    {
      name: 'Markdown',
      detail: '可视化编辑与受控 Markdown 源码',
      component: 'MarkdownEditor',
    },
    {
      name: '表格',
      detail: '稀疏工作簿、原生表格与 A3S Office 计算内核',
      component: 'SpreadsheetEditor',
    },
    {
      name: '演示文稿',
      detail: '场景图与按需加载的文本编辑',
      component: 'PresentationEditor',
    },
    {
      name: 'PDF',
      detail: 'PDFium WebAssembly 渲染与批注控制器',
      component: 'PdfViewer',
    },
  ],
  en: [
    {
      name: 'Document',
      detail: 'TipTap and ProseMirror with the A3S Office layout kernel',
      component: 'DocumentEditor',
    },
    {
      name: 'Markdown',
      detail: 'Visual editing with a controlled Markdown source model',
      component: 'MarkdownEditor',
    },
    {
      name: 'Spreadsheet',
      detail:
        'Sparse workbooks, native tables, and the A3S Office calculation kernel',
      component: 'SpreadsheetEditor',
    },
    {
      name: 'Presentation',
      detail: 'A scene graph with on-demand text editing',
      component: 'PresentationEditor',
    },
    {
      name: 'PDF',
      detail: 'PDFium WebAssembly rendering and annotation controllers',
      component: 'PdfViewer',
    },
  ],
};

const latestCapabilities: Record<
  Language,
  Array<{ label: string; type: string }>
> = {
  zh: [
    { label: '入场动画', type: 'Presentation' },
    { label: '文档比较', type: 'Writer' },
    { label: '原生索引', type: 'Writer' },
    { label: '可更新目录', type: 'Writer' },
    { label: '字符底纹', type: 'Writer' },
    { label: '校对语言', type: 'Writer' },
    { label: '数据验证', type: 'Spreadsheet' },
    { label: '页面组织', type: 'PDF' },
  ],
  en: [
    { label: 'Entrance animations', type: 'Presentation' },
    { label: 'Document compare', type: 'Writer' },
    { label: 'Document index', type: 'Writer' },
    { label: 'Table of contents', type: 'Writer' },
    { label: 'Character shading', type: 'Writer' },
    { label: 'Proofing languages', type: 'Writer' },
    { label: 'Data validation', type: 'Spreadsheet' },
    { label: 'Page organization', type: 'PDF' },
  ],
};

const capabilityHrefs: Record<Language, string[]> = {
  zh: [
    'components/presentation.html#入场动画',
    'components/document.html#文档比较与合并',
    'components/document.html#原生文档索引',
    'components/document.html#原生可更新目录',
    'components/document.html#原生字符底纹',
    'components/document.html#原生校对语言',
    'components/spreadsheet.html#数据验证',
    'components/pdf.html#页面组织',
  ],
  en: [
    'components/presentation.html#entrance-animations',
    'components/document.html#document-compare-and-combine',
    'components/document.html#native-document-index',
    'components/document.html#native-table-of-contents',
    'components/document.html#native-character-shading',
    'components/document.html#native-proofing-languages',
    'components/spreadsheet.html#data-validation',
    'components/pdf.html#page-organization',
  ],
};

function docsPath(path: string, language: Language) {
  const locale = language === 'en' ? 'en/' : '';
  return withBase(`/docs/${locale}${path}`);
}

function surfaceDocsPage(component: string): string {
  const pages: Record<string, string> = {
    DocumentEditor: 'document',
    MarkdownEditor: 'markdown',
    SpreadsheetEditor: 'spreadsheet',
    PresentationEditor: 'presentation',
    PdfViewer: 'pdf',
  };
  return pages[component] ?? 'index';
}

function productCopy(language: Language) {
  const zh = language === 'zh';
  return {
    lead: zh
      ? '把文档、Markdown、表格、演示文稿和 PDF 嵌入你的产品。宿主管理文件与权限，团队共享可审计的实时状态。'
      : 'Embed documents, Markdown, spreadsheets, presentations, and PDFs. Your host keeps files and permissions while teams share live, auditable state.',
    primary: zh ? '查看接入文档' : 'Read the integration docs',
    secondary: zh ? '打开 Playground' : 'Open Playground',
    install: zh ? '安装 A3S Office' : 'Install A3S Office',
    assurance: zh ? '统一宿主契约' : 'One host contract',
    latest: zh ? 'main 最新能力' : 'Latest on main',
    workflowTitle: zh ? '按工作流接入' : 'Integrate by workflow',
    workflowLead: zh
      ? '从第一个组件开始，逐步接入文件能力、实时协作和智能体自动化。'
      : 'Start with one component, then add file behavior, live collaboration, and agent automation.',
    collaborationTitle: zh
      ? '协作是一条完整链路'
      : 'Collaboration is a complete path',
    collaborationLead: zh
      ? 'Office 负责格式绑定、成员状态、远端位置与本地撤销。宿主继续拥有房间、鉴权、权限、网络 Provider、离线队列和持久化。'
      : 'Office owns format bindings, participants, remote locations, and local undo. The host keeps rooms, authentication, authorization, networking, offline buffering, and persistence.',
    editorsTitle: zh ? '选择编辑器界面' : 'Choose an editor surface',
    editorsLead: zh
      ? '每种文件格式都有自己的数据模型，同时共享统一的宿主接入方式。'
      : 'Each file format keeps its own data model while sharing one host integration pattern.',
    finalTitle: zh
      ? '把编辑器带进你的产品。'
      : 'Bring the editor into your product.',
    finalLead: zh
      ? '从文档开始，再按需打开协作、自动化和原生能力。'
      : 'Start with a document, then open collaboration, automation, and native capabilities as you need them.',
    docs: zh ? '文档' : 'Docs',
    collaboration: zh ? '多人实时协作' : 'Live collaboration',
    backend: zh ? 'A3S Boot 后端' : 'A3S Boot backend',
    automation: zh ? 'CLI 与 Skill' : 'CLI and Skill',
  };
}

function ArrowIcon({ external = false }: { external?: boolean }) {
  return (
    <svg aria-hidden="true" className="docs-home-arrow" viewBox="0 0 16 16">
      {external ? (
        <>
          <path d="M6 3h7v7" />
          <path d="m13 3-8 8" />
          <path d="M11 9v4H3V5h4" />
        </>
      ) : (
        <>
          <path d="M3 8h10" />
          <path d="m9 4 4 4-4 4" />
        </>
      )}
    </svg>
  );
}

function HomeSurfaceMap({ language }: { language: Language }) {
  const zh = language === 'zh';
  return (
    <div
      className="office-home-surface-map"
      role="img"
      aria-label={
        zh
          ? 'A3S Office 统一宿主契约连接文档、Markdown、表格、演示文稿和 PDF 五种编辑器'
          : 'One A3S Office host contract connects document, Markdown, spreadsheet, presentation, and PDF editors'
      }
    >
      <div className="office-home-surface-map__window">
        <header className="office-home-surface-map__bar">
          <span className="office-home-surface-map__dots" aria-hidden="true">
            <i />
            <i />
            <i />
          </span>
          <code>A3S OFFICE / SURFACE MAP</code>
          <b>
            <i aria-hidden="true" />
            {zh ? '统一宿主契约' : 'One host contract'}
          </b>
        </header>
        <div className="office-home-surface-map__body">
          <div className="office-home-surface-map__intro">
            <small>
              {zh ? '一个宿主 · 五种界面' : 'ONE HOST · FIVE SURFACES'}
            </small>
            <strong>
              {zh ? '按文件模型选择体验' : 'Choose by file model'}
            </strong>
            <span>
              {zh
                ? '格式、权限和协作状态沿同一条接入路径流动。'
                : 'Format, permissions, and collaboration state travel through one integration path.'}
            </span>
          </div>
          <ol className="office-home-surface-map__list">
            {surfaces[language].map((surface, index) => (
              <li key={surface.component}>
                <span>{`0${index + 1}`}</span>
                <strong>{surface.name}</strong>
                <code>{surface.component}</code>
                <i aria-hidden="true" />
              </li>
            ))}
          </ol>
          <div className="office-home-surface-map__signal" aria-hidden="true">
            <i />
            <i />
            <i />
          </div>
        </div>
        <footer className="office-home-surface-map__footer">
          <span>
            {zh
              ? '按需加载 · 状态可控 · 文件原生'
              : 'On demand · controlled state · native files'}
          </span>
          <b>READY</b>
        </footer>
      </div>
    </div>
  );
}

export function HomeLayout() {
  const language = (useLang() === 'en' ? 'en' : 'zh') as Language;
  const copy = productCopy(language);
  const zh = language === 'zh';
  const playground = withBase('/playground/');
  const languageHref = (target: Language) =>
    target === 'en' ? withBase('/en/') : withBase('/');
  const latest = latestCapabilities[language];
  const hrefs = capabilityHrefs[language];

  useEffect(() => {
    // Rspress preserves the previous scroll offset when a viewport changes or
    // a localized home route is hydrated. Product home must always begin at
    // its hero; in-page hash navigation remains opt-in.
    const previousRestoration = window.history.scrollRestoration;
    window.history.scrollRestoration = 'manual';
    const resetHomeScroll = () => {
      if (window.location.hash) return;

      // The shared A3S shell uses smooth scrolling for in-page links. Reusing
      // that behavior for arrival makes the async editor mount look like a
      // restored scroll position on compact screens. Temporarily override the
      // root behavior so the product home always paints from its hero.
      const root = document.documentElement;
      const previousBehavior = root.style.scrollBehavior;
      root.style.scrollBehavior = 'auto';
      window.scrollTo({ top: 0, left: 0, behavior: 'auto' });
      if (document.scrollingElement) document.scrollingElement.scrollTop = 0;
      document.documentElement.scrollTop = 0;
      document.body.scrollTop = 0;

      // Rspress may own the scroll position from a layout container instead
      // of the document element. Reset only scrollable ancestors of the
      // product surface so docs pages keep their own reading position.
      let ancestor = document.querySelector<HTMLElement>(
        '.office-product-home',
      )?.parentElement;
      while (ancestor && ancestor !== document.body) {
        const style = window.getComputedStyle(ancestor);
        if (
          (style.overflowY === 'auto' || style.overflowY === 'scroll') &&
          ancestor.scrollHeight > ancestor.clientHeight
        ) {
          ancestor.scrollTop = 0;
        }
        ancestor = ancestor.parentElement;
      }
      root.style.scrollBehavior = previousBehavior;
    };

    resetHomeScroll();
    const frame = window.requestAnimationFrame(resetHomeScroll);
    const secondFrame = window.requestAnimationFrame(() => {
      window.requestAnimationFrame(resetHomeScroll);
    });
    const deferred = window.setTimeout(resetHomeScroll, 180);
    window.addEventListener('pageshow', resetHomeScroll);
    window.addEventListener('popstate', resetHomeScroll);
    return () => {
      window.cancelAnimationFrame(frame);
      window.cancelAnimationFrame(secondFrame);
      window.clearTimeout(deferred);
      window.removeEventListener('pageshow', resetHomeScroll);
      window.removeEventListener('popstate', resetHomeScroll);
      window.history.scrollRestoration = previousRestoration;
    };
  }, []);

  const workflow: WorkflowStage[] = [
    {
      index: '01',
      name: zh ? '开始' : 'Start',
      detail: zh
        ? '安装包并在现有应用中挂载第一个编辑器。'
        : 'Install the package and mount the first editor in an existing application.',
      links: [
        { label: copy.docs, href: docsPath('guide/', language) },
        {
          label: zh ? '选择框架' : 'Choose a framework',
          href: docsPath('components/', language),
        },
      ],
    },
    {
      index: '02',
      name: zh ? '构建' : 'Build',
      detail: zh
        ? '配置文件模型、受控状态、导入导出和宿主扩展。'
        : 'Configure file models, controlled state, import and export, and host extensions.',
      links: [
        {
          label: zh ? '编辑器 API' : 'Editor APIs',
          href: docsPath('components/document.html', language),
        },
        {
          label: zh ? '扩展机制' : 'Extension model',
          href: docsPath('components/extensions.html', language),
        },
      ],
    },
    {
      index: '03',
      name: zh ? '协作' : 'Collaborate',
      detail: zh
        ? '让用户与编码智能体通过 Yjs/Yrs 实时编辑或审阅同一文件。'
        : 'Let people and coding agents edit or review the same file through Yjs and Yrs.',
      accent: true,
      links: [
        {
          label: copy.collaboration,
          href: docsPath('components/collaboration.html', language),
        },
        {
          label: copy.backend,
          href: docsPath('components/collaboration-server.html', language),
        },
      ],
    },
    {
      index: '04',
      name: zh ? '自动化' : 'Automate',
      detail: zh
        ? '通过 CLI、MCP、A3S Code 与 Skill 安全处理 Office 文件。'
        : 'Process Office files safely through the CLI, MCP, A3S Code, and the Skill.',
      links: [
        { label: copy.automation, href: docsPath('automation/', language) },
        {
          label: zh ? 'CLI 参考' : 'CLI reference',
          href: docsPath('cli-reference.html', language),
        },
      ],
    },
  ];

  return (
    <main className="docs-home office-product-home" data-home-surface="product">
      <section
        className="docs-home-hero"
        aria-labelledby="office-product-title"
      >
        <div className="docs-home-hero__copy">
          <div className="docs-home-hero__meta office-home-hero__meta">
            <span>
              A3S OFFICE <i aria-hidden="true">·</i> EMBEDDED EDITORS
            </span>
            <nav aria-label={zh ? '首页语言' : 'Homepage language'}>
              <a
                href={languageHref('zh')}
                aria-current={zh ? 'page' : undefined}
              >
                中文
              </a>
              <a
                href={languageHref('en')}
                aria-current={!zh ? 'page' : undefined}
              >
                EN
              </a>
            </nav>
          </div>
          <h1 id="office-product-title">
            <span className="docs-home-hero__brand">A3S Office</span>
            <strong className="docs-home-hero__promise">
              {zh
                ? '嵌入编辑器，协作自然发生'
                : 'Embedded editors for live collaboration'}
            </strong>
          </h1>
          <p className="docs-home-hero__lead">{copy.lead}</p>
          <nav
            className="docs-home-actions"
            aria-label={zh ? '产品入口' : 'Product entry points'}
          >
            <a
              className="docs-home-action docs-home-action--primary"
              href={docsPath('guide/', language)}
            >
              {copy.primary} <ArrowIcon />
            </a>
            <a
              className="docs-home-action docs-home-action--secondary"
              href={playground}
            >
              {copy.secondary} <ArrowIcon external />
            </a>
          </nav>
          <section className="docs-home-install" aria-label={copy.install}>
            <div className="docs-home-install__header">
              <span>INSTALL</span>
              <strong>@a3s-lab/office</strong>
            </div>
            <pre>
              <code>bun add @a3s-lab/office</code>
            </pre>
          </section>
        </div>

        <div className="docs-home-hero__visual">
          <HomeSurfaceMap language={language} />
        </div>
      </section>

      <section
        className="docs-home-assurance"
        aria-label={zh ? '编辑器能力范围' : 'Editor surfaces'}
      >
        <strong>{copy.assurance}</strong>
        <ul>
          {surfaces[language].map((surface, index) => (
            <li key={surface.component}>
              <b>0{index + 1}</b>
              <span>
                {surface.name}
                <small>{surface.component}</small>
              </span>
            </li>
          ))}
        </ul>
      </section>

      <nav
        className="docs-home-hero__latest docs-home-latest"
        aria-label={copy.latest}
      >
        <span>{copy.latest}</span>
        {latest.map((capability, index) => (
          <a href={docsPath(hrefs[index], language)} key={capability.label}>
            <small>{capability.type}</small>
            {capability.label}
          </a>
        ))}
      </nav>

      <HomeEditorDemo language={language} />

      <section
        className="docs-home-section docs-home-journey"
        aria-labelledby="office-workflow-title"
      >
        <div className="docs-home-section__intro">
          <h2 id="office-workflow-title">{copy.workflowTitle}</h2>
          <p>{copy.workflowLead}</p>
        </div>
        <ol className="docs-home-flow">
          {workflow.map((stage) => (
            <li
              className={`docs-home-stage${stage.accent ? ' docs-home-stage--collaborate' : ''}`}
              key={stage.index}
            >
              <span className="docs-home-stage__index" aria-hidden="true">
                {stage.index}
              </span>
              <div className="docs-home-stage__body">
                <h3>{stage.name}</h3>
                <p>{stage.detail}</p>
              </div>
              <div className="docs-home-stage__links">
                {stage.links.map((link) => (
                  <a href={link.href} key={link.href}>
                    {link.label} <ArrowIcon />
                  </a>
                ))}
              </div>
            </li>
          ))}
        </ol>
      </section>

      <section
        className="docs-home-collaboration"
        aria-labelledby="office-collaboration-title"
      >
        <div className="docs-home-collaboration__copy">
          <h2 id="office-collaboration-title">{copy.collaborationTitle}</h2>
          <p>{copy.collaborationLead}</p>
          <ul className="docs-home-collaboration__facts">
            {(zh
              ? [
                  '五种编辑器使用同一套类型化协作边界',
                  'Document 评论与建议模式保留署名、线程和不可变决定',
                  'Spreadsheet 表格/ListObject 以校验后的 ID 索引记录收敛',
                  '浏览器 Yjs 与原生 Yrs 交换标准增量更新',
                  'A3S Boot 在持久化和广播前执行语义授权',
                ]
              : [
                  'All five editors share one typed collaboration boundary',
                  'Document comments and suggestions retain attribution and immutable decisions',
                  'Spreadsheet Tables/ListObjects converge as validated ID-keyed records',
                  'Browser Yjs and native Yrs exchange standard incremental updates',
                  'A3S Boot authorizes review semantics before persistence and broadcast',
                ]
            ).map((fact) => (
              <li key={fact}>{fact}</li>
            ))}
          </ul>
          <div className="docs-home-collaboration__links">
            <a href={docsPath('components/collaboration.html', language)}>
              {zh ? '阅读接入流程' : 'Read the integration flow'} <ArrowIcon />
            </a>
            <a
              href={docsPath('components/collaboration-server.html', language)}
            >
              {zh ? '部署完整后端' : 'Deploy the full backend'} <ArrowIcon />
            </a>
          </div>
        </div>
        <figure
          className="docs-home-collaboration__diagram"
          role="img"
          aria-label={
            zh
              ? '浏览器和原生智能体通过 A3S Boot 共享协作状态'
              : 'Browsers and a native agent share collaboration state through A3S Boot'
          }
        >
          <div className="docs-home-diagram__peers" aria-hidden="true">
            <span>
              <b>A</b>Browser A<small>Yjs</small>
            </span>
            <span>
              <b>B</b>Browser B<small>Yjs</small>
            </span>
            <span>
              <b>C</b>A3S Agent<small>Yrs</small>
            </span>
          </div>
          <div className="docs-home-diagram__connector" aria-hidden="true">
            <i />
            <b>state vectors + updates</b>
            <i />
          </div>
          <div className="docs-home-diagram__service" aria-hidden="true">
            <strong>A3S Boot</strong>
            <span>room auth · awareness · persistence</span>
          </div>
        </figure>
      </section>

      <section
        className="docs-home-section docs-home-editors"
        aria-labelledby="office-editors-title"
      >
        <div className="docs-home-section__intro">
          <h2 id="office-editors-title">{copy.editorsTitle}</h2>
          <p>{copy.editorsLead}</p>
        </div>
        <ul className="docs-home-surface-list">
          {surfaces[language].map((surface) => (
            <li key={surface.component}>
              <a
                href={docsPath(
                  `components/${surfaceDocsPage(surface.component)}.html`,
                  language,
                )}
              >
                <strong>{surface.name}</strong>
                <span>{surface.detail}</span>
                <code>{surface.component}</code>
              </a>
            </li>
          ))}
        </ul>
      </section>

      <section className="docs-home-final" aria-labelledby="office-final-title">
        <div>
          <h2 id="office-final-title">{copy.finalTitle}</h2>
          <p>{copy.finalLead}</p>
        </div>
        <div className="docs-home-final__actions">
          <a
            className="docs-home-action docs-home-action--primary"
            href={docsPath('', language)}
          >
            {zh ? '浏览文档' : 'Browse docs'} <ArrowIcon />
          </a>
          <a
            className="docs-home-action docs-home-action--secondary"
            href={playground}
          >
            {copy.secondary}
          </a>
        </div>
      </section>
    </main>
  );
}
