import { useLang, withBase } from '@rspress/core/runtime';
import { useEffect } from 'react';
import { HomeEditorDemo } from './HomeEditorDemo';
import { editorScreenshotFiles } from './home-editor-assets';
import { chapterOrder } from './HomeEditorContent';
import type { ChapterKind } from './HomeEditorContent';

type Language = 'zh' | 'en';

type Surface = {
  name: string;
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
      component: 'DocumentEditor',
    },
    {
      name: 'Markdown',
      component: 'MarkdownEditor',
    },
    {
      name: '表格',
      component: 'SpreadsheetEditor',
    },
    {
      name: '演示文稿',
      component: 'PresentationEditor',
    },
    {
      name: 'PDF',
      component: 'PdfViewer',
    },
  ],
  en: [
    {
      name: 'Document',
      component: 'DocumentEditor',
    },
    {
      name: 'Markdown',
      component: 'MarkdownEditor',
    },
    {
      name: 'Spreadsheet',
      component: 'SpreadsheetEditor',
    },
    {
      name: 'Presentation',
      component: 'PresentationEditor',
    },
    {
      name: 'PDF',
      component: 'PdfViewer',
    },
  ],
};

const latestCapabilities: Record<Language, Array<{ label: string }>> = {
  zh: [
    { label: '入场动画' },
    { label: '文档比较' },
    { label: '原生索引' },
    { label: '可更新目录' },
    { label: '字符底纹' },
    { label: '校对语言' },
    { label: '数据验证' },
    { label: '页面组织' },
  ],
  en: [
    { label: 'Entrance animations' },
    { label: 'Document compare' },
    { label: 'Document index' },
    { label: 'Table of contents' },
    { label: 'Character shading' },
    { label: 'Proofing languages' },
    { label: 'Data validation' },
    { label: 'Page organization' },
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
      ? '在你的应用里打开、编辑和保存文档、Markdown、表格、演示文稿与 PDF。文件、权限和存储仍由你的应用管理；多人协作可以单独接入。'
      : 'Open, edit, and save documents, Markdown, spreadsheets, presentations, and PDFs in your app. Your app keeps files, permissions, and storage; collaboration can be added separately.',
    primary: zh ? '开始接入' : 'Start integrating',
    secondary: zh ? '打开 Playground' : 'Open Playground',
    assurance: zh ? '支持的编辑器' : 'Supported editors',
    latest: zh ? '最近更新' : 'Recent updates',
    workflowTitle: zh ? '接入步骤' : 'Integration steps',
    workflowLead: zh
      ? '先渲染一个编辑器，再接上文件数据。需要多人协作时，继续接入协作服务。'
      : 'Render one editor, then connect your file data. Add collaboration when you need it.',
    collaborationTitle: zh
      ? '成员和 Agent 同时改一份文件'
      : 'Let people and an agent edit one file',
    collaborationLead: zh
      ? '成员 A、成员 B 和 A3S Agent 打开同一份项目方案。文字、评论和审阅状态会同步到三块屏幕。'
      : 'Person A, Person B, and A3S Agent open the same project brief. Text, comments, and review state appear on all three screens.',
    editorsTitle: zh ? '五种编辑器' : 'Five editors',
    editorsLead: zh
      ? '文档、Markdown、表格、演示文稿和 PDF，各有对应的编辑器。'
      : 'Documents, Markdown, spreadsheets, presentations, and PDFs each have a matching editor.',
    finalTitle: zh ? '开始接入' : 'Start integrating',
    finalLead: zh
      ? '先从一个编辑器开始，其他编辑器可以随后加入。'
      : 'Start with one editor and add the others when you need them.',
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
      data-real-editor-gallery="hero"
      role="img"
      aria-label={
        zh
          ? '五种真实 A3S Office 编辑器截图的堆叠预览'
          : 'Stacked previews of five real A3S Office editor screenshots'
      }
    >
      <div
        className="office-home-screenshot-stack"
        data-stack-style="poker-hand"
        aria-hidden="true"
      >
        {chapterOrder.map((kind: ChapterKind, index) => (
          <figure
            className={`office-home-screenshot-stack__card office-home-screenshot-stack__card--${kind}`}
            key={kind}
          >
            <img
              src={withBase(`/editor-previews/${editorScreenshotFiles[kind]}`)}
              alt=""
              loading={index === 0 ? 'eager' : 'lazy'}
              decoding="async"
            />
          </figure>
        ))}
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
  const collaborationParticipants = [
    {
      key: 'human-a',
      avatar: 'A',
      name: zh ? '成员 A' : 'Person A',
      role: zh ? '编辑' : 'Editor',
    },
    {
      key: 'human-b',
      avatar: 'B',
      name: zh ? '成员 B' : 'Person B',
      role: zh ? '审阅' : 'Reviewer',
    },
    {
      key: 'agent',
      avatar: 'AI',
      name: 'A3S Agent',
      role: zh ? 'Agent' : 'Agent',
    },
  ] as const;

  useEffect(() => {
    // Rspress preserves the previous scroll offset when a viewport changes or
    // a localized home route is hydrated. Product home must always begin at
    // its hero; in-page hash navigation remains opt-in.
    const previousRestoration = window.history.scrollRestoration;
    window.history.scrollRestoration = 'manual';
    let userHasScrolled = false;
    const markUserScroll = () => {
      userHasScrolled = true;
    };
    const resetHomeScroll = () => {
      if (window.location.hash || userHasScrolled) return;

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
    window.addEventListener('scroll', markUserScroll, {
      passive: true,
      once: true,
    });
    return () => {
      window.cancelAnimationFrame(frame);
      window.removeEventListener('scroll', markUserScroll);
      window.history.scrollRestoration = previousRestoration;
    };
  }, []);

  const workflow: WorkflowStage[] = [
    {
      index: '01',
      name: zh ? '安装' : 'Install',
      detail: zh
        ? '安装包，在现有页面里渲染第一个编辑器。'
        : 'Install the package and render the first editor in an existing page.',
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
      name: zh ? '接入' : 'Connect',
      detail: zh
        ? '接上文件数据，处理更新、导入导出和自定义按钮。'
        : 'Connect file data, updates, import and export, and custom commands.',
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
        ? '把同一份文件交给成员和 Agent 一起编辑。Yjs/Yrs 传递变更。'
        : 'Let people and an agent edit the same file. Yjs and Yrs carry the changes.',
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
        ? '用 CLI、MCP、A3S Code 或 Skill 批量处理 Office 文件。'
        : 'Process Office files in batches with the CLI, MCP, A3S Code, or a Skill.',
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
              {zh ? (
                <>
                  在应用里编辑
                  <br />
                  Office 文件
                </>
              ) : (
                <>
                  Edit Office files
                  <br />
                  in your app
                </>
              )}
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
          {surfaces[language].map((surface) => (
            <li key={surface.component} data-surface={surface.component}>
              <span>{surface.name}</span>
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
                  '评论、建议和修订都写回同一份文件',
                  '成员和 A3S Agent 能看到彼此的光标',
                  'A3S Boot 负责房间鉴权、消息转发和存储',
                ]
              : [
                  'Comments, suggestions, and revisions write to one file',
                  'People and the A3S Agent can see each other’s cursors',
                  'A3S Boot handles room auth, message routing, and storage',
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
              ? '两位成员和 A3S Agent 在同一编辑器中实时协作'
              : 'Two people and an A3S Agent collaborate in the same editor in real time'
          }
        >
          <div
            className="office-collab-demo"
            data-collaboration-animation="true"
            data-collaboration-participants="3"
            aria-hidden="true"
          >
            <div className="office-collab-demo__header">
              <strong>{zh ? '项目方案' : 'Project brief'}</strong>
              <span className="office-collab-demo__live">
                <i />
                {zh ? '同步中' : 'Syncing'}
              </span>
            </div>
            <div className="office-collab-demo__participants">
              {collaborationParticipants.map((participant) => (
                <article
                  className={`office-collab-peer office-collab-peer--${participant.key}`}
                  key={participant.key}
                  aria-label={`${participant.name}, ${participant.role}`}
                  data-peer-role={participant.role}
                >
                  <header className="office-collab-peer__header">
                    <b>{participant.avatar}</b>
                    <strong>{participant.name}</strong>
                  </header>
                  <div className="office-collab-peer__screen">
                    <img
                      src={withBase(
                        `/editor-previews/${editorScreenshotFiles.document}`,
                      )}
                      alt=""
                      loading="lazy"
                      decoding="async"
                    />
                    <span className="office-collab-peer__selection" />
                    <span className="office-collab-peer__cursor">
                      <i />
                      <b>{participant.avatar}</b>
                    </span>
                  </div>
                </article>
              ))}
            </div>
            <div className="office-collab-demo__wire">
              <span className="office-collab-demo__wire-line">
                <i />
              </span>
              <strong>{zh ? '编辑变更' : 'Edit changes'}</strong>
              <span className="office-collab-demo__wire-line office-collab-demo__wire-line--reverse">
                <i />
              </span>
            </div>
            <div className="office-collab-demo__service">
              <div>
                <strong>A3S Boot</strong>
                <span>{zh ? '同一房间' : 'One room'}</span>
              </div>
              <span className="office-collab-demo__service-state">
                <i />
                {zh ? '已连接' : 'Connected'}
              </span>
            </div>
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
                <span
                  className="docs-home-surface-list__arrow"
                  aria-hidden="true"
                >
                  <ArrowIcon />
                </span>
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
