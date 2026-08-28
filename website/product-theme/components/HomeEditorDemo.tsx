import { withBase } from '@rspress/core/runtime';
import { useEffect, useRef, useState } from 'react';
import {
  chapterOrder,
  MotionArrow,
  MotionRibbon,
  SyncIcon,
  WindowChrome,
} from './HomeEditorContent';
import type { ChapterKind, HomeLanguage } from './HomeEditorContent';

interface ChapterCopy {
  index: string;
  label: string;
  component: string;
  title: string;
  detail: string;
  features: string[];
  linkLabel: string;
}

const chapterCopy: Record<HomeLanguage, Record<ChapterKind, ChapterCopy>> = {
  zh: {
    document: {
      index: '01',
      label: '文档',
      component: 'DocumentEditor',
      title: '写作、审阅和协作在同一页完成。',
      detail:
        '保留熟悉的功能区和页面感，把批注、建议、远端光标与撤销历史放进同一条编辑路径。',
      features: ['功能区与页面布局', '批注和修订', '多人实时状态'],
      linkLabel: '查看 DocumentEditor',
    },
    markdown: {
      index: '02',
      label: 'Markdown',
      component: 'MarkdownEditor',
      title: '源码和可视化预览保持同步。',
      detail:
        '作者可以在 Markdown 源码、分屏和可视化模式之间切换，光标位置和受控内容始终属于同一个模型。',
      features: ['源码 / 分屏 / 预览', 'CommonMark 结构', '受控 Markdown 状态'],
      linkLabel: '查看 MarkdownEditor',
    },
    spreadsheet: {
      index: '03',
      label: '表格',
      component: 'SpreadsheetEditor',
      title: '选择一个单元格，就能读懂整张表。',
      detail:
        '公式栏、稀疏网格、表格样式和数据验证围绕当前选择展开；大量空白区域不会变成大量 DOM。',
      features: ['公式栏和范围选择', '稀疏渲染', '数据验证与表格'],
      linkLabel: '查看 SpreadsheetEditor',
    },
    presentation: {
      index: '04',
      label: '演示文稿',
      component: 'PresentationEditor',
      title: '从场景图到入场动画，顺着对象工作。',
      detail:
        '幻灯片、对象、动画时间线和演讲者备注各自有清晰的焦点，按需加载文本编辑，不牺牲画布流畅度。',
      features: ['场景图对象', '入场动画时间线', '演讲者备注'],
      linkLabel: '查看 PresentationEditor',
    },
    pdf: {
      index: '05',
      label: 'PDF',
      component: 'PdfViewer',
      title: '页面组织和批注，不打断阅读。',
      detail:
        'PDFium WebAssembly 负责渲染，缩略图、批注和页面操作保持在同一工作台，并把重写任务交给 Worker。',
      features: ['PDFium WebAssembly', '缩略图与批注', '页面组织 Worker'],
      linkLabel: '查看 PdfViewer',
    },
  },
  en: {
    document: {
      index: '01',
      label: 'Document',
      component: 'DocumentEditor',
      title: 'Write, review, and collaborate on one page.',
      detail:
        'Keep the familiar ribbon and page model while comments, suggestions, remote cursors, and undo history share one editing path.',
      features: [
        'Ribbon and page layout',
        'Comments and revisions',
        'Live participant state',
      ],
      linkLabel: 'View DocumentEditor',
    },
    markdown: {
      index: '02',
      label: 'Markdown',
      component: 'MarkdownEditor',
      title: 'Source and visual preview stay in sync.',
      detail:
        'Authors can move between Markdown source, split view, and visual mode while caret position and controlled content remain one model.',
      features: [
        'Source / split / preview',
        'CommonMark structure',
        'Controlled Markdown state',
      ],
      linkLabel: 'View MarkdownEditor',
    },
    spreadsheet: {
      index: '03',
      label: 'Spreadsheet',
      component: 'SpreadsheetEditor',
      title: 'Select one cell and understand the whole sheet.',
      detail:
        'The formula bar, sparse grid, table styles, and data validation follow the active selection; empty maximum-size ranges stay unmaterialized.',
      features: [
        'Formula bar and ranges',
        'Sparse rendering',
        'Validation and tables',
      ],
      linkLabel: 'View SpreadsheetEditor',
    },
    presentation: {
      index: '04',
      label: 'Presentation',
      component: 'PresentationEditor',
      title: 'Work from the scene graph to the entrance cue.',
      detail:
        'Slides, objects, animation timing, and presenter notes each get a clear focus, with on-demand text editing that keeps the canvas responsive.',
      features: [
        'Scene-graph objects',
        'Entrance animation timing',
        'Presenter notes',
      ],
      linkLabel: 'View PresentationEditor',
    },
    pdf: {
      index: '05',
      label: 'PDF',
      component: 'PdfViewer',
      title: 'Organize pages and annotate without leaving reading mode.',
      detail:
        'PDFium WebAssembly renders the source while thumbnails, annotations, and page operations share one workbench; rewrites run in a Worker.',
      features: [
        'PDFium WebAssembly',
        'Thumbnails and annotations',
        'Worker page operations',
      ],
      linkLabel: 'View PdfViewer',
    },
  },
};

function DocumentPreview({ language }: { language: HomeLanguage }) {
  const zh = language === 'zh';
  return (
    <WindowChrome
      file="launch-brief.docx"
      status={zh ? '示意 · 3 位在线' : 'Sample · 3 online'}
    >
      <MotionRibbon
        tabs={
          zh
            ? ['开始', '插入', '页面布局', '引用', '审阅', '视图']
            : ['Home', 'Insert', 'Layout', 'References', 'Review', 'View']
        }
        active={zh ? '开始' : 'Home'}
        commands={
          zh
            ? [
                { icon: 'undo', label: '撤销' },
                { icon: 'redo', label: '重做' },
                { icon: 'paste', label: '粘贴' },
                { icon: 'paint', label: '格式刷' },
                { icon: 'font', label: '字体' },
                { icon: 'align', label: '段落' },
                { icon: 'comment', label: '批注', active: true },
                { icon: 'share', label: '共享' },
              ]
            : [
                { icon: 'undo', label: 'Undo' },
                { icon: 'redo', label: 'Redo' },
                { icon: 'paste', label: 'Paste' },
                { icon: 'paint', label: 'Painter' },
                { icon: 'font', label: 'Font' },
                { icon: 'align', label: 'Paragraph' },
                { icon: 'comment', label: 'Comment', active: true },
                { icon: 'share', label: 'Share' },
              ]
        }
        accent="blue"
      />
      <div className="office-document-preview__ruler" aria-hidden="true">
        <span>0</span>
        <i />
        <span>1</span>
        <i />
        <span>2</span>
        <i />
        <span>3</span>
        <i />
        <span>4</span>
        <i />
        <span>5</span>
      </div>
      <div className="office-document-preview__body">
        <article className="office-document-preview__page" aria-hidden="true">
          <small>PRODUCT TEAM · SHARED BRIEF</small>
          <h4>{zh ? '发布简报' : 'Launch brief'}</h4>
          <p className="office-document-preview__meta">
            {zh
              ? '一份可审阅的协作工作文档'
              : 'A reviewed, shared working document'}
          </p>
          <div className="office-document-preview__rule" />
          <p className="office-document-preview__line is-wide" />
          <p className="office-document-preview__line" />
          <p className="office-document-preview__line is-medium" />
          <p className="office-document-preview__highlight">
            {zh
              ? '评审小组已确认首个发布窗口。'
              : 'The review group aligned on the release window.'}
            <span className="office-document-preview__caret">Mina</span>
          </p>
          <p className="office-document-preview__line is-short" />
          <p className="office-document-preview__line" />
          <p className="office-document-preview__line is-medium" />
        </article>
        <aside className="office-document-preview__comment" aria-hidden="true">
          <span className="office-document-preview__comment-pin">M</span>
          <strong>{zh ? 'Mina 的批注' : "Mina's comment"}</strong>
          <p>
            {zh
              ? '可以把这个决定保留给整个发布小组吗？'
              : 'Can we keep this decision visible to the launch group?'}
          </p>
          <small>{zh ? '刚刚 · 未解决' : 'Just now · Open'}</small>
        </aside>
      </div>
      <footer className="office-motion-window__footer">
        <span>{zh ? '修订 2 · 评论 1' : '2 revisions · 1 comment'}</span>
        <b>{zh ? '已同步' : 'Synced'}</b>
      </footer>
    </WindowChrome>
  );
}

function MarkdownPreview({ language }: { language: HomeLanguage }) {
  const zh = language === 'zh';
  return (
    <WindowChrome
      file="README.md"
      status={zh ? '示意 · 受控内容' : 'Sample · Controlled'}
    >
      <MotionRibbon
        tabs={
          zh
            ? ['编辑', '源码', '预览', '协作', '设置']
            : ['Edit', 'Source', 'Preview', 'Collab', 'Settings']
        }
        active={zh ? '编辑' : 'Edit'}
        commands={
          zh
            ? [
                { icon: 'undo', label: '撤销' },
                { icon: 'redo', label: '重做' },
                { icon: 'source', label: '源码', active: true },
                { icon: 'preview', label: '预览' },
                { icon: 'link', label: '链接' },
                { icon: 'search', label: '查找' },
              ]
            : [
                { icon: 'undo', label: 'Undo' },
                { icon: 'redo', label: 'Redo' },
                { icon: 'source', label: 'Source', active: true },
                { icon: 'preview', label: 'Preview' },
                { icon: 'link', label: 'Link' },
                { icon: 'search', label: 'Find' },
              ]
        }
        accent="violet"
      />
      <div className="office-markdown-preview__body">
        <section className="office-markdown-preview__source" aria-hidden="true">
          <header>
            <span>MARKDOWN</span>
            <code>source</code>
          </header>
          <div className="office-markdown-preview__source-editor">
            <ol aria-hidden="true">
              <li>1</li>
              <li>2</li>
              <li>3</li>
              <li>4</li>
              <li>5</li>
              <li>6</li>
              <li>7</li>
              <li>8</li>
            </ol>
            <pre>
              <code>
                <span className="is-muted">#</span> Launch plan{`\n`}
                <span className="is-muted">-</span> Owner: Mina{`\n`}
                <span className="is-muted">-</span> Status: <b>ready</b>
                {`\n\n`}
                <span className="is-muted">##</span> Next steps{`\n`}
                <span className="is-muted">1.</span> Publish docs{`\n`}
                <span className="is-muted">2.</span> Open the room
              </code>
            </pre>
          </div>
          <i className="office-markdown-preview__source-caret" />
        </section>
        <div className="office-markdown-preview__bridge" aria-hidden="true">
          <i />
          <span>
            <SyncIcon />
          </span>
          <i />
        </div>
        <article
          className="office-markdown-preview__rendered"
          aria-hidden="true"
        >
          <small>VISUAL PREVIEW</small>
          <h4>{zh ? '发布计划' : 'Launch plan'}</h4>
          <p>
            {zh
              ? '一份可以直接审阅的 Markdown 文档。'
              : 'A Markdown document ready for review.'}
          </p>
          <div className="office-markdown-preview__checklist">
            <span>
              <i />
              {zh ? '发布文档' : 'Publish docs'}
            </span>
            <span>
              <i />
              {zh ? '打开协作房间' : 'Open the room'}
            </span>
          </div>
          <span className="office-markdown-preview__selection">
            {zh ? '光标同步' : 'Caret synced'}
          </span>
        </article>
      </div>
      <footer className="office-motion-window__footer">
        <span>
          {zh ? '源码和预览共享同一状态' : 'Source and preview share one state'}
        </span>
        <b>{zh ? '已更新' : 'Updated'}</b>
      </footer>
    </WindowChrome>
  );
}

function SpreadsheetPreview({ language }: { language: HomeLanguage }) {
  const zh = language === 'zh';
  const rows = [
    ['Docs', 'Lin', 'Done', '83%'],
    ['Realtime room', 'Mina', 'In progress', '58%'],
    ['Release checks', 'Noah', 'Done', '100%'],
    ['Customer preview', 'Avery', 'Next', '32%'],
  ];
  return (
    <WindowChrome
      file="launch-plan.xlsx"
      status={zh ? '示意 · 自动保存' : 'Sample · Autosaved'}
    >
      <MotionRibbon
        tabs={
          zh
            ? ['开始', '插入', '页面布局', '公式', '数据', '审阅', '视图']
            : ['Home', 'Insert', 'Layout', 'Formulas', 'Data', 'Review', 'View']
        }
        active={zh ? '开始' : 'Home'}
        commands={
          zh
            ? [
                { icon: 'undo', label: '撤销' },
                { icon: 'redo', label: '重做' },
                { icon: 'paste', label: '粘贴' },
                { icon: 'bold', label: '加粗' },
                { icon: 'align', label: '对齐' },
                { icon: 'table', label: '表格' },
                { icon: 'filter', label: '筛选', active: true },
                { icon: 'more', label: '更多' },
              ]
            : [
                { icon: 'undo', label: 'Undo' },
                { icon: 'redo', label: 'Redo' },
                { icon: 'paste', label: 'Paste' },
                { icon: 'bold', label: 'Bold' },
                { icon: 'align', label: 'Align' },
                { icon: 'table', label: 'Table' },
                { icon: 'filter', label: 'Filter', active: true },
                { icon: 'more', label: 'More' },
              ]
        }
        accent="green"
      />
      <div className="office-spreadsheet-preview__formula" aria-hidden="true">
        <span className="office-spreadsheet-preview__name-box">F3</span>
        <b>fx</b>
        <code>=SUM(C3:C6)</code>
      </div>
      <div className="office-spreadsheet-preview__body">
        <span
          className="office-spreadsheet-preview__freeze"
          aria-hidden="true"
        />
        <table aria-hidden="true">
          <thead>
            <tr>
              <th />
              <th>A</th>
              <th>B</th>
              <th>C</th>
              <th>D</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, rowIndex) => (
              <tr key={row[0]}>
                <th>{rowIndex + 3}</th>
                <td>{row[0]}</td>
                <td>{row[1]}</td>
                <td className={rowIndex === 1 ? 'is-selected' : undefined}>
                  {row[2]}
                </td>
                <td className={rowIndex === 1 ? 'is-focus' : undefined}>
                  {row[3]}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <span
          className="office-spreadsheet-preview__validation"
          aria-hidden="true"
        >
          {zh ? '数据验证 · 3 个选项' : 'Validation · 3 options'}
        </span>
        <i
          className="office-spreadsheet-preview__fill-handle"
          aria-hidden="true"
        />
      </div>
      <div className="office-spreadsheet-preview__sheetbar" aria-hidden="true">
        <span className="is-active">{zh ? 'Launch plan' : 'Launch plan'}</span>
        <span>{zh ? 'Owners' : 'Owners'}</span>
        <b>+</b>
      </div>
      <footer className="office-motion-window__footer">
        <span>
          {zh ? 'Launch plan · 4 行已填充' : 'Launch plan · 4 populated rows'}
        </span>
        <b>{zh ? '范围 F3' : 'Range F3'}</b>
      </footer>
    </WindowChrome>
  );
}

function PresentationPreview({ language }: { language: HomeLanguage }) {
  const zh = language === 'zh';
  return (
    <WindowChrome
      file="launch-story.pptx"
      status={zh ? '示意 · 动画就绪' : 'Sample · Animation ready'}
    >
      <MotionRibbon
        tabs={
          zh
            ? ['开始', '插入', '设计', '切换', '动画', '幻灯片放映', '审阅']
            : [
                'Home',
                'Insert',
                'Design',
                'Transitions',
                'Animations',
                'Slide Show',
                'Review',
              ]
        }
        active={zh ? '动画' : 'Animations'}
        commands={
          zh
            ? [
                { icon: 'undo', label: '撤销' },
                { icon: 'redo', label: '重做' },
                { icon: 'slide', label: '新建幻灯片' },
                { icon: 'arrange', label: '排列' },
                { icon: 'play', label: '预览', active: true },
                { icon: 'rotate', label: '时序' },
                { icon: 'comment', label: '批注' },
              ]
            : [
                { icon: 'undo', label: 'Undo' },
                { icon: 'redo', label: 'Redo' },
                { icon: 'slide', label: 'New slide' },
                { icon: 'arrange', label: 'Arrange' },
                { icon: 'play', label: 'Preview', active: true },
                { icon: 'rotate', label: 'Timing' },
                { icon: 'comment', label: 'Comment' },
              ]
        }
        accent="orange"
      />
      <div className="office-presentation-preview__body">
        <aside className="office-presentation-preview__rail" aria-hidden="true">
          <span className="is-active">
            <b>01</b>
            <i />
          </span>
          <span>
            <b>02</b>
            <i />
          </span>
          <span>
            <b>03</b>
            <i />
          </span>
          <span className="office-presentation-preview__add">+</span>
        </aside>
        <div className="office-presentation-preview__stage" aria-hidden="true">
          <div className="office-presentation-preview__stagebar">
            <span>{zh ? '幻灯片 1 · 编辑' : 'Slide 1 · Editing'}</span>
            <span>{zh ? '适应窗口' : 'Fit to window'}</span>
          </div>
          <article className="office-presentation-preview__slide">
            <small>A3S OFFICE · SHARED WORKSPACE</small>
            <h4>{zh ? '从草稿到决定' : 'Move from draft to decision.'}</h4>
            <p>
              {zh
                ? '一个共享状态，三种编辑界面。'
                : 'One shared state · three surfaces.'}
            </p>
            <div className="office-presentation-preview__object-box">
              <span className="office-presentation-preview__object">
                {zh ? '入场 · 淡入' : 'Entrance · Fade'}
              </span>
              <i />
              <i />
              <i />
              <i />
            </div>
          </article>
          <div className="office-presentation-preview__timeline">
            <span>{zh ? '单击时' : 'On click'}</span>
            <i className="is-active" />
            <span>{zh ? '与上一项同时' : 'With previous'}</span>
            <i />
            <span>{zh ? '持续 0.42s' : '0.42s'}</span>
          </div>
        </div>
      </div>
      <footer className="office-motion-window__footer">
        <span>
          {zh ? '幻灯片 1 / 3 · 演讲者备注' : 'Slide 1 / 3 · Presenter notes'}
        </span>
        <b>{zh ? '可播放' : 'Playable'}</b>
      </footer>
    </WindowChrome>
  );
}

function PdfPreview({ language }: { language: HomeLanguage }) {
  const zh = language === 'zh';
  return (
    <WindowChrome
      file="review-packet.pdf"
      status={zh ? '示意 · PDFium' : 'Sample · PDFium rendered'}
    >
      <MotionRibbon
        tabs={
          zh
            ? ['阅读', '批注', '页面', '表单', '视图']
            : ['Read', 'Annotate', 'Pages', 'Forms', 'View']
        }
        active={zh ? '阅读' : 'Read'}
        commands={
          zh
            ? [
                { icon: 'undo', label: '撤销' },
                { icon: 'redo', label: '重做' },
                { icon: 'comment', label: '批注', active: true },
                { icon: 'pages', label: '页面' },
                { icon: 'rotate', label: '旋转' },
                { icon: 'zoom', label: '缩放' },
                { icon: 'search', label: '查找' },
              ]
            : [
                { icon: 'undo', label: 'Undo' },
                { icon: 'redo', label: 'Redo' },
                { icon: 'comment', label: 'Annotate', active: true },
                { icon: 'pages', label: 'Pages' },
                { icon: 'rotate', label: 'Rotate' },
                { icon: 'zoom', label: 'Zoom' },
                { icon: 'search', label: 'Find' },
              ]
        }
        accent="violet"
      />
      <div className="office-pdf-preview__controls" aria-hidden="true">
        <span>1 / 12</span>
        <i />
        <span>{zh ? '单页' : 'Single page'}</span>
        <b>−</b>
        <strong>90%</strong>
        <b>＋</b>
      </div>
      <div className="office-pdf-preview__body">
        <aside className="office-pdf-preview__thumbs" aria-hidden="true">
          <span className="is-active">
            <b>1</b>
            <i />
          </span>
          <span>
            <b>2</b>
            <i />
          </span>
          <span className="is-moving">
            <b>3</b>
            <i />
          </span>
        </aside>
        <article className="office-pdf-preview__page" aria-hidden="true">
          <small>REVIEW PACKET · PAGE 1</small>
          <h4>{zh ? '发布检查清单' : 'Release checklist'}</h4>
          <p>
            {zh
              ? '在阅读路径上完成页面组织和批注。'
              : 'Organize pages and annotate without leaving reading mode.'}
          </p>
          <div className="office-pdf-preview__highlight">
            {zh ? '需要负责人确认' : 'Owner confirmation needed'}
          </div>
          <div className="office-pdf-preview__line" />
          <div className="office-pdf-preview__line is-short" />
          <span className="office-pdf-preview__note">
            {zh ? 'Mina · 批注' : 'Mina · annotation'}
          </span>
        </article>
      </div>
      <footer className="office-motion-window__footer">
        <span>
          {zh
            ? '第 1 页 / 共 12 页 · 缩略图窗口化'
            : 'Page 1 / 12 · Windowed thumbnails'}
        </span>
        <b>{zh ? '可组织' : 'Organize'}</b>
      </footer>
    </WindowChrome>
  );
}

function renderPreview(kind: ChapterKind, language: HomeLanguage) {
  if (kind === 'document') return <DocumentPreview language={language} />;
  if (kind === 'markdown') return <MarkdownPreview language={language} />;
  if (kind === 'spreadsheet') return <SpreadsheetPreview language={language} />;
  if (kind === 'presentation')
    return <PresentationPreview language={language} />;
  return <PdfPreview language={language} />;
}

function docsHref(kind: ChapterKind, language: HomeLanguage) {
  const page: Record<ChapterKind, string> = {
    document: 'document',
    markdown: 'markdown',
    spreadsheet: 'spreadsheet',
    presentation: 'presentation',
    pdf: 'pdf',
  };
  const locale = language === 'en' ? 'en/' : '';
  return withBase(`/docs/${locale}components/${page[kind]}.html`);
}

export function HomeEditorDemo({ language }: { language: HomeLanguage }) {
  const zh = language === 'zh';
  const [paused, setPaused] = useState(false);
  const [reducedMotion, setReducedMotion] = useState(false);
  const [currentChapter, setCurrentChapter] = useState<ChapterKind>('document');
  const [activeChapters, setActiveChapters] = useState<
    Record<ChapterKind, boolean>
  >(
    () =>
      Object.fromEntries(chapterOrder.map((kind) => [kind, true])) as Record<
        ChapterKind,
        boolean
      >,
  );
  const [settledAnchor, setSettledAnchor] = useState<ChapterKind | null>(null);
  const chapterRefs = useRef(new Map<ChapterKind, HTMLElement>());
  const copy = chapterCopy[language];

  useEffect(() => {
    const media = window.matchMedia('(prefers-reduced-motion: reduce)');
    const onChange = () => setReducedMotion(media.matches);
    onChange();
    media.addEventListener?.('change', onChange);
    return () => media.removeEventListener?.('change', onChange);
  }, []);

  useEffect(() => {
    if (typeof IntersectionObserver === 'undefined') return;

    const nodes = Array.from(chapterRefs.current.values());
    if (nodes.length === 0) return;

    const observer = new IntersectionObserver(
      (entries) => {
        setActiveChapters((current) => {
          let changed = false;
          const next = { ...current };
          for (const entry of entries) {
            const kind = entry.target.getAttribute(
              'data-editor-chapter',
            ) as ChapterKind | null;
            if (!kind || next[kind] === entry.isIntersecting) continue;
            next[kind] = entry.isIntersecting;
            changed = true;
          }
          return changed ? next : current;
        });
      },
      { rootMargin: '180px 0px' },
    );

    nodes.forEach((node) => {
      observer.observe(node);
    });

    // Keep a compact, persistent reading position in sync with the chapter
    // that occupies the visual center. The existing chapter index remains the
    // only navigation model; this observer only supplies orientation state.
    const chapterVisibility = new Set<ChapterKind>();
    const currentChapterObserver = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          const kind = entry.target.getAttribute(
            'data-editor-chapter',
          ) as ChapterKind | null;
          if (!kind) continue;
          if (entry.isIntersecting) chapterVisibility.add(kind);
          else chapterVisibility.delete(kind);
        }

        const viewportCenter = window.innerHeight / 2;
        const closest = Array.from(chapterVisibility).sort((first, second) => {
          const firstRect = chapterRefs.current
            .get(first)
            ?.getBoundingClientRect();
          const secondRect = chapterRefs.current
            .get(second)
            ?.getBoundingClientRect();
          if (!firstRect || !secondRect) return 0;
          const firstCenter = firstRect.top + firstRect.height / 2;
          const secondCenter = secondRect.top + secondRect.height / 2;
          return (
            Math.abs(firstCenter - viewportCenter) -
            Math.abs(secondCenter - viewportCenter)
          );
        })[0];

        if (closest) {
          setCurrentChapter((current) =>
            current === closest ? current : closest,
          );
        }
      },
      {
        rootMargin: '-26% 0px -54% 0px',
        threshold: [0, 0.15, 0.5, 0.85],
      },
    );
    nodes.forEach((node) => {
      currentChapterObserver.observe(node);
    });
    // Use a viewport-only observer to start hash settling. The motion observer
    // above intentionally has a preload margin and is therefore unsuitable
    // for proving that a requested anchor is actually on screen. A fragment
    // navigation can be smooth and content-visibility can expand sections as
    // they approach the viewport, so one intersecting callback is not enough
    // to announce that the destination is settled.
    let probeTimer: number | undefined;
    let probeToken = 0;

    const cancelAnchorProbe = () => {
      if (probeTimer !== undefined) {
        window.clearTimeout(probeTimer);
        probeTimer = undefined;
      }
      probeToken += 1;
    };

    const requestedKindFromHash = () => {
      const requestedId = window.location.hash.slice(1);
      return chapterOrder.find(
        (kind) => `editor-chapter-${kind}` === requestedId,
      );
    };

    const scheduleAnchorProbe = (kind: ChapterKind) => {
      cancelAnchorProbe();
      const token = probeToken;
      let previousTop: number | null = null;
      let stableSamples = 0;
      const startedAt = performance.now();

      const probe = () => {
        if (token !== probeToken || requestedKindFromHash() !== kind) return;

        const node = chapterRefs.current.get(kind);
        const rect = node?.getBoundingClientRect();
        const inViewport =
          rect !== undefined &&
          rect.bottom > 0 &&
          rect.top < window.innerHeight;

        if (!node || !rect || !inViewport) {
          previousTop = null;
          stableSamples = 0;
        } else {
          const top = Math.round(rect.top * 10) / 10;
          stableSamples =
            previousTop !== null && Math.abs(top - previousTop) <= 0.5
              ? stableSamples + 1
              : 0;
          previousTop = top;
          if (stableSamples >= 2) {
            probeTimer = undefined;
            setSettledAnchor(kind);
            return;
          }
        }

        // Smooth scrolling normally settles in under a second. Keep a bounded
        // fallback for browsers that do not emit `scrollend`, while avoiding a
        // permanent timer when a malformed or stale hash cannot be resolved.
        if (performance.now() - startedAt < 5000) {
          probeTimer = window.setTimeout(probe, 80);
        } else {
          probeTimer = undefined;
        }
      };

      probe();
    };

    const anchorObserver = new IntersectionObserver(
      (entries) => {
        const requestedKind = requestedKindFromHash();
        if (!requestedKind) return;
        for (const entry of entries) {
          if (
            entry.target.getAttribute('data-editor-chapter') ===
              requestedKind &&
            entry.isIntersecting
          ) {
            scheduleAnchorProbe(requestedKind);
            break;
          }
        }
      },
      { rootMargin: '0px' },
    );
    nodes.forEach((node) => {
      anchorObserver.observe(node);
    });
    const onHashChange = () => {
      cancelAnchorProbe();
      const requestedKind = requestedKindFromHash();
      if (!requestedKind) {
        setSettledAnchor(null);
        return;
      }

      setCurrentChapter(requestedKind);
      setSettledAnchor(null);
      scheduleAnchorProbe(requestedKind);
    };
    const onScrollEnd = () => {
      const requestedKind = requestedKindFromHash();
      if (requestedKind) scheduleAnchorProbe(requestedKind);
    };
    window.addEventListener('hashchange', onHashChange);
    window.addEventListener('scrollend', onScrollEnd);
    onHashChange();
    return () => {
      cancelAnchorProbe();
      observer.disconnect();
      currentChapterObserver.disconnect();
      anchorObserver.disconnect();
      window.removeEventListener('hashchange', onHashChange);
      window.removeEventListener('scrollend', onScrollEnd);
    };
  }, []);

  const motionPaused = paused || reducedMotion;
  const motionToggleLabel = reducedMotion
    ? zh
      ? '动画已按系统设置关闭'
      : 'Animations disabled by system'
    : paused
      ? zh
        ? '播放预览'
        : 'Play previews'
      : zh
        ? '暂停动画'
        : 'Pause previews';
  const currentChapterCopy = copy[currentChapter];
  const playgroundHref = withBase('/playground/');
  const currentChapterNumber = chapterOrder.indexOf(currentChapter) + 1;

  return (
    <section
      id="editor-chapters"
      className="office-editor-chapters"
      data-editor-runtime="preview"
      data-preview-kind="editor-chapters"
      data-motion-paused={motionPaused ? 'true' : 'false'}
      data-current-chapter={currentChapter}
      data-anchor-settled={settledAnchor ?? undefined}
      aria-labelledby="office-editor-chapters-title"
      aria-describedby="office-editor-chapters-preview-note"
    >
      <header className="office-editor-chapters__header">
        <div>
          <span className="docs-home-section__eyebrow">
            01 / {zh ? '编辑器界面' : 'EDITOR SURFACES'}
          </span>
          <h2 id="office-editor-chapters-title">
            {zh
              ? '五种编辑器，五段交互语言。'
              : 'Five editors. Five interaction languages.'}
          </h2>
          <p>
            {zh
              ? '下面是每个可嵌入编辑器的 UI/UX 动画预览。它们展示真实组件的工作方式，完整编辑器和可编辑内容请在 Playground 中体验。'
              : 'These UI/UX motion previews show how each embeddable editor works. Open the Playground for the complete editors and editable content.'}
          </p>
        </div>
        <div className="office-editor-chapters__header-actions">
          <span
            className="office-editor-chapters__preview-note"
            id="office-editor-chapters-preview-note"
          >
            <i aria-hidden="true" />
            {zh ? '示意 UI/UX · 仅预览' : 'Illustrative UI/UX · preview only'}
          </span>
          <button
            className="office-editor-chapters__motion-toggle"
            type="button"
            aria-pressed={motionPaused}
            disabled={reducedMotion}
            onClick={() => setPaused((current) => !current)}
          >
            <span
              className="office-editor-chapters__motion-icon"
              aria-hidden="true"
            >
              {motionPaused ? (
                <span className="is-play" />
              ) : (
                <span className="is-pause">
                  <i />
                  <i />
                </span>
              )}
            </span>
            {motionToggleLabel}
          </button>
        </div>
        <span
          className="office-editor-chapters__anchor-status"
          aria-live="polite"
          aria-atomic="true"
        >
          {settledAnchor
            ? `${copy[settledAnchor].label} ${zh ? '章节已定位' : 'chapter located'}`
            : ''}
        </span>
      </header>

      <nav
        className="office-editor-chapters__index"
        aria-label={zh ? '编辑器章节' : 'Editor chapters'}
      >
        {chapterOrder.map((kind) => {
          const chapter = copy[kind];
          return (
            <a
              href={`#editor-chapter-${kind}`}
              key={kind}
              aria-current={currentChapter === kind ? 'location' : undefined}
            >
              <span>{chapter.index}</span>
              <strong>{chapter.label}</strong>
              <small>{chapter.component}</small>
            </a>
          );
        })}
      </nav>

      <div
        className="office-editor-chapters__progress"
        data-current-chapter={currentChapter}
        role="status"
        aria-live="polite"
        aria-atomic="true"
      >
        <div className="office-editor-chapters__progress-copy">
          <span>{zh ? '当前章节' : 'Now viewing'}</span>
          <strong>
            {currentChapterCopy.index} · {currentChapterCopy.label}
          </strong>
          <small>{currentChapterCopy.component}</small>
        </div>
        <div className="office-editor-chapters__progress-meter">
          <span>
            {currentChapterNumber} /{' '}
            {String(chapterOrder.length).padStart(2, '0')}
          </span>
          <div
            aria-label={
              zh
                ? `${currentChapterCopy.label}，第 ${currentChapterNumber} 章，共 ${chapterOrder.length} 章`
                : `${currentChapterCopy.label}, chapter ${currentChapterNumber} of ${chapterOrder.length}`
            }
            aria-valuemax={chapterOrder.length}
            aria-valuemin={1}
            aria-valuenow={currentChapterNumber}
            className="office-editor-chapters__progress-dots"
            role="progressbar"
          >
            {chapterOrder.map((kind) => (
              <i
                className={kind === currentChapter ? 'is-active' : undefined}
                key={kind}
              />
            ))}
          </div>
        </div>
      </div>

      <div className="office-editor-chapters__list">
        {chapterOrder.map((kind) => {
          const chapter = copy[kind];
          return (
            <article
              className={`office-editor-chapter office-editor-chapter--${kind}`}
              id={`editor-chapter-${kind}`}
              key={kind}
              ref={(node) => {
                if (node) chapterRefs.current.set(kind, node);
                else chapterRefs.current.delete(kind);
              }}
              aria-labelledby={`editor-chapter-${kind}-title`}
              data-editor-chapter={kind}
              data-motion-active={
                !motionPaused && activeChapters[kind] ? 'true' : 'false'
              }
            >
              <div className="office-editor-chapter__copy">
                <div className="office-editor-chapter__heading">
                  <span className="office-editor-chapter__index">
                    {chapter.index}
                  </span>
                  <div>
                    <small>{chapter.component}</small>
                    <h3 id={`editor-chapter-${kind}-title`}>{chapter.title}</h3>
                  </div>
                </div>
                <p>{chapter.detail}</p>
                <ul>
                  {chapter.features.map((feature) => (
                    <li key={feature}>{feature}</li>
                  ))}
                </ul>
                <a
                  className="office-editor-chapter__link"
                  href={docsHref(kind, language)}
                >
                  {chapter.linkLabel}
                  <MotionArrow />
                </a>
                <details className="office-editor-chapter__recovery">
                  <summary>
                    {zh ? '预览无法加载？' : 'Preview unavailable?'}
                  </summary>
                  <p>
                    {zh
                      ? '打开组件文档或 Playground 继续。'
                      : 'Open the component docs or Playground to continue.'}{' '}
                    <a href={docsHref(kind, language)}>
                      {zh ? '组件文档' : 'Component docs'}
                    </a>
                    <span aria-hidden="true"> · </span>
                    <a href={playgroundHref}>Playground</a>
                  </p>
                </details>
              </div>
              <figure
                className="office-editor-chapter__preview"
                role="img"
                aria-label={
                  zh
                    ? `${chapter.label} ${chapter.component} UI/UX 动画预览`
                    : `${chapter.label} ${chapter.component} UI/UX motion preview`
                }
              >
                {renderPreview(kind, language)}
                <figcaption>
                  <span>{zh ? 'UI/UX 动画预览' : 'UI/UX motion preview'}</span>
                  <b>
                    {zh
                      ? '轻量展示 · 不挂载编辑器运行时'
                      : 'Lightweight preview · no editor runtime mounted'}
                  </b>
                </figcaption>
              </figure>
            </article>
          );
        })}
      </div>
    </section>
  );
}
