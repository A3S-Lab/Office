import { withBase } from '@rspress/core/runtime';
import { useEffect, useRef, useState } from 'react';
import { chapterOrder, MotionArrow } from './HomeEditorContent';
import type { ChapterKind, HomeLanguage } from './HomeEditorContent';
import {
  editorScreenshotFiles,
  editorScreenshotLabels,
} from './home-editor-assets';

interface ChapterCopy {
  index: string;
  label: string;
  component: string;
  title: string;
  detail: string;
  linkLabel: string;
}

const chapterCopy: Record<HomeLanguage, Record<ChapterKind, ChapterCopy>> = {
  zh: {
    document: {
      index: '01',
      label: '文档',
      component: 'DocumentEditor',
      title: '写文档，批注和修订一起看。',
      detail:
        '功能区、页面、批注和修订在同一个编辑器里；协作者的光标会显示在页面上。',
      linkLabel: '查看 DocumentEditor',
    },
    markdown: {
      index: '02',
      label: 'Markdown',
      component: 'MarkdownEditor',
      title: '改源码，预览就更新。',
      detail: '在源码、分屏和预览之间切换，内容和光标位置都保留。',
      linkLabel: '查看 MarkdownEditor',
    },
    spreadsheet: {
      index: '03',
      label: '表格',
      component: 'SpreadsheetEditor',
      title: '选中单元格，看公式和校验。',
      detail:
        '公式栏、表格样式和数据验证围绕当前单元格显示；空白区域按需绘制。',
      linkLabel: '查看 SpreadsheetEditor',
    },
    presentation: {
      index: '04',
      label: '演示文稿',
      component: 'PresentationEditor',
      title: '选中对象，设置入场动画。',
      detail:
        '幻灯片、对象、动画时间线和演讲者备注分别显示；编辑文字时再打开文字编辑器。',
      linkLabel: '查看 PresentationEditor',
    },
    pdf: {
      index: '05',
      label: 'PDF',
      component: 'PdfViewer',
      title: '看 PDF 时整理页面和批注。',
      detail: '缩略图、批注和页面操作都在阅读区里；页面重排交给 Worker。',
      linkLabel: '查看 PdfViewer',
    },
  },
  en: {
    document: {
      index: '01',
      label: 'Document',
      component: 'DocumentEditor',
      title: 'Write, comment, and review in one view.',
      detail:
        'The ribbon, page, comments, and revisions stay in one editor; collaborators see remote cursors on the page.',
      linkLabel: 'View DocumentEditor',
    },
    markdown: {
      index: '02',
      label: 'Markdown',
      component: 'MarkdownEditor',
      title: 'Edit the source. The preview follows.',
      detail:
        'Move between source, split view, and preview without losing the content or caret position.',
      linkLabel: 'View MarkdownEditor',
    },
    spreadsheet: {
      index: '03',
      label: 'Spreadsheet',
      component: 'SpreadsheetEditor',
      title: 'Select a cell. See its formula and validation.',
      detail:
        'The formula bar, table styles, and data validation follow the active cell; empty areas are drawn on demand.',
      linkLabel: 'View SpreadsheetEditor',
    },
    presentation: {
      index: '04',
      label: 'Presentation',
      component: 'PresentationEditor',
      title: 'Select an object. Set its entrance animation.',
      detail:
        'Slides, objects, animation timing, and presenter notes have separate views; text editing opens when needed.',
      linkLabel: 'View PresentationEditor',
    },
    pdf: {
      index: '05',
      label: 'PDF',
      component: 'PdfViewer',
      title: 'Read the PDF while you organize pages and notes.',
      detail:
        'Thumbnails, annotations, and page operations share the reading workspace; page reordering runs in a Worker.',
      linkLabel: 'View PdfViewer',
    },
  },
};

function RealEditorPreview({
  kind,
  language,
}: {
  kind: ChapterKind;
  language: HomeLanguage;
}) {
  const zh = language === 'zh';
  const label = editorScreenshotLabels[language][kind];

  return (
    <div
      className={`office-motion-window office-real-editor-shot office-real-editor-shot--${kind}`}
      data-real-editor-screenshot="true"
      data-screenshot-kind={kind}
    >
      <div className="office-real-editor-shot__viewport">
        <img
          src={withBase(`/editor-previews/${editorScreenshotFiles[kind]}`)}
          alt={
            zh ? `${label}真实界面截图` : `Real ${label} interface screenshot`
          }
          loading={kind === 'document' ? 'eager' : 'lazy'}
          decoding="async"
        />
        <span className="office-real-editor-shot__scan" aria-hidden="true" />
      </div>
    </div>
  );
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
    >
      <header className="office-editor-chapters__header">
        <div>
          <h2 id="office-editor-chapters-title">
            {zh ? (
              <>
                五种编辑器，
                <br />
                分别看一个操作。
              </>
            ) : (
              <>
                Five editors.
                <br />
                One action from each.
              </>
            )}
          </h2>
          <p>
            {zh
              ? '下面的截图来自实际编辑器，每段只展示一个常用操作。完整界面可在 Playground 打开。'
              : 'These captures come from the editors themselves. Each section shows one common action; open the Playground for the full interface.'}
          </p>
        </div>
        <div className="office-editor-chapters__header-actions">
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
          <strong>
            {currentChapterCopy.index} · {currentChapterCopy.label}
          </strong>
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
                    <h3 id={`editor-chapter-${kind}-title`}>{chapter.title}</h3>
                  </div>
                </div>
                <p>{chapter.detail}</p>
                <a
                  className="office-editor-chapter__link"
                  href={docsHref(kind, language)}
                >
                  {chapter.linkLabel}
                  <MotionArrow />
                </a>
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
                <RealEditorPreview kind={kind} language={language} />
              </figure>
            </article>
          );
        })}
      </div>
    </section>
  );
}
