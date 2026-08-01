import {
  DocumentManagerPlugin,
  type PluginRegistry,
  ThumbnailPlugin,
} from '@embedpdf/react-pdf-viewer';
import { X } from 'lucide-react';
import {
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type RefObject,
} from 'react';

const PDF_THUMBNAIL_ITEM_HEIGHT = 190;
const PDF_THUMBNAIL_WINDOW_THRESHOLD = 48;
const PDF_THUMBNAIL_OVERSCAN = 5;
const PDF_THUMBNAIL_FALLBACK_VIEWPORT_HEIGHT = 720;
export const PDF_THUMBNAIL_WINDOW_LIMIT = 32;

interface PdfThumbnailRange {
  end: number;
  start: number;
  windowed: boolean;
}

export interface PdfThumbnailRailProps {
  currentPage: number;
  mobileCloseButtonRef?: RefObject<HTMLButtonElement | null>;
  mobileNavigationId?: string;
  mobileNavigationModal?: boolean;
  registry: PluginRegistry;
  totalPages: number;
  onCloseMobileNavigation?: () => void;
  onSelectPage: (page: number) => void;
}

export function PdfThumbnailRail({
  currentPage,
  mobileCloseButtonRef,
  mobileNavigationId,
  mobileNavigationModal = false,
  registry,
  totalPages,
  onCloseMobileNavigation,
  onSelectPage,
}: PdfThumbnailRailProps) {
  const viewportRef = useRef<HTMLElement>(null);
  const pendingKeyboardFocusPageRef = useRef<number | null>(null);
  const [anchorIndex, setAnchorIndex] = useState(() =>
    pageIndex(currentPage, totalPages),
  );
  const [viewportHeight, setViewportHeight] = useState(
    PDF_THUMBNAIL_FALLBACK_VIEWPORT_HEIGHT,
  );
  const [itemHeight, setItemHeight] = useState(PDF_THUMBNAIL_ITEM_HEIGHT);

  useEffect(() => {
    setAnchorIndex(pageIndex(currentPage, totalPages));
  }, [currentPage, totalPages]);

  useLayoutEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;
    const update = () => {
      if (viewport.clientHeight > 0) setViewportHeight(viewport.clientHeight);
    };
    update();
    if (typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(update);
    observer.observe(viewport);
    return () => observer.disconnect();
  }, []);

  useLayoutEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport || currentPage < 1) return;
    const index = pageIndex(currentPage, totalPages);
    const pageTop = index * itemHeight;
    const pageBottom = pageTop + itemHeight;
    const visibleTop = viewport.scrollTop;
    const visibleBottom = visibleTop + viewportHeight;
    if (pageTop >= visibleTop && pageBottom <= visibleBottom) return;
    const top = Math.max(0, pageTop - (viewportHeight - itemHeight) / 2);
    if (typeof viewport.scrollTo === 'function') {
      viewport.scrollTo({ top, behavior: 'smooth' });
    } else {
      viewport.scrollTop = top;
    }
  }, [currentPage, itemHeight, totalPages, viewportHeight]);

  const range = useMemo(
    () =>
      calculatePdfThumbnailRange({
        anchorIndex,
        itemHeight,
        totalPages,
        viewportHeight,
      }),
    [anchorIndex, itemHeight, totalPages, viewportHeight],
  );
  const pages = useMemo(
    () =>
      Array.from(
        { length: Math.max(0, range.end - range.start) },
        (_, index) => range.start + index + 1,
      ),
    [range.end, range.start],
  );
  useLayoutEffect(() => {
    const pendingPage = pendingKeyboardFocusPageRef.current;
    if (pendingPage === null || currentPage !== pendingPage) return;
    const thumbnail = viewportRef.current?.querySelector<HTMLButtonElement>(
      `[data-pdf-page-index="${pendingPage - 1}"]`,
    );
    if (!thumbnail) return;
    pendingKeyboardFocusPageRef.current = null;
    thumbnail.focus({ preventScroll: true });
  }, [currentPage, range.end, range.start]);
  useLayoutEffect(() => {
    const viewport = viewportRef.current;
    const thumbnail = viewport?.querySelector<HTMLElement>(
      '[data-pdf-page-thumbnail]',
    );
    const list = viewport?.querySelector<HTMLElement>(
      '.work-pdf-thumbnail-list',
    );
    if (!(thumbnail && list)) return;
    const measure = () => {
      const height = thumbnail.getBoundingClientRect().height;
      const gap = Number.parseFloat(getComputedStyle(list).rowGap) || 0;
      if (height > 0) setItemHeight(Math.round(height + gap));
    };
    measure();
    if (typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(measure);
    observer.observe(thumbnail);
    return () => observer.disconnect();
  }, [range.end, range.start]);

  const topSpacerHeight = range.start * itemHeight;
  const bottomSpacerHeight = Math.max(0, totalPages - range.end) * itemHeight;
  const modalAttributes = mobileNavigationModal
    ? ({ role: 'dialog', 'aria-modal': true } as const)
    : {};

  return (
    <aside
      {...modalAttributes}
      id={mobileNavigationId}
      className="work-pdf-thumbnail-rail"
      aria-label="PDF 页面"
      data-pdf-page-count={totalPages}
      data-pdf-thumbnail-window-end={range.end}
      data-pdf-thumbnail-window-start={range.start}
      data-pdf-thumbnail-windowed={range.windowed ? 'true' : 'false'}
    >
      <header className="work-pdf-thumbnail-header">
        <strong>页面</strong>
        <button
          ref={mobileCloseButtonRef}
          type="button"
          aria-label="关闭 PDF 页面导航"
          onClick={onCloseMobileNavigation}
        >
          <X size={16} />
        </button>
      </header>
      <nav
        ref={viewportRef}
        className="work-pdf-thumbnail-viewport"
        aria-label="PDF 页面缩略图"
        onScroll={(event) => {
          setAnchorIndex(
            Math.floor(event.currentTarget.scrollTop / itemHeight),
          );
        }}
      >
        <div className="work-pdf-thumbnail-list">
          <PdfThumbnailSpacer height={topSpacerHeight} position="before" />
          {pages.map((page) => (
            <PdfPageThumbnail
              current={page === currentPage}
              key={page}
              page={page}
              registry={registry}
              totalPages={totalPages}
              onSelectPage={onSelectPage}
              onSelectPageFromKeyboard={(nextPage) => {
                if (!mobileNavigationModal) {
                  pendingKeyboardFocusPageRef.current = nextPage;
                }
                onSelectPage(nextPage);
              }}
            />
          ))}
          <PdfThumbnailSpacer height={bottomSpacerHeight} position="after" />
        </div>
      </nav>
    </aside>
  );
}

export function calculatePdfThumbnailRange({
  anchorIndex,
  itemHeight = PDF_THUMBNAIL_ITEM_HEIGHT,
  totalPages,
  viewportHeight,
}: {
  anchorIndex: number;
  itemHeight?: number;
  totalPages: number;
  viewportHeight: number;
}): PdfThumbnailRange {
  const pageCount = Math.max(0, Math.floor(totalPages));
  if (pageCount <= PDF_THUMBNAIL_WINDOW_THRESHOLD) {
    return { end: pageCount, start: 0, windowed: false };
  }
  const visibleCount = Math.max(
    1,
    Math.ceil(Math.max(1, viewportHeight) / Math.max(1, itemHeight)),
  );
  const windowCount = Math.min(
    PDF_THUMBNAIL_WINDOW_LIMIT,
    pageCount,
    visibleCount + PDF_THUMBNAIL_OVERSCAN * 2,
  );
  const boundedAnchor = Math.min(
    pageCount - 1,
    Math.max(0, Math.floor(anchorIndex)),
  );
  const start = Math.min(
    pageCount - windowCount,
    Math.max(0, boundedAnchor - PDF_THUMBNAIL_OVERSCAN),
  );
  return { end: start + windowCount, start, windowed: true };
}

function PdfPageThumbnail({
  current,
  page,
  registry,
  totalPages,
  onSelectPage,
  onSelectPageFromKeyboard,
}: {
  current: boolean;
  page: number;
  registry: PluginRegistry;
  totalPages: number;
  onSelectPage: (page: number) => void;
  onSelectPageFromKeyboard: (page: number) => void;
}) {
  const { sourceUrl, state } = usePdfThumbnailSource(registry, page);
  return (
    <button
      type="button"
      className={current ? 'active' : undefined}
      aria-current={current ? 'page' : undefined}
      aria-label={`第 ${page} 页`}
      data-pdf-page-thumbnail
      data-pdf-page-index={page - 1}
      onClick={() => onSelectPage(page)}
      onKeyDown={(event) => {
        const nextPage = pdfThumbnailKeyboardPage(event.key, page, totalPages);
        if (nextPage === null) return;
        event.preventDefault();
        event.stopPropagation();
        onSelectPageFromKeyboard(nextPage);
      }}
    >
      <span className="work-pdf-thumbnail-number" aria-hidden="true">
        {page}
      </span>
      <span
        className="work-pdf-thumbnail-preview"
        data-pdf-thumbnail-state={state}
      >
        {sourceUrl && <img src={sourceUrl} alt="" draggable={false} />}
      </span>
    </button>
  );
}

function usePdfThumbnailSource(
  registry: PluginRegistry,
  page: number,
): { sourceUrl: string | null; state: 'error' | 'loading' | 'ready' } {
  const [sourceUrl, setSourceUrl] = useState<string | null>(null);
  const [state, setState] = useState<'error' | 'loading' | 'ready'>('loading');

  useEffect(() => {
    let disposed = false;
    let objectUrl: string | null = null;
    setSourceUrl(null);
    setState('loading');
    void renderPdfThumbnail(registry, page)
      .then((blob) => {
        objectUrl = URL.createObjectURL(blob);
        if (disposed) {
          URL.revokeObjectURL(objectUrl);
          return;
        }
        setSourceUrl(objectUrl);
        setState('ready');
      })
      .catch(() => {
        if (!disposed) setState('error');
      });
    return () => {
      disposed = true;
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [page, registry]);

  return { sourceUrl, state };
}

async function renderPdfThumbnail(
  registry: PluginRegistry,
  page: number,
): Promise<Blob> {
  await registry.pluginsReady();
  const documentManager = registry
    .getPlugin<DocumentManagerPlugin>(DocumentManagerPlugin.id)
    ?.provides();
  const thumbnail = registry
    .getPlugin<ThumbnailPlugin>(ThumbnailPlugin.id)
    ?.provides();
  const documentId = documentManager?.getActiveDocumentId();
  if (!documentId || !thumbnail) {
    throw new Error('PDF thumbnail rendering is unavailable.');
  }
  const devicePixelRatio =
    typeof window === 'undefined'
      ? 1
      : Math.min(2, Math.max(1, window.devicePixelRatio || 1));
  return thumbnail
    .forDocument(documentId)
    .renderThumb(page - 1, devicePixelRatio)
    .toPromise();
}

function PdfThumbnailSpacer({
  height,
  position,
}: {
  height: number;
  position: 'after' | 'before';
}) {
  if (height <= 0) return null;
  return (
    <span
      aria-hidden="true"
      className="work-pdf-thumbnail-spacer"
      data-pdf-thumbnail-spacer={position}
      style={{ height: `${height}px` }}
    />
  );
}

function pageIndex(currentPage: number, totalPages: number): number {
  return Math.min(
    Math.max(0, Math.floor(totalPages) - 1),
    Math.max(0, Math.floor(currentPage) - 1),
  );
}

function pdfThumbnailKeyboardPage(
  key: string,
  page: number,
  totalPages: number,
): number | null {
  if (key === 'Home') return 1;
  if (key === 'End') return totalPages;
  if (key === 'ArrowUp' || key === 'ArrowLeft') return Math.max(1, page - 1);
  if (key === 'ArrowDown' || key === 'ArrowRight') {
    return Math.min(totalPages, page + 1);
  }
  return null;
}
