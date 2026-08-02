import {
  type KeyboardEvent,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import type { WorkDocumentSectionLayout } from '../work-types';
import {
  DocumentPageThumbnail,
  type WorkDocumentPageThumbnailSource,
} from './document-page-thumbnail';

const DOCUMENT_PAGE_WINDOW_THRESHOLD = 48;
const DOCUMENT_PAGE_WINDOW_OVERSCAN = 4;
const DOCUMENT_PAGE_FALLBACK_VIEWPORT_HEIGHT = 720;
const DOCUMENT_PAGE_LIST_PADDING_TOP = 12;
const DOCUMENT_PAGE_ITEM_GAP = 12;
const DOCUMENT_PAGE_PORTRAIT_WIDTH = 138;
const DOCUMENT_PAGE_LANDSCAPE_WIDTH = 170;
const DOCUMENT_PAGE_BUTTON_CHROME_HEIGHT = 36;
export const DOCUMENT_PAGE_WINDOW_LIMIT = 24;

interface DocumentPageRange {
  end: number;
  start: number;
  windowed: boolean;
}

interface DocumentPageSpacerEntry {
  end: number;
  height: number;
  kind: 'spacer';
  position: 'after' | 'before' | 'between';
  start: number;
}

interface DocumentPageItemEntry {
  index: number;
  kind: 'page';
}

type DocumentPageListEntry = DocumentPageItemEntry | DocumentPageSpacerEntry;

export interface DocumentNavigationPage {
  backgroundColor?: string;
  physicalPage: number;
  pageNumber: number;
  orientation: WorkDocumentSectionLayout['orientation'];
  previewText: string;
  selectionPosition: number;
}

export function DocumentPageNavigation({
  currentPage,
  pages,
  thumbnailSource,
  onSelectPage,
}: {
  currentPage: number;
  pages: readonly DocumentNavigationPage[];
  thumbnailSource?: WorkDocumentPageThumbnailSource;
  onSelectPage: (page: DocumentNavigationPage) => void | Promise<void>;
}) {
  const viewportRef = useRef<HTMLElement>(null);
  const pageRefs = useRef(new Map<number, HTMLButtonElement>());
  const pagesRef = useRef(pages);
  const pendingFocusPageRef = useRef<number | null>(null);
  const initialIndex = documentPageIndex(pages, currentPage);
  const initialPhysicalPage = pages[initialIndex]?.physicalPage ?? currentPage;
  const [selectedPage, setSelectedPage] = useState(initialPhysicalPage);
  const [rovingPage, setRovingPage] = useState(initialPhysicalPage);
  const [anchorIndex, setAnchorIndex] = useState(Math.max(0, initialIndex));
  const [viewportHeight, setViewportHeight] = useState(
    DOCUMENT_PAGE_FALLBACK_VIEWPORT_HEIGHT,
  );
  pagesRef.current = pages;

  const currentPageIndex = documentPageIndex(pages, currentPage);
  const resolvedCurrentPage = pages[currentPageIndex]?.physicalPage ?? null;
  useEffect(() => {
    if (resolvedCurrentPage === null) return;
    setSelectedPage(resolvedCurrentPage);
    setRovingPage(resolvedCurrentPage);
    setAnchorIndex(currentPageIndex);
  }, [currentPageIndex, resolvedCurrentPage]);

  useEffect(() => {
    if (!pages.length) return;
    const fallback = pages[Math.max(0, currentPageIndex)] ?? pages[0];
    if (!fallback) return;
    if (!pages.some((page) => page.physicalPage === selectedPage)) {
      setSelectedPage(fallback.physicalPage);
    }
    if (!pages.some((page) => page.physicalPage === rovingPage)) {
      setRovingPage(fallback.physicalPage);
    }
  }, [currentPageIndex, pages, rovingPage, selectedPage]);

  useLayoutEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;
    const update = () => {
      const nextHeight =
        viewport.clientHeight || DOCUMENT_PAGE_FALLBACK_VIEWPORT_HEIGHT;
      setViewportHeight((current) =>
        current === nextHeight ? current : nextHeight,
      );
    };
    update();
    if (typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(update);
    observer.observe(viewport);
    return () => observer.disconnect();
  }, []);

  const itemHeights = useMemo(
    () =>
      pages.map((page) =>
        estimatedDocumentPageItemHeight(page, thumbnailSource),
      ),
    [pages, thumbnailSource?.pageHeight, thumbnailSource?.pageWidth],
  );
  const pageStarts = useMemo(
    () => documentPageStarts(itemHeights),
    [itemHeights],
  );
  const averageItemHeight = itemHeights.length
    ? pageStarts[pageStarts.length - 1] / itemHeights.length
    : 1;
  const range = useMemo(
    () =>
      calculateDocumentPageRange({
        anchorIndex,
        averageItemHeight,
        pageCount: pages.length,
        viewportHeight,
      }),
    [anchorIndex, averageItemHeight, pages.length, viewportHeight],
  );
  const mountedIndices = useMemo(
    () =>
      documentMountedPageIndices({
        pages,
        range,
        rovingPage,
        selectedPage,
      }),
    [pages, range, rovingPage, selectedPage],
  );
  const listEntries = useMemo(
    () => documentPageListEntries(mountedIndices, pageStarts, pages.length),
    [mountedIndices, pageStarts, pages.length],
  );
  const mountedPageKey = mountedIndices.join(',');

  useLayoutEffect(() => {
    pageRefs.current
      .get(selectedPage)
      ?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
  }, [selectedPage]);

  useLayoutEffect(() => {
    const pendingPage = pendingFocusPageRef.current;
    if (pendingPage === null) return;
    const element = pageRefs.current.get(pendingPage);
    if (!element) return;
    pendingFocusPageRef.current = null;
    focusDocumentPageButton(element);
  }, [mountedPageKey, rovingPage]);

  const focusPage = useCallback((pageNumber: number) => {
    const pageIndex = documentPageIndex(pagesRef.current, pageNumber);
    if (pageIndex < 0) return;
    setRovingPage(pageNumber);
    setAnchorIndex(pageIndex);
    const mounted = pageRefs.current.get(pageNumber);
    if (mounted) {
      pendingFocusPageRef.current = null;
      focusDocumentPageButton(mounted);
      return;
    }
    pendingFocusPageRef.current = pageNumber;
  }, []);

  const handleKeyDown = (
    event: KeyboardEvent<HTMLButtonElement>,
    page: DocumentNavigationPage,
  ) => {
    const index = documentPageIndex(pages, page.physicalPage);
    const focusAt = (requestedIndex: number) => {
      const next = pages[requestedIndex];
      if (next) focusPage(next.physicalPage);
    };
    if (event.key === 'ArrowDown' || event.key === 'ArrowRight') {
      event.preventDefault();
      focusAt(Math.min(pages.length - 1, index + 1));
    } else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') {
      event.preventDefault();
      focusAt(Math.max(0, index - 1));
    } else if (event.key === 'Home') {
      event.preventDefault();
      focusAt(0);
    } else if (event.key === 'End') {
      event.preventDefault();
      focusAt(pages.length - 1);
    }
  };

  return (
    <nav
      ref={viewportRef}
      className="work-document-task-pane-body work-document-page-navigation"
      aria-label="文档页面"
      data-document-page-count={pages.length}
      data-document-page-mounted-count={mountedIndices.length}
      data-document-page-window-end={range.end}
      data-document-page-window-limit={DOCUMENT_PAGE_WINDOW_LIMIT}
      data-document-page-window-start={range.start}
      data-document-page-windowed={range.windowed ? 'true' : 'false'}
      onScroll={(event) => {
        if (!range.windowed) return;
        const nextAnchor = documentPageIndexAtOffset(
          pageStarts,
          Math.max(
            0,
            event.currentTarget.scrollTop - DOCUMENT_PAGE_LIST_PADDING_TOP,
          ),
          pages.length,
        );
        setAnchorIndex((current) =>
          current === nextAnchor ? current : nextAnchor,
        );
      }}
    >
      {pages.length ? (
        <ol>
          {listEntries.map((entry) => {
            if (entry.kind === 'spacer') {
              return (
                <li
                  aria-hidden="true"
                  className="work-document-page-navigation-spacer"
                  data-document-page-spacer={entry.position}
                  data-document-page-spacer-end={entry.end}
                  data-document-page-spacer-start={entry.start + 1}
                  key={`spacer-${entry.start}-${entry.end}`}
                  role="presentation"
                  style={{ height: `${entry.height}px` }}
                />
              );
            }
            const page = pages[entry.index];
            if (!page) return null;
            const current = page.physicalPage === selectedPage;
            return (
              <li
                aria-posinset={entry.index + 1}
                aria-setsize={pages.length}
                data-document-page-item={page.physicalPage}
                data-document-page-orientation={page.orientation}
                key={page.physicalPage}
              >
                <button
                  ref={(element) => {
                    if (element) {
                      pageRefs.current.set(page.physicalPage, element);
                    } else {
                      pageRefs.current.delete(page.physicalPage);
                    }
                  }}
                  type="button"
                  className={current ? 'active' : undefined}
                  aria-current={current ? 'page' : undefined}
                  aria-label={`第 ${page.physicalPage} 页`}
                  data-document-page-thumbnail={page.physicalPage}
                  tabIndex={page.physicalPage === rovingPage ? 0 : -1}
                  onFocus={() => setRovingPage(page.physicalPage)}
                  onKeyDown={(event) => handleKeyDown(event, page)}
                  onClick={() => {
                    setSelectedPage(page.physicalPage);
                    setAnchorIndex(entry.index);
                    void onSelectPage(page);
                  }}
                >
                  <DocumentPageThumbnail
                    backgroundColor={page.backgroundColor}
                    fallbackText={page.previewText}
                    orientation={page.orientation}
                    pageIndex={page.physicalPage - 1}
                    priority={Math.abs(page.physicalPage - selectedPage) <= 1}
                    source={thumbnailSource}
                  />
                  <span className="work-document-page-thumbnail-label">
                    第 {page.physicalPage} 页
                    {page.pageNumber !== page.physicalPage && (
                      <small>页码 {page.pageNumber}</small>
                    )}
                  </span>
                </button>
              </li>
            );
          })}
        </ol>
      ) : (
        <div className="work-document-outline-empty">正在生成页面预览…</div>
      )}
    </nav>
  );
}

export function calculateDocumentPageRange({
  anchorIndex,
  averageItemHeight,
  pageCount,
  viewportHeight,
}: {
  anchorIndex: number;
  averageItemHeight: number;
  pageCount: number;
  viewportHeight: number;
}): DocumentPageRange {
  const count = Math.max(0, Math.floor(pageCount));
  if (count <= DOCUMENT_PAGE_WINDOW_THRESHOLD) {
    return { end: count, start: 0, windowed: false };
  }
  const visibleCount = Math.max(
    1,
    Math.ceil(Math.max(1, viewportHeight) / Math.max(1, averageItemHeight)),
  );
  const windowCount = Math.min(
    count,
    DOCUMENT_PAGE_WINDOW_LIMIT,
    visibleCount + DOCUMENT_PAGE_WINDOW_OVERSCAN * 2,
  );
  const boundedAnchor = Math.min(
    count - 1,
    Math.max(0, Math.floor(anchorIndex)),
  );
  const start = Math.min(
    count - windowCount,
    Math.max(0, boundedAnchor - DOCUMENT_PAGE_WINDOW_OVERSCAN),
  );
  return { end: start + windowCount, start, windowed: true };
}

function documentPageIndex(
  pages: readonly DocumentNavigationPage[],
  physicalPage: number,
): number {
  return pages.findIndex((page) => page.physicalPage === physicalPage);
}

function estimatedDocumentPageItemHeight(
  page: DocumentNavigationPage,
  source: WorkDocumentPageThumbnailSource | undefined,
): number {
  const width =
    page.orientation === 'landscape'
      ? DOCUMENT_PAGE_LANDSCAPE_WIDTH
      : DOCUMENT_PAGE_PORTRAIT_WIDTH;
  const sourceRatio =
    source && source.pageWidth > 0 && source.pageHeight > 0
      ? source.pageHeight / source.pageWidth
      : null;
  const fallbackRatio =
    page.orientation === 'landscape' ? 210 / 297 : 297 / 210;
  return Math.ceil(
    width * (sourceRatio ?? fallbackRatio) + DOCUMENT_PAGE_BUTTON_CHROME_HEIGHT,
  );
}

function documentPageStarts(itemHeights: readonly number[]): number[] {
  const starts = [0];
  for (const height of itemHeights) {
    starts.push(
      (starts[starts.length - 1] ?? 0) +
        Math.max(1, height) +
        DOCUMENT_PAGE_ITEM_GAP,
    );
  }
  return starts;
}

function documentMountedPageIndices({
  pages,
  range,
  rovingPage,
  selectedPage,
}: {
  pages: readonly DocumentNavigationPage[];
  range: DocumentPageRange;
  rovingPage: number;
  selectedPage: number;
}): number[] {
  if (!range.windowed) {
    return Array.from({ length: pages.length }, (_, index) => index);
  }
  const mounted = new Set<number>();
  for (let index = range.start; index < range.end; index += 1) {
    mounted.add(index);
  }
  const selectedIndex = documentPageIndex(pages, selectedPage);
  const rovingIndex = documentPageIndex(pages, rovingPage);
  if (selectedIndex >= 0) mounted.add(selectedIndex);
  if (rovingIndex >= 0) mounted.add(rovingIndex);
  return [...mounted].sort((left, right) => left - right);
}

function documentPageListEntries(
  mountedIndices: readonly number[],
  pageStarts: readonly number[],
  pageCount: number,
): DocumentPageListEntry[] {
  const entries: DocumentPageListEntry[] = [];
  let cursor = 0;
  for (const index of mountedIndices) {
    if (index > cursor) {
      entries.push(
        documentPageSpacerEntry(cursor, index, pageStarts, pageCount),
      );
    }
    entries.push({ index, kind: 'page' });
    cursor = index + 1;
  }
  if (cursor < pageCount) {
    entries.push(
      documentPageSpacerEntry(cursor, pageCount, pageStarts, pageCount),
    );
  }
  return entries;
}

function documentPageSpacerEntry(
  start: number,
  end: number,
  pageStarts: readonly number[],
  pageCount: number,
): DocumentPageSpacerEntry {
  return {
    end,
    height: Math.max(
      0,
      (pageStarts[end] ?? 0) -
        (pageStarts[start] ?? 0) -
        DOCUMENT_PAGE_ITEM_GAP,
    ),
    kind: 'spacer',
    position: start === 0 ? 'before' : end === pageCount ? 'after' : 'between',
    start,
  };
}

function documentPageIndexAtOffset(
  pageStarts: readonly number[],
  offset: number,
  pageCount: number,
): number {
  if (pageCount <= 1) return 0;
  let low = 0;
  let high = pageCount - 1;
  while (low <= high) {
    const middle = Math.floor((low + high) / 2);
    const start = pageStarts[middle] ?? 0;
    const next = pageStarts[middle + 1] ?? Number.POSITIVE_INFINITY;
    if (offset < start) high = middle - 1;
    else if (offset >= next) low = middle + 1;
    else return middle;
  }
  return Math.min(pageCount - 1, Math.max(0, low));
}

function focusDocumentPageButton(element: HTMLButtonElement): void {
  element.focus({ preventScroll: true });
  element.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
}
