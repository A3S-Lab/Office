import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type RefObject,
} from 'react';

const THUMBNAIL_WINDOW_THRESHOLD = 60;
const NORMAL_INITIAL_COLUMNS = 1;
const SORTER_INITIAL_COLUMNS = 4;
const NORMAL_ITEM_HEIGHT = 88;
const SORTER_ITEM_HEIGHT = 160;
const NORMAL_ROW_GAP = 10;
const SORTER_ROW_GAP = 22;
const NORMAL_OVERSCAN_ROWS = 4;
const SORTER_OVERSCAN_ROWS = 2;
const FALLBACK_VIEWPORT_HEIGHT = 720;

interface ThumbnailRange {
  start: number;
  end: number;
}

interface ThumbnailMetrics {
  columns: number;
  itemHeight: number;
  rowGap: number;
  viewportRows: number;
}

interface ThumbnailLayoutAnchor {
  key: string;
  slideId: string;
}

export interface PresentationThumbnailWindow {
  bottomSpacerHeight: number;
  end: number;
  requestFocus: (index: number) => void;
  requestFocusById: (slideId: string) => void;
  start: number;
  topSpacerHeight: number;
  windowed: boolean;
}

export function usePresentationThumbnailWindow({
  slideIds,
  selectedSlideId,
  viewMode,
  viewportRef,
  zoom,
}: {
  slideIds: readonly string[];
  selectedSlideId: string;
  viewMode: 'normal' | 'sorter';
  viewportRef: RefObject<HTMLElement | null>;
  zoom: number;
}): PresentationThumbnailWindow {
  const itemCount = slideIds.length;
  const windowed = itemCount > THUMBNAIL_WINDOW_THRESHOLD;
  const selectedIndex = Math.max(0, slideIds.indexOf(selectedSlideId));
  const initialMetrics = defaultMetrics(viewMode, zoom);
  const metricsRef = useRef(initialMetrics);
  const slideIdsRef = useRef(slideIds);
  const [, setMetricsVersion] = useState(0);
  const pendingFocusIdRef = useRef<string | null>(null);
  const pendingRevealIdRef = useRef<string | null>(null);
  const pendingLayoutAnchorRef = useRef<ThumbnailLayoutAnchor | null>(null);
  const layoutKey = `${viewMode}:${zoom}`;
  const layoutKeyRef = useRef(layoutKey);
  const activeLayoutKeyRef = useRef(layoutKey);
  const frameRef = useRef<number | null>(null);
  const scheduleMeasurementRef = useRef<() => void>(() => undefined);
  const [range, setRange] = useState<ThumbnailRange>(() =>
    centeredRange(itemCount, selectedIndex, initialMetrics),
  );
  slideIdsRef.current = slideIds;
  activeLayoutKeyRef.current = layoutKey;
  const slideIdSignature = useMemo(
    () => slideIds.map((slideId) => `${slideId.length}:${slideId}`).join(''),
    [slideIds],
  );

  const focusPendingThumbnail = useCallback(() => {
    const slideId = pendingFocusIdRef.current;
    const viewport = viewportRef.current;
    if (!slideId || !viewport) return;
    const thumbnail = Array.from(
      viewport.querySelectorAll<HTMLButtonElement>('[data-slide-thumbnail]'),
    ).find((candidate) => candidate.dataset.slideId === slideId);
    if (!thumbnail) return;
    pendingFocusIdRef.current = null;
    thumbnail.focus();
    thumbnail.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }, [viewportRef]);

  const revealThumbnail = useCallback(
    (slideId: string): boolean => {
      const viewport = viewportRef.current;
      if (!viewport) return false;
      const thumbnail = Array.from(
        viewport.querySelectorAll<HTMLButtonElement>('[data-slide-thumbnail]'),
      ).find((candidate) => candidate.dataset.slideId === slideId);
      if (!thumbnail) return false;
      thumbnail.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      return true;
    },
    [viewportRef],
  );

  const revealPendingThumbnail = useCallback(() => {
    const slideId = pendingRevealIdRef.current;
    if (!slideId || !revealThumbnail(slideId)) return;
    pendingRevealIdRef.current = null;
  }, [revealThumbnail]);

  const scheduleMeasurement = useCallback(() => {
    if (frameRef.current !== null) cancelAnimationFrame(frameRef.current);
    frameRef.current = requestAnimationFrame(() => {
      frameRef.current = null;
      if (activeLayoutKeyRef.current !== layoutKey) return;
      const pendingLayoutAnchor = pendingLayoutAnchorRef.current;
      if (pendingLayoutAnchor && pendingLayoutAnchor.key !== layoutKey) return;
      const viewport = viewportRef.current;
      if (!viewport || !windowed) return;
      const list = viewport.querySelector<HTMLElement>(
        '[data-slide-thumbnail-list]',
      );
      const thumbnail = list?.querySelector<HTMLElement>(
        '[data-slide-thumbnail]',
      );
      if (!list || !thumbnail) return;
      const listStyle = getComputedStyle(list);
      const measuredHeight = thumbnail.getBoundingClientRect().height;
      const columns =
        viewMode === 'sorter'
          ? renderedGridColumnCount(listStyle.gridTemplateColumns)
          : 1;
      const rowGap =
        finiteCssPixels(listStyle.rowGap) ??
        (viewMode === 'sorter' ? SORTER_ROW_GAP : NORMAL_ROW_GAP);
      const itemHeight =
        measuredHeight > 0
          ? measuredHeight
          : metricsRef.current.itemHeight ||
            defaultMetrics(viewMode, zoom).itemHeight;
      const viewportHeight = viewport.clientHeight || FALLBACK_VIEWPORT_HEIGHT;
      const viewportRows = Math.max(
        1,
        Math.ceil(viewportHeight / Math.max(1, itemHeight + rowGap)),
      );
      const metrics = {
        columns,
        itemHeight,
        rowGap,
        viewportRows,
      };
      const metricsChanged = !sameMetrics(metricsRef.current, metrics);
      metricsRef.current = metrics;
      if (metricsChanged) setMetricsVersion((version) => version + 1);
      const listTop = list.offsetTop;
      const firstVisibleRow = Math.max(
        0,
        Math.floor(
          Math.max(0, viewport.scrollTop - listTop) /
            Math.max(1, itemHeight + rowGap),
        ),
      );
      const overscanRows =
        viewMode === 'sorter' ? SORTER_OVERSCAN_ROWS : NORMAL_OVERSCAN_ROWS;
      const visibleRange = rowRange(
        itemCount,
        columns,
        Math.max(0, firstVisibleRow - overscanRows),
        viewportRows + overscanRows * 2,
      );
      if (pendingLayoutAnchor?.key === layoutKey) {
        const anchorIndex = slideIdsRef.current.indexOf(
          pendingLayoutAnchor.slideId,
        );
        if (anchorIndex < 0) {
          pendingLayoutAnchorRef.current = null;
        } else if (
          !metricsChanged &&
          anchorIndex >= visibleRange.start &&
          anchorIndex < visibleRange.end
        ) {
          pendingLayoutAnchorRef.current = null;
        } else {
          setRange((current) => {
            const next = centeredRange(itemCount, anchorIndex, metrics);
            return sameRange(current, next) ? current : next;
          });
          requestAnimationFrame(() => {
            revealThumbnail(pendingLayoutAnchor.slideId);
            requestAnimationFrame(() => scheduleMeasurementRef.current());
          });
          return;
        }
      }
      if (!pendingFocusIdRef.current && !pendingRevealIdRef.current) {
        setRange((current) => {
          return sameRange(current, visibleRange) ? current : visibleRange;
        });
      }
      focusPendingThumbnail();
      revealPendingThumbnail();
    });
  }, [
    focusPendingThumbnail,
    itemCount,
    layoutKey,
    revealPendingThumbnail,
    revealThumbnail,
    viewMode,
    viewportRef,
    windowed,
    zoom,
  ]);
  scheduleMeasurementRef.current = scheduleMeasurement;

  useEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport || !windowed) {
      setRange({ start: 0, end: itemCount });
      return;
    }
    const onScroll = () => scheduleMeasurement();
    viewport.addEventListener('scroll', onScroll, { passive: true });
    const resizeObserver =
      typeof ResizeObserver === 'undefined'
        ? null
        : new ResizeObserver(scheduleMeasurement);
    resizeObserver?.observe(viewport);
    const list = viewport.querySelector<HTMLElement>(
      '[data-slide-thumbnail-list]',
    );
    if (list) resizeObserver?.observe(list);
    window.addEventListener('resize', scheduleMeasurement);
    scheduleMeasurement();
    return () => {
      viewport.removeEventListener('scroll', onScroll);
      resizeObserver?.disconnect();
      window.removeEventListener('resize', scheduleMeasurement);
      if (frameRef.current !== null) {
        cancelAnimationFrame(frameRef.current);
        frameRef.current = null;
      }
    };
  }, [
    itemCount,
    scheduleMeasurement,
    slideIdSignature,
    viewMode,
    viewportRef,
    windowed,
    zoom,
  ]);

  useEffect(() => {
    if (!windowed) return;
    const layoutChanged = layoutKeyRef.current !== layoutKey;
    layoutKeyRef.current = layoutKey;
    pendingRevealIdRef.current = selectedSlideId;
    if (layoutChanged) {
      const metrics = defaultMetrics(viewMode, zoom);
      metricsRef.current = metrics;
      pendingLayoutAnchorRef.current = {
        key: layoutKey,
        slideId: selectedSlideId,
      };
      setRange(centeredRange(itemCount, selectedIndex, metrics));
      return;
    }
    if (pendingLayoutAnchorRef.current?.key === layoutKey) {
      pendingLayoutAnchorRef.current = {
        key: layoutKey,
        slideId: selectedSlideId,
      };
    }
    setRange((current) => {
      if (selectedIndex >= current.start && selectedIndex < current.end) {
        return current;
      }
      return centeredRange(itemCount, selectedIndex, metricsRef.current);
    });
    const frame = requestAnimationFrame(revealPendingThumbnail);
    return () => cancelAnimationFrame(frame);
  }, [
    itemCount,
    revealPendingThumbnail,
    selectedIndex,
    selectedSlideId,
    layoutKey,
    viewMode,
    windowed,
    zoom,
  ]);

  useEffect(() => {
    if (pendingLayoutAnchorRef.current) return;
    if (!pendingFocusIdRef.current && !pendingRevealIdRef.current) return;
    const frame = requestAnimationFrame(() => {
      focusPendingThumbnail();
      revealPendingThumbnail();
    });
    return () => cancelAnimationFrame(frame);
  });

  const requestFocusById = useCallback(
    (slideId: string) => {
      const currentSlideIds = slideIdsRef.current;
      const index = currentSlideIds.indexOf(slideId);
      if (index < 0) return;
      pendingFocusIdRef.current = slideId;
      if (windowed) {
        setRange((current) =>
          index >= current.start && index < current.end
            ? current
            : centeredRange(currentSlideIds.length, index, metricsRef.current),
        );
      }
      requestAnimationFrame(focusPendingThumbnail);
    },
    [focusPendingThumbnail, windowed],
  );
  const requestFocus = useCallback(
    (index: number) => {
      const currentSlideIds = slideIdsRef.current;
      const boundedIndex = Math.min(
        Math.max(0, index),
        Math.max(0, currentSlideIds.length - 1),
      );
      const slideId = currentSlideIds[boundedIndex];
      if (slideId) requestFocusById(slideId);
    },
    [requestFocusById],
  );

  const effectiveRange = windowed
    ? normalizedRange(range, itemCount)
    : { start: 0, end: itemCount };
  const spacers = windowed
    ? spacerHeights(itemCount, effectiveRange, metricsRef.current)
    : { top: 0, bottom: 0 };

  return {
    bottomSpacerHeight: spacers.bottom,
    end: effectiveRange.end,
    requestFocus,
    requestFocusById,
    start: effectiveRange.start,
    topSpacerHeight: spacers.top,
    windowed,
  };
}

function defaultMetrics(
  viewMode: 'normal' | 'sorter',
  zoom: number,
): ThumbnailMetrics {
  if (viewMode === 'sorter') {
    return {
      columns: SORTER_INITIAL_COLUMNS,
      itemHeight: Math.round(SORTER_ITEM_HEIGHT * (zoom / 100)),
      rowGap: SORTER_ROW_GAP,
      viewportRows: 5,
    };
  }
  return {
    columns: NORMAL_INITIAL_COLUMNS,
    itemHeight: NORMAL_ITEM_HEIGHT,
    rowGap: NORMAL_ROW_GAP,
    viewportRows: 9,
  };
}

function centeredRange(
  itemCount: number,
  selectedIndex: number,
  metrics: ThumbnailMetrics,
): ThumbnailRange {
  const overscanRows =
    metrics.columns > 1 ? SORTER_OVERSCAN_ROWS : NORMAL_OVERSCAN_ROWS;
  const rowCount = metrics.viewportRows + overscanRows * 2;
  const selectedRow = Math.floor(selectedIndex / metrics.columns);
  return rowRange(
    itemCount,
    metrics.columns,
    Math.max(0, selectedRow - Math.floor(rowCount / 2)),
    rowCount,
  );
}

function rowRange(
  itemCount: number,
  columns: number,
  startRow: number,
  rowCount: number,
): ThumbnailRange {
  const totalRows = Math.ceil(itemCount / Math.max(1, columns));
  const boundedRowCount = Math.min(Math.max(1, rowCount), totalRows);
  const boundedStartRow = Math.min(
    Math.max(0, startRow),
    Math.max(0, totalRows - boundedRowCount),
  );
  return {
    start: boundedStartRow * columns,
    end: Math.min(itemCount, (boundedStartRow + boundedRowCount) * columns),
  };
}

function normalizedRange(
  range: ThumbnailRange,
  itemCount: number,
): ThumbnailRange {
  const start = Math.min(Math.max(0, range.start), itemCount);
  const end = Math.min(Math.max(start, range.end), itemCount);
  return end > start || itemCount === 0
    ? { start, end }
    : { start: 0, end: Math.min(itemCount, 1) };
}

function spacerHeights(
  itemCount: number,
  range: ThumbnailRange,
  metrics: ThumbnailMetrics,
) {
  const columns = Math.max(1, metrics.columns);
  const totalRows = Math.ceil(itemCount / columns);
  const startRow = Math.floor(range.start / columns);
  const endRow = Math.ceil(range.end / columns);
  return {
    top: skippedRowsHeight(startRow, metrics),
    bottom: skippedRowsHeight(Math.max(0, totalRows - endRow), metrics),
  };
}

function skippedRowsHeight(
  rowCount: number,
  metrics: ThumbnailMetrics,
): number {
  if (!rowCount) return 0;
  return Math.max(
    0,
    rowCount * (metrics.itemHeight + metrics.rowGap) - metrics.rowGap,
  );
}

function renderedGridColumnCount(template: string): number {
  const tracks = template.trim().match(/(?:[^\s(]+|\([^)]*\))+/g);
  return Math.max(1, tracks?.length ?? 1);
}

function finiteCssPixels(value: string): number | undefined {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function sameRange(left: ThumbnailRange, right: ThumbnailRange): boolean {
  return left.start === right.start && left.end === right.end;
}

function sameMetrics(left: ThumbnailMetrics, right: ThumbnailMetrics): boolean {
  return (
    left.columns === right.columns &&
    Math.abs(left.itemHeight - right.itemHeight) < 0.5 &&
    Math.abs(left.rowGap - right.rowGap) < 0.5 &&
    left.viewportRows === right.viewportRows
  );
}
