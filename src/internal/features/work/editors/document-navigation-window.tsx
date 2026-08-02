import {
  type UIEventHandler,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';

export const DOCUMENT_NAVIGATION_COLLECTION_WINDOW_THRESHOLD = 48;
export const DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT = 32;
export const DOCUMENT_NAVIGATION_COLLECTION_WINDOW_OVERSCAN = 4;
const DOCUMENT_NAVIGATION_FALLBACK_VIEWPORT_HEIGHT = 720;

export interface DocumentNavigationWindowRange {
  end: number;
  start: number;
  windowed: boolean;
}

export interface DocumentNavigationWindowSpacerEntry {
  end: number;
  height: number;
  kind: 'spacer';
  position: 'after' | 'before' | 'between';
  start: number;
}

export interface DocumentNavigationWindowItemEntry {
  index: number;
  kind: 'item';
}

export type DocumentNavigationWindowEntry =
  | DocumentNavigationWindowItemEntry
  | DocumentNavigationWindowSpacerEntry;

export interface DocumentNavigationListHandle {
  focusFirst: () => void;
  focusLast: () => void;
}

export function useDocumentNavigationWindow({
  estimatedItemHeight,
  itemGap,
  keys,
  listPaddingTop,
  onRovingKeyChange,
  pinnedKeys = [],
  rovingKey,
}: {
  estimatedItemHeight: number;
  itemGap: number;
  keys: readonly string[];
  listPaddingTop: number;
  onRovingKeyChange: (key: string) => void;
  pinnedKeys?: readonly (string | null | undefined)[];
  rovingKey: string | null;
}) {
  const viewportRef = useRef<HTMLElement>(null);
  const itemRefs = useRef(new Map<string, HTMLElement>());
  const keysRef = useRef(keys);
  const pendingFocusKeyRef = useRef<string | null>(null);
  const rovingIndex = rovingKey ? keys.indexOf(rovingKey) : -1;
  const [anchorIndex, setAnchorIndex] = useState(Math.max(0, rovingIndex));
  const [viewportHeight, setViewportHeight] = useState(
    DOCUMENT_NAVIGATION_FALLBACK_VIEWPORT_HEIGHT,
  );
  keysRef.current = keys;

  useEffect(() => {
    if (rovingIndex < 0) return;
    setAnchorIndex(rovingIndex);
  }, [rovingIndex, rovingKey]);

  useLayoutEffect(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;
    const update = () => {
      const nextHeight =
        viewport.clientHeight || DOCUMENT_NAVIGATION_FALLBACK_VIEWPORT_HEIGHT;
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
      Array.from({ length: keys.length }, () =>
        Math.max(1, estimatedItemHeight),
      ),
    [estimatedItemHeight, keys.length],
  );
  const itemStarts = useMemo(
    () => documentNavigationItemStarts(itemHeights, itemGap),
    [itemGap, itemHeights],
  );
  const averageItemHeight = itemHeights.length
    ? (itemStarts[itemStarts.length - 1] ?? 1) / itemHeights.length
    : 1;
  const range = useMemo(
    () =>
      calculateDocumentNavigationWindowRange({
        anchorIndex,
        averageItemHeight,
        itemCount: keys.length,
        limit: DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT,
        overscan: DOCUMENT_NAVIGATION_COLLECTION_WINDOW_OVERSCAN,
        threshold: DOCUMENT_NAVIGATION_COLLECTION_WINDOW_THRESHOLD,
        viewportHeight,
      }),
    [anchorIndex, averageItemHeight, keys.length, viewportHeight],
  );
  const pinnedIndices = useMemo(() => {
    const indices = new Set<number>();
    if (rovingIndex >= 0) indices.add(rovingIndex);
    for (const key of pinnedKeys) {
      if (!key) continue;
      const index = keys.indexOf(key);
      if (index >= 0) indices.add(index);
    }
    return [...indices];
  }, [keys, pinnedKeys, rovingIndex]);
  const mountedIndices = useMemo(
    () =>
      documentNavigationWindowIndices({
        itemCount: keys.length,
        pinnedIndices,
        range,
      }),
    [keys.length, pinnedIndices, range],
  );
  const entries = useMemo(
    () =>
      documentNavigationWindowEntries(
        mountedIndices,
        itemStarts,
        keys.length,
        itemGap,
      ),
    [itemGap, itemStarts, keys.length, mountedIndices],
  );
  const mountedIndicesKey = mountedIndices.join(',');

  const registerItem = useCallback(
    (key: string, element: HTMLElement | null) => {
      if (element) itemRefs.current.set(key, element);
      else itemRefs.current.delete(key);
    },
    [],
  );

  const focusAt = useCallback(
    (requestedIndex: number) => {
      const currentKeys = keysRef.current;
      if (!currentKeys.length) return;
      const index = Math.min(
        currentKeys.length - 1,
        Math.max(0, Math.floor(requestedIndex)),
      );
      const key = currentKeys[index];
      if (!key) return;
      setAnchorIndex(index);
      onRovingKeyChange(key);
      const mounted = itemRefs.current.get(key);
      if (mounted) {
        pendingFocusKeyRef.current = null;
        focusDocumentNavigationItem(mounted);
        return;
      }
      pendingFocusKeyRef.current = key;
    },
    [onRovingKeyChange],
  );

  const onItemFocus = useCallback(
    (index: number) => {
      const key = keysRef.current[index];
      if (!key) return;
      pendingFocusKeyRef.current = null;
      setAnchorIndex(index);
      onRovingKeyChange(key);
    },
    [onRovingKeyChange],
  );

  useLayoutEffect(() => {
    const pendingKey = pendingFocusKeyRef.current;
    if (!pendingKey) return;
    if (!keys.includes(pendingKey)) {
      pendingFocusKeyRef.current = null;
      return;
    }
    const element = itemRefs.current.get(pendingKey);
    if (!element) return;
    pendingFocusKeyRef.current = null;
    focusDocumentNavigationItem(element);
  }, [keys, mountedIndicesKey, rovingKey]);

  const onScroll: UIEventHandler<HTMLElement> = useCallback(
    (event) => {
      if (!range.windowed) return;
      const nextAnchor = documentNavigationIndexAtOffset(
        itemStarts,
        Math.max(0, event.currentTarget.scrollTop - listPaddingTop),
        keys.length,
      );
      setAnchorIndex((current) =>
        current === nextAnchor ? current : nextAnchor,
      );
    },
    [itemStarts, keys.length, listPaddingTop, range.windowed],
  );

  return {
    entries,
    focusAt,
    mountedCount: mountedIndices.length,
    onItemFocus,
    onScroll,
    range,
    registerItem,
    rovingIndex,
    viewportRef,
  };
}

export function DocumentNavigationWindowSpacer({
  entry,
}: {
  entry: DocumentNavigationWindowSpacerEntry;
}) {
  return (
    <li
      aria-hidden="true"
      className="work-document-navigation-window-spacer"
      data-document-navigation-spacer={entry.position}
      data-document-navigation-spacer-end={entry.end}
      data-document-navigation-spacer-start={entry.start + 1}
      role="presentation"
      style={{ height: `${entry.height}px` }}
    />
  );
}

export function calculateDocumentNavigationWindowRange({
  anchorIndex,
  averageItemHeight,
  itemCount,
  limit,
  overscan,
  threshold,
  viewportHeight,
}: {
  anchorIndex: number;
  averageItemHeight: number;
  itemCount: number;
  limit: number;
  overscan: number;
  threshold: number;
  viewportHeight: number;
}): DocumentNavigationWindowRange {
  const count = Math.max(0, Math.floor(itemCount));
  if (count <= Math.max(0, Math.floor(threshold))) {
    return { end: count, start: 0, windowed: false };
  }
  const visibleCount = Math.max(
    1,
    Math.ceil(Math.max(1, viewportHeight) / Math.max(1, averageItemHeight)),
  );
  const boundedOverscan = Math.max(0, Math.floor(overscan));
  const windowCount = Math.min(
    count,
    Math.max(1, Math.floor(limit)),
    visibleCount + boundedOverscan * 2,
  );
  const boundedAnchor = Math.min(
    count - 1,
    Math.max(0, Math.floor(anchorIndex)),
  );
  const start = Math.min(
    count - windowCount,
    Math.max(0, boundedAnchor - boundedOverscan),
  );
  return { end: start + windowCount, start, windowed: true };
}

export function documentNavigationItemStarts(
  itemHeights: readonly number[],
  itemGap: number,
): number[] {
  const starts = [0];
  const gap = Math.max(0, itemGap);
  for (const height of itemHeights) {
    starts.push((starts[starts.length - 1] ?? 0) + Math.max(1, height) + gap);
  }
  return starts;
}

export function documentNavigationWindowIndices({
  itemCount,
  pinnedIndices,
  range,
}: {
  itemCount: number;
  pinnedIndices: readonly number[];
  range: DocumentNavigationWindowRange;
}): number[] {
  const count = Math.max(0, Math.floor(itemCount));
  if (!range.windowed) {
    return Array.from({ length: count }, (_, index) => index);
  }
  const mounted = new Set<number>();
  for (let index = range.start; index < range.end; index += 1) {
    if (index >= 0 && index < count) mounted.add(index);
  }
  for (const index of pinnedIndices) {
    if (index >= 0 && index < count) mounted.add(index);
  }
  return [...mounted].sort((left, right) => left - right);
}

export function documentNavigationWindowEntries(
  mountedIndices: readonly number[],
  itemStarts: readonly number[],
  itemCount: number,
  itemGap: number,
): DocumentNavigationWindowEntry[] {
  const entries: DocumentNavigationWindowEntry[] = [];
  let cursor = 0;
  for (const index of mountedIndices) {
    if (index > cursor) {
      entries.push(
        documentNavigationWindowSpacerEntry(
          cursor,
          index,
          itemStarts,
          itemCount,
          itemGap,
        ),
      );
    }
    entries.push({ index, kind: 'item' });
    cursor = index + 1;
  }
  if (cursor < itemCount) {
    entries.push(
      documentNavigationWindowSpacerEntry(
        cursor,
        itemCount,
        itemStarts,
        itemCount,
        itemGap,
      ),
    );
  }
  return entries;
}

export function documentNavigationIndexAtOffset(
  itemStarts: readonly number[],
  offset: number,
  itemCount: number,
): number {
  if (itemCount <= 1) return 0;
  let low = 0;
  let high = itemCount - 1;
  while (low <= high) {
    const middle = Math.floor((low + high) / 2);
    const start = itemStarts[middle] ?? 0;
    const next = itemStarts[middle + 1] ?? Number.POSITIVE_INFINITY;
    if (offset < start) high = middle - 1;
    else if (offset >= next) low = middle + 1;
    else return middle;
  }
  return Math.min(itemCount - 1, Math.max(0, low));
}

function documentNavigationWindowSpacerEntry(
  start: number,
  end: number,
  itemStarts: readonly number[],
  itemCount: number,
  itemGap: number,
): DocumentNavigationWindowSpacerEntry {
  return {
    end,
    height: Math.max(
      0,
      (itemStarts[end] ?? 0) - (itemStarts[start] ?? 0) - Math.max(0, itemGap),
    ),
    kind: 'spacer',
    position: start === 0 ? 'before' : end === itemCount ? 'after' : 'between',
    start,
  };
}

function focusDocumentNavigationItem(element: HTMLElement): void {
  element.focus({ preventScroll: true });
  element.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
}
