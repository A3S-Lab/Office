import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';

export const DOCUMENT_COMMENT_WINDOW_THRESHOLD = 48;
export const DOCUMENT_COMMENT_WINDOW_LIMIT = 32;
export const DOCUMENT_COMMENT_WINDOW_OVERSCAN = 4;

export interface DocumentCommentWindowRange {
  end: number;
  start: number;
  windowed: boolean;
}

export function useDocumentCommentWindow({
  keys,
  onRovingKeyChange,
  pinnedKeys = [],
  rovingKey,
}: {
  keys: readonly string[];
  onRovingKeyChange: (key: string) => void;
  pinnedKeys?: readonly (string | null | undefined)[];
  rovingKey: string | null;
}) {
  const itemRefs = useRef(new Map<string, HTMLElement>());
  const keysRef = useRef(keys);
  const pendingFocusKeyRef = useRef<string | null>(null);
  const rovingIndex = rovingKey ? keys.indexOf(rovingKey) : -1;
  const [anchorIndex, setAnchorIndex] = useState(Math.max(0, rovingIndex));
  keysRef.current = keys;

  useEffect(() => {
    if (rovingIndex < 0) return;
    setAnchorIndex(rovingIndex);
  }, [rovingIndex, rovingKey]);

  const range = useMemo(
    () =>
      calculateDocumentCommentWindowRange({
        anchorIndex,
        itemCount: keys.length,
      }),
    [anchorIndex, keys.length],
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
      documentCommentWindowIndices({
        itemCount: keys.length,
        pinnedIndices,
        range,
      }),
    [keys.length, pinnedIndices, range],
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
        focusDocumentCommentItem(mounted);
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
  const onViewportAnchorChange = useCallback((index: number) => {
    setAnchorIndex((current) => (current === index ? current : index));
  }, []);

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
    focusDocumentCommentItem(element);
  }, [keys, mountedIndicesKey, rovingKey]);

  return {
    focusAt,
    mountedIndices,
    onItemFocus,
    onViewportAnchorChange,
    range,
    registerItem,
    rovingIndex,
  };
}

export function calculateDocumentCommentWindowRange({
  anchorIndex,
  itemCount,
  limit = DOCUMENT_COMMENT_WINDOW_LIMIT,
  overscan = DOCUMENT_COMMENT_WINDOW_OVERSCAN,
  threshold = DOCUMENT_COMMENT_WINDOW_THRESHOLD,
}: {
  anchorIndex: number;
  itemCount: number;
  limit?: number;
  overscan?: number;
  threshold?: number;
}): DocumentCommentWindowRange {
  const count = Math.max(0, Math.floor(itemCount));
  if (count <= Math.max(0, Math.floor(threshold))) {
    return { end: count, start: 0, windowed: false };
  }
  const windowCount = Math.min(count, Math.max(1, Math.floor(limit)));
  const boundedAnchor = Math.min(
    count - 1,
    Math.max(0, Math.floor(anchorIndex)),
  );
  const start = Math.min(
    count - windowCount,
    Math.max(0, boundedAnchor - Math.max(0, Math.floor(overscan))),
  );
  return { end: start + windowCount, start, windowed: true };
}

export function documentCommentWindowIndices({
  itemCount,
  pinnedIndices,
  range,
}: {
  itemCount: number;
  pinnedIndices: readonly number[];
  range: DocumentCommentWindowRange;
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

export function documentCommentKeyboardDestination(
  key: string,
  index: number,
  itemCount: number,
): number | null {
  if (!itemCount) return null;
  if (key === 'ArrowDown') return Math.min(itemCount - 1, index + 1);
  if (key === 'ArrowUp') return Math.max(0, index - 1);
  if (key === 'PageDown') return Math.min(itemCount - 1, index + 8);
  if (key === 'PageUp') return Math.max(0, index - 8);
  if (key === 'Home') return 0;
  if (key === 'End') return itemCount - 1;
  return null;
}

export function focusDocumentCommentItem(element: HTMLElement): void {
  element.focus({ preventScroll: true });
  scrollDocumentCommentItemIntoView(element);
}

export function scrollDocumentCommentItemIntoView(element: HTMLElement): void {
  const panel = element.closest<HTMLElement>('.work-document-comments-panel');
  const view = element.ownerDocument.defaultView;
  const panelOverflow = panel
    ? view?.getComputedStyle(panel).overflowY
    : undefined;
  const panelOwnsScroll =
    Boolean(panel) &&
    /^(auto|scroll)$/.test(panelOverflow ?? '') &&
    (panel?.scrollHeight ?? 0) > (panel?.clientHeight ?? 0) + 1;
  const scrollContainer = panelOwnsScroll
    ? panel
    : panel?.closest<HTMLElement>('.work-document-scroll');
  if (!scrollContainer) {
    element.scrollIntoView?.({ block: 'nearest' });
    return;
  }

  const scrollTarget =
    element.closest<HTMLElement>('.work-document-comment-card') ?? element;
  const elementRect = scrollTarget.getBoundingClientRect();
  const containerRect = scrollContainer.getBoundingClientRect();
  const header = panel?.querySelector<HTMLElement>(':scope > header');
  const headerHeight = header?.getBoundingClientRect().height ?? 0;
  const visibleTop = containerRect.top + headerHeight + 8;
  const visibleBottom = containerRect.bottom - 8;
  let nextScrollTop: number | null = null;
  if (elementRect.top < visibleTop) {
    nextScrollTop = scrollContainer.scrollTop + elementRect.top - visibleTop;
  } else if (elementRect.bottom > visibleBottom) {
    nextScrollTop =
      scrollContainer.scrollTop + elementRect.bottom - visibleBottom;
  }
  if (nextScrollTop === null) return;

  const top = Math.max(0, nextScrollTop);
  if (typeof scrollContainer.scrollTo === 'function') {
    scrollContainer.scrollTo({ behavior: 'instant', top });
    return;
  }

  const previousScrollBehavior = scrollContainer.style.scrollBehavior;
  scrollContainer.style.scrollBehavior = 'auto';
  void view?.getComputedStyle(scrollContainer).scrollBehavior;
  scrollContainer.scrollTop = top;
  scrollContainer.style.scrollBehavior = previousScrollBehavior;
}
