import { useEffect, useMemo, useRef, useState } from 'react';

export function usePresentationThumbnailVisibility(
  slideIds: readonly string[],
  viewMode: 'normal' | 'sorter',
) {
  const viewportRef = useRef<HTMLElement | null>(null);
  const slideIdsRef = useRef(slideIds);
  const [visibleIds, setVisibleIds] = useState<ReadonlySet<string>>(
    () => new Set(),
  );
  slideIdsRef.current = slideIds;
  const slideIdSignature = useMemo(
    () => slideIds.map((slideId) => `${slideId.length}:${slideId}`).join(''),
    [slideIds],
  );

  useEffect(() => {
    const viewport = viewportRef.current;
    const availableIds = new Set(slideIdsRef.current);
    setVisibleIds((current) => retainedSlideIds(current, availableIds));
    if (!viewport) return;
    if (typeof IntersectionObserver === 'undefined') {
      setVisibleIds(availableIds);
      return;
    }

    let active = true;
    const observer = new IntersectionObserver(
      (entries) => {
        if (!active) return;
        setVisibleIds((current) => updatedVisibleSlideIds(current, entries));
      },
      {
        root: viewport,
        rootMargin: viewMode === 'sorter' ? '480px 240px' : '360px 0px',
      },
    );
    const observed = new Set<HTMLElement>();
    const syncObservedThumbnails = () => {
      const mounted = new Set(
        viewport.querySelectorAll<HTMLElement>('[data-slide-thumbnail]'),
      );
      for (const thumbnail of observed) {
        if (mounted.has(thumbnail)) continue;
        observer.unobserve(thumbnail);
        observed.delete(thumbnail);
      }
      for (const thumbnail of mounted) {
        if (observed.has(thumbnail)) continue;
        observed.add(thumbnail);
        observer.observe(thumbnail);
      }
      const mountedIds = new Set(
        [...mounted]
          .map((thumbnail) => thumbnail.dataset.slideId)
          .filter((slideId): slideId is string => Boolean(slideId)),
      );
      setVisibleIds((current) => retainedSlideIds(current, mountedIds));
    };
    syncObservedThumbnails();
    const mutationObserver =
      typeof MutationObserver === 'undefined'
        ? null
        : new MutationObserver(syncObservedThumbnails);
    mutationObserver?.observe(viewport, { childList: true, subtree: true });
    return () => {
      active = false;
      mutationObserver?.disconnect();
      observer.disconnect();
      observed.clear();
    };
  }, [slideIdSignature, viewMode]);

  return { viewportRef, visibleIds };
}

function retainedSlideIds(
  current: ReadonlySet<string>,
  available: ReadonlySet<string>,
): ReadonlySet<string> {
  if ([...current].every((slideId) => available.has(slideId))) return current;
  return new Set([...current].filter((slideId) => available.has(slideId)));
}

function updatedVisibleSlideIds(
  current: ReadonlySet<string>,
  entries: readonly IntersectionObserverEntry[],
): ReadonlySet<string> {
  const next = new Set(current);
  let changed = false;
  for (const entry of entries) {
    const slideId =
      entry.target instanceof HTMLElement
        ? entry.target.dataset.slideId
        : undefined;
    if (!slideId) continue;
    if (entry.isIntersecting) {
      if (next.has(slideId)) continue;
      next.add(slideId);
      changed = true;
      continue;
    }
    changed = next.delete(slideId) || changed;
  }
  return changed ? next : current;
}
