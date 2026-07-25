import { useEffect, useRef, useState } from 'react';

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
  const slideIdSignature = slideIds
    .map((slideId) => `${slideId.length}:${slideId}`)
    .join('');

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
    for (const thumbnail of viewport.querySelectorAll<HTMLElement>(
      '[data-slide-thumbnail]',
    )) {
      observer.observe(thumbnail);
    }
    return () => {
      active = false;
      observer.disconnect();
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
