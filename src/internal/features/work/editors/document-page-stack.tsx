import { useLayoutEffect, useMemo, useRef, useState } from 'react';
import { WorkDocumentPageBorder } from '../components/work-document-page-border';
import { resolveDocumentPageBorders } from '../work-document-page-borders';
import { resolveDocumentPageMargins } from '../work-document-page-margins';
import { documentPageSurfaceGeometry } from '../work-document-page-frames';
import { registerDocumentPageSurfaceGeometry } from '../work-document-page-surface-registry';
import { resolveDocumentPageSize } from '../work-document-page-size';
import type { WorkDocumentSectionLayout } from '../work-types';
import type { OfficeKernelPageMetrics } from '../../../kernel/office-kernel-protocol';

interface DocumentPageStackProps {
  pageColor: string;
  pageCount: number;
  pageGap: number;
  pageHeight: number;
  pageWidth?: number;
  pages?: readonly {
    layout: WorkDocumentSectionLayout;
    page?: OfficeKernelPageMetrics;
    physicalPage?: number;
    sectionPage: number;
  }[];
}

const EMPTY_DOCUMENT_PAGE_DESCRIPTORS: NonNullable<
  DocumentPageStackProps['pages']
> = [];

interface DocumentPageWindow {
  end: number;
  start: number;
}

const DOCUMENT_PAGE_WINDOW_THRESHOLD = 24;
const DOCUMENT_PAGE_INITIAL_WINDOW_SIZE = 8;
const DOCUMENT_PAGE_WINDOW_OVERSCAN = 3;

export function DocumentPageStack({
  pageColor,
  pageCount,
  pageGap,
  pageHeight,
  pageWidth = 794,
  pages = EMPTY_DOCUMENT_PAGE_DESCRIPTORS,
}: DocumentPageStackProps) {
  const count = Math.max(1, Math.trunc(pageCount));
  const gap = Math.max(0, pageGap);
  const height = Math.max(1, pageHeight);
  const fallbackPage = useMemo<OfficeKernelPageMetrics>(
    () => ({
      width: Math.max(1, pageWidth),
      height,
      marginTop: 0,
      marginRight: 0,
      marginBottom: 0,
      marginLeft: 0,
      headerHeight: 0,
      footerHeight: 0,
      pageGap: gap,
    }),
    [gap, height, pageWidth],
  );
  const geometry = useMemo(
    () =>
      documentPageSurfaceGeometry(
        pages.map((page) => page.page ?? fallbackPage),
        fallbackPage,
        count,
      ),
    [count, fallbackPage, pages],
  );
  const windowed = count > DOCUMENT_PAGE_WINDOW_THRESHOLD;
  const stackRef = useRef<HTMLDivElement>(null);
  const [requestedWindow, setRequestedWindow] = useState<DocumentPageWindow>(
    () => initialDocumentPageWindow(count),
  );
  const pageWindow = windowed
    ? clampDocumentPageWindow(requestedWindow, count)
    : { start: 0, end: count };
  const descriptors = geometry.frames
    .slice(pageWindow.start, pageWindow.end)
    .map((frame) => ({
      frame,
      page: pages[frame.pageIndex],
    }));

  useLayoutEffect(() => {
    const stack = stackRef.current;
    if (!stack) return;
    return registerDocumentPageSurfaceGeometry(stack, () => {
      if (pages.length < count) return [];
      return geometry.frames.map((frame) => {
        const descriptor = pages[frame.pageIndex];
        const pageSize = descriptor
          ? resolveDocumentPageSize(descriptor.layout)
          : null;
        return {
          height: frame.height,
          left: frame.left,
          orientation: pageSize?.orientation ?? 'portrait',
          pageHeightPoints: pageSize?.heightPoints ?? (frame.height * 72) / 96,
          pageIndex: frame.pageIndex,
          pageSize:
            pageSize?.preset && pageSize.preset !== 'custom'
              ? pageSize.preset
              : 'a4',
          pageWidthPoints: pageSize?.widthPoints ?? (frame.width * 72) / 96,
          top: frame.top,
          width: frame.width,
        };
      });
    });
  }, [count, geometry.frames, pages]);

  useLayoutEffect(() => {
    const stack = stackRef.current;
    if (!stack || !windowed) {
      return;
    }
    const root = stack.closest<HTMLElement>('.work-document-scroll');
    if (!root) {
      const next = initialDocumentPageWindow(count);
      setRequestedWindow((current) =>
        documentPageWindowsEqual(current, next) ? current : next,
      );
      return;
    }
    let animationFrame = 0;
    const update = () => {
      animationFrame = 0;
      const next = documentPageWindowForViewport(geometry.frames, stack, root);
      setRequestedWindow((current) =>
        documentPageWindowsEqual(current, next) ? current : next,
      );
    };
    const schedule = () => {
      if (animationFrame) return;
      animationFrame = requestAnimationFrame(update);
    };
    const resizeObserver =
      typeof ResizeObserver === 'undefined'
        ? null
        : new ResizeObserver(schedule);
    root.addEventListener('scroll', schedule, { passive: true });
    globalThis.addEventListener('resize', schedule);
    resizeObserver?.observe(root);
    resizeObserver?.observe(stack);
    update();
    return () => {
      if (animationFrame) cancelAnimationFrame(animationFrame);
      resizeObserver?.disconnect();
      root.removeEventListener('scroll', schedule);
      globalThis.removeEventListener('resize', schedule);
    };
  }, [count, geometry.frames, windowed]);

  return (
    <>
      <div
        ref={stackRef}
        className="work-document-page-stack"
        data-page-count={count}
        data-page-surface-height={geometry.height}
        data-page-surface-width={geometry.width}
        data-page-window-end={pageWindow.end}
        data-page-window-start={pageWindow.start + 1}
        data-page-windowed={String(windowed)}
        aria-hidden="true"
      >
        {descriptors.map(({ frame, page }) => {
          const resolvedPageSize = page
            ? resolveDocumentPageSize(page.layout)
            : null;
          return (
            <div
              className="work-document-page-sheet"
              data-work-document-page-sheet=""
              data-page-index={frame.pageIndex + 1}
              data-page-top={frame.top}
              data-page-left={frame.left}
              data-page-width={frame.width}
              data-page-height={frame.height}
              data-pdf-orientation={resolvedPageSize?.orientation}
              data-pdf-page-size={resolvedPageSize?.preset}
              data-pdf-page-width-points={
                resolvedPageSize?.widthPoints ?? (frame.width * 72) / 96
              }
              data-pdf-page-height-points={
                resolvedPageSize?.heightPoints ?? (frame.height * 72) / 96
              }
              key={frame.pageIndex}
              style={{
                backgroundColor: pageColor,
                height: frame.height,
                left: frame.left,
                top: frame.top,
                width: frame.width,
              }}
            />
          );
        })}
      </div>
      {(['back', 'front'] as const).map((zOrder) => (
        <div
          className={`work-document-page-border-stack ${zOrder}`}
          data-document-page-border-stack={zOrder}
          aria-hidden="true"
          key={zOrder}
        >
          {descriptors.flatMap(({ frame, page }) => {
            const resolved = page
              ? resolveDocumentPageBorders(
                  page.layout.pageBorders,
                  resolveDocumentPageMargins(
                    page.layout,
                    page.physicalPage ?? frame.pageIndex + 1,
                  ).body,
                )
              : null;
            if (!page || resolved?.zOrder !== zOrder) return [];
            return [
              <div
                className="work-document-page-border-surface"
                data-page-index={frame.pageIndex + 1}
                key={frame.pageIndex}
                style={{
                  height: frame.height,
                  left: frame.left,
                  top: frame.top,
                  width: frame.width,
                }}
              >
                <WorkDocumentPageBorder
                  layout={page.layout}
                  physicalPage={page.physicalPage ?? frame.pageIndex + 1}
                  sectionPage={page.sectionPage}
                />
              </div>,
            ];
          })}
        </div>
      ))}
    </>
  );
}

function initialDocumentPageWindow(pageCount: number): DocumentPageWindow {
  return {
    start: 0,
    end: Math.min(pageCount, DOCUMENT_PAGE_INITIAL_WINDOW_SIZE),
  };
}

function clampDocumentPageWindow(
  pageWindow: DocumentPageWindow,
  pageCount: number,
): DocumentPageWindow {
  const start = Math.max(0, Math.min(pageCount - 1, pageWindow.start));
  const end = Math.max(start + 1, Math.min(pageCount, pageWindow.end));
  return { start, end };
}

function documentPageWindowForViewport(
  frames: readonly { height: number; top: number }[],
  stack: HTMLElement,
  root: HTMLElement,
): DocumentPageWindow {
  if (!frames.length) return { start: 0, end: 1 };
  const stackRect = stack.getBoundingClientRect();
  const rootRect = root.getBoundingClientRect();
  if (!(stackRect.width > 0) || !(rootRect.height > 0)) {
    return initialDocumentPageWindow(frames.length);
  }
  const scale = Math.max(
    0.01,
    stackRect.width / Math.max(1, stack.offsetWidth),
  );
  const viewportTop = Math.max(0, (rootRect.top - stackRect.top) / scale);
  const viewportBottom = Math.max(
    viewportTop,
    (rootRect.bottom - stackRect.top) / scale,
  );
  const firstVisible = firstDocumentPageEndingAfter(frames, viewportTop);
  const lastVisible = lastDocumentPageStartingBefore(frames, viewportBottom);
  return {
    start: Math.max(0, firstVisible - DOCUMENT_PAGE_WINDOW_OVERSCAN),
    end: Math.min(
      frames.length,
      Math.max(firstVisible + 1, lastVisible + 1) +
        DOCUMENT_PAGE_WINDOW_OVERSCAN,
    ),
  };
}

function firstDocumentPageEndingAfter(
  frames: readonly { height: number; top: number }[],
  offset: number,
): number {
  let lower = 0;
  let upper = frames.length - 1;
  let result = upper;
  while (lower <= upper) {
    const middle = Math.floor((lower + upper) / 2);
    const frame = frames[middle];
    if (frame && frame.top + frame.height >= offset) {
      result = middle;
      upper = middle - 1;
    } else {
      lower = middle + 1;
    }
  }
  return result;
}

function lastDocumentPageStartingBefore(
  frames: readonly { top: number }[],
  offset: number,
): number {
  let lower = 0;
  let upper = frames.length - 1;
  let result = 0;
  while (lower <= upper) {
    const middle = Math.floor((lower + upper) / 2);
    const frame = frames[middle];
    if (frame && frame.top <= offset) {
      result = middle;
      lower = middle + 1;
    } else {
      upper = middle - 1;
    }
  }
  return result;
}

function documentPageWindowsEqual(
  left: DocumentPageWindow,
  right: DocumentPageWindow,
): boolean {
  return left.start === right.start && left.end === right.end;
}
