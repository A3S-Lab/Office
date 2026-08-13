import type { OfficeKernelPageMetrics } from '../../kernel/office-kernel-protocol';

export interface DocumentPageFrame {
  height: number;
  left: number;
  page: OfficeKernelPageMetrics;
  pageIndex: number;
  top: number;
  width: number;
}

export interface DocumentPageSurfaceGeometry {
  frames: DocumentPageFrame[];
  height: number;
  width: number;
}

export function documentPageSurfaceGeometry(
  pages: readonly OfficeKernelPageMetrics[],
  fallback: OfficeKernelPageMetrics,
  requestedPageCount = pages.length,
): DocumentPageSurfaceGeometry {
  const pageCount = Math.max(1, Math.trunc(requestedPageCount));
  const metrics = Array.from(
    { length: pageCount },
    (_, pageIndex) => pages[pageIndex] ?? fallback,
  );
  const width = Math.max(1, ...metrics.map((page) => page.width));
  let top = 0;
  const frames = metrics.map((page, pageIndex): DocumentPageFrame => {
    const frame = {
      height: page.height,
      left: (width - page.width) / 2,
      page,
      pageIndex,
      top,
      width: page.width,
    };
    top += page.height;
    if (pageIndex + 1 < metrics.length) top += page.pageGap;
    return frame;
  });
  return {
    frames,
    height: Math.max(1, top),
    width,
  };
}
