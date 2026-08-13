import { WorkDocumentPageBorder } from '../components/work-document-page-border';
import { resolveDocumentPageBorders } from '../work-document-page-borders';
import { resolveDocumentPageMargins } from '../work-document-page-margins';
import { documentPageSurfaceGeometry } from '../work-document-page-frames';
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

export function DocumentPageStack({
  pageColor,
  pageCount,
  pageGap,
  pageHeight,
  pageWidth = 794,
  pages = [],
}: DocumentPageStackProps) {
  const count = Math.max(1, Math.trunc(pageCount));
  const gap = Math.max(0, pageGap);
  const height = Math.max(1, pageHeight);
  const fallbackPage: OfficeKernelPageMetrics = {
    width: Math.max(1, pageWidth),
    height,
    marginTop: 0,
    marginRight: 0,
    marginBottom: 0,
    marginLeft: 0,
    headerHeight: 0,
    footerHeight: 0,
    pageGap: gap,
  };
  const geometry = documentPageSurfaceGeometry(
    pages.map((page) => page.page ?? fallbackPage),
    fallbackPage,
    count,
  );
  const descriptors = geometry.frames.map((frame) => ({
    frame,
    page: pages[frame.pageIndex],
  }));
  return (
    <>
      <div
        className="work-document-page-stack"
        data-page-count={count}
        data-page-surface-height={geometry.height}
        data-page-surface-width={geometry.width}
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
