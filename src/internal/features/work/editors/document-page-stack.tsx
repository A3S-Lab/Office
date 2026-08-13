import { WorkDocumentPageBorder } from '../components/work-document-page-border';
import { resolveDocumentPageBorders } from '../work-document-page-borders';
import type { WorkDocumentSectionLayout } from '../work-types';

interface DocumentPageStackProps {
  pageColor: string;
  pageCount: number;
  pageGap: number;
  pageHeight: number;
  pages?: readonly {
    layout: WorkDocumentSectionLayout;
    sectionPage: number;
  }[];
}

export function DocumentPageStack({
  pageColor,
  pageCount,
  pageGap,
  pageHeight,
  pages = [],
}: DocumentPageStackProps) {
  const count = Math.max(1, Math.trunc(pageCount));
  const gap = Math.max(0, pageGap);
  const height = Math.max(1, pageHeight);

  const descriptors = Array.from({ length: count }, (_, pageIndex) => ({
    page: pages[pageIndex],
    pageIndex,
  }));
  return (
    <>
      <div
        className="work-document-page-stack"
        data-page-count={count}
        aria-hidden="true"
      >
        {descriptors.map(({ pageIndex }) => (
          <div
            className="work-document-page-sheet"
            data-page-index={pageIndex + 1}
            key={pageIndex}
            style={{
              backgroundColor: pageColor,
              height,
              top: pageIndex * (height + gap),
            }}
          />
        ))}
      </div>
      {(['back', 'front'] as const).map((zOrder) => (
        <div
          className={`work-document-page-border-stack ${zOrder}`}
          data-document-page-border-stack={zOrder}
          aria-hidden="true"
          key={zOrder}
        >
          {descriptors.flatMap(({ page, pageIndex }) => {
            const resolved = page
              ? resolveDocumentPageBorders(
                  page.layout.pageBorders,
                  page.layout.margins,
                )
              : null;
            if (!page || resolved?.zOrder !== zOrder) return [];
            return [
              <div
                className="work-document-page-border-surface"
                data-page-index={pageIndex + 1}
                key={pageIndex}
                style={{
                  height,
                  top: pageIndex * (height + gap),
                }}
              >
                <WorkDocumentPageBorder
                  layout={page.layout}
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
