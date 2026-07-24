import { resolveDocumentPageChrome } from '../work-document-page-chrome';
import type { documentInitialSectionLayout } from '../work-document-section';
import type { DocumentPaginationPageDescriptor } from './use-document-pagination';

export function fallbackPaginationPageDescriptor(
  sectionId: string | undefined,
  sectionIndex: number | undefined,
  layout: ReturnType<typeof documentInitialSectionLayout>,
  physicalPage: number,
): DocumentPaginationPageDescriptor {
  const sectionPage = Math.max(1, physicalPage);
  return {
    pageIndex: sectionPage - 1,
    physicalPage: sectionPage,
    pageNumber: Math.max(1, layout.pageNumberStart ?? 1) + sectionPage - 1,
    sectionPage,
    sectionId: sectionId ?? 'document-section-1',
    sectionIndex: sectionIndex ?? 0,
    layout,
    pageChrome: resolveDocumentPageChrome(layout, sectionPage, sectionPage),
  };
}
