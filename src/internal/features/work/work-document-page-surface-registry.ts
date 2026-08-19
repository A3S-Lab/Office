import type { WorkSpreadsheetPaperSize } from './work-types';

export interface RegisteredDocumentPageSurfaceFrame {
  height: number;
  left: number;
  orientation: 'landscape' | 'portrait';
  pageHeightPoints: number;
  pageIndex: number;
  pageSize: WorkSpreadsheetPaperSize;
  pageWidthPoints: number;
  top: number;
  width: number;
}

type DocumentPageSurfaceProvider =
  () => readonly RegisteredDocumentPageSurfaceFrame[];

const documentPageSurfaceProviders = new WeakMap<
  HTMLElement,
  DocumentPageSurfaceProvider
>();

export function registerDocumentPageSurfaceGeometry(
  element: HTMLElement,
  provider: DocumentPageSurfaceProvider,
): () => void {
  documentPageSurfaceProviders.set(element, provider);
  return () => {
    if (documentPageSurfaceProviders.get(element) === provider) {
      documentPageSurfaceProviders.delete(element);
    }
  };
}

export function documentPageSurfaceGeometryForElement(
  element: HTMLElement,
): readonly RegisteredDocumentPageSurfaceFrame[] | null {
  return documentPageSurfaceProviders.get(element)?.() ?? null;
}
