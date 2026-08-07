export const MIN_DOCUMENT_ZOOM = 50;
export const MAX_DOCUMENT_ZOOM = 200;

export type DocumentZoomFit = 'page' | 'width';

export interface DocumentZoomFitMetrics {
  pageHeight: number;
  pageWidth: number;
  viewportHeight: number;
  viewportWidth: number;
  viewportPadding: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };
}

export function clampDocumentZoom(zoom: number): number {
  return Math.min(
    MAX_DOCUMENT_ZOOM,
    Math.max(MIN_DOCUMENT_ZOOM, Math.round(zoom)),
  );
}

export function documentZoomForFit(
  fit: DocumentZoomFit,
  metrics: DocumentZoomFitMetrics,
): number {
  const availableWidth = Math.max(
    1,
    metrics.viewportWidth -
      metrics.viewportPadding.left -
      metrics.viewportPadding.right,
  );
  const widthZoom = (availableWidth / Math.max(1, metrics.pageWidth)) * 100;
  if (fit === 'width') return clampDocumentZoom(widthZoom);

  const availableHeight = Math.max(
    1,
    metrics.viewportHeight -
      metrics.viewportPadding.top -
      metrics.viewportPadding.bottom,
  );
  const heightZoom = (availableHeight / Math.max(1, metrics.pageHeight)) * 100;
  return clampDocumentZoom(Math.min(widthZoom, heightZoom));
}
