import {
  DocumentManagerPlugin,
  type DocumentManagerCapability,
  type PluginRegistry,
  ScrollPlugin,
  type ScrollCapability,
  ViewportPlugin,
  type ViewportCapability,
  ZoomPlugin,
  type ZoomCapability,
} from '@embedpdf/react-pdf-viewer';
import { useEffect, useMemo, useRef, useState } from 'react';

export const PDF_EVIDENCE_COORDINATE_BASIS = 1_000_000 as const;

export interface PdfEvidenceBounds {
  bottom: number;
  left: number;
  right: number;
  top: number;
}

export interface PdfEvidenceRegion {
  bounds: PdfEvidenceBounds;
  id: string;
  label?: string;
  sourceRegionIds: readonly string[];
  targetIds: readonly string[];
}

export interface PdfEvidencePage {
  canvasHeight: number;
  canvasId: string;
  canvasKind: 'source';
  canvasWidth: number;
  coordinateBasis: typeof PDF_EVIDENCE_COORDINATE_BASIS;
  pageNumber: number;
  regions: readonly PdfEvidenceRegion[];
  renderProfileSha256: string;
  rotationDegrees: 0 | 90 | 180 | 270;
  sourceSha256: string;
}

export interface PdfEvidenceRegionLocation {
  pageNumber: number;
}

export interface PdfEvidenceOverlay {
  coordinateBasis: typeof PDF_EVIDENCE_COORDINATE_BASIS;
  loadPage: (
    pageNumber: number,
    signal: AbortSignal,
  ) => Promise<PdfEvidencePage | null>;
  locateRegion?: (
    regionId: string,
    signal: AbortSignal,
  ) => Promise<PdfEvidenceRegionLocation | null>;
  renderProfileSha256: string;
  sourceSha256: string;
}

interface PdfEvidenceOverlayLayerProps {
  currentPage: number;
  evidenceOverlay: PdfEvidenceOverlay;
  onEvidenceRegionSelect?: (region: PdfEvidenceRegion) => void;
  registry: PluginRegistry;
  selectedEvidenceRegionId?: string;
  sourceKey?: string;
}

interface EvidenceRuntime {
  documentManager: DocumentManagerCapability;
  scroll: ScrollCapability;
  viewport: ViewportCapability;
  zoom: ZoomCapability | null;
}

type PageLoadState =
  | { status: 'idle' | 'loading' }
  | { message: string; status: 'error' }
  | { page: PdfEvidencePage | null; status: 'ready' };

const MAX_EVIDENCE_REGIONS_PER_PAGE = 2_000;
const MAX_REGION_TARGETS = 128;
const MAX_ID_LENGTH = 512;
const PAGE_ASPECT_RATIO_TOLERANCE = 0.005;

export function PdfEvidenceOverlayLayer({
  currentPage,
  evidenceOverlay,
  onEvidenceRegionSelect,
  registry,
  selectedEvidenceRegionId,
  sourceKey,
}: PdfEvidenceOverlayLayerProps) {
  const [runtime, setRuntime] = useState<EvidenceRuntime | null>(null);
  const [pageState, setPageState] = useState<PageLoadState>({ status: 'idle' });
  const [viewportRevision, setViewportRevision] = useState(0);
  const lastNavigationRef = useRef<string | null>(null);

  useEffect(() => {
    let disposed = false;
    void registry
      .pluginsReady()
      .then(() => {
        if (!disposed) setRuntime(readEvidenceRuntime(registry));
      })
      .catch((error: unknown) => {
        if (!disposed) {
          setPageState({
            message: evidenceErrorMessage(error),
            status: 'error',
          });
        }
      });
    return () => {
      disposed = true;
    };
  }, [registry]);

  useEffect(() => {
    if (!runtime) return;
    let frame = 0;
    const update = () => {
      if (frame) return;
      frame = window.requestAnimationFrame(() => {
        frame = 0;
        setViewportRevision((value) => value + 1);
      });
    };
    const unsubscribe = [
      runtime.scroll.onScroll(() => update()),
      runtime.scroll.onLayoutChange(() => update()),
      runtime.viewport.onScrollChange(() => update()),
      runtime.viewport.onViewportResize(() => update()),
      runtime.zoom?.onStateChange(() => update()),
    ].filter((stop): stop is () => void => Boolean(stop));
    return () => {
      if (frame) window.cancelAnimationFrame(frame);
      for (const stop of unsubscribe) stop();
    };
  }, [runtime]);

  useEffect(() => {
    const abort = new AbortController();
    if (!runtime || currentPage < 1) {
      setPageState({ status: 'idle' });
      return () => abort.abort();
    }
    const identityError = validateOverlayIdentity(evidenceOverlay, sourceKey);
    if (identityError) {
      setPageState({ message: identityError, status: 'error' });
      return () => abort.abort();
    }
    setPageState({ status: 'loading' });
    void evidenceOverlay
      .loadPage(currentPage, abort.signal)
      .then((page) => {
        if (abort.signal.aborted) return;
        if (page) {
          validatePdfEvidencePage(page, evidenceOverlay, runtime, currentPage);
        }
        setPageState({ page, status: 'ready' });
      })
      .catch((error: unknown) => {
        if (!abort.signal.aborted) {
          setPageState({
            message: evidenceErrorMessage(error),
            status: 'error',
          });
        }
      });
    return () => abort.abort();
  }, [currentPage, evidenceOverlay, runtime, sourceKey]);

  const page = pageState.status === 'ready' ? pageState.page : null;
  const regions = useMemo(
    () =>
      page && runtime
        ? positionEvidenceRegions(page, runtime, viewportRevision)
        : [],
    [page, runtime, viewportRevision],
  );

  useEffect(() => {
    if (!runtime || !page || !selectedEvidenceRegionId) return;
    const selected = page.regions.find(
      (region) => region.id === selectedEvidenceRegionId,
    );
    if (!selected) return;
    const navigationKey = `${page.pageNumber}:${selected.id}`;
    if (lastNavigationRef.current === navigationKey) return;
    lastNavigationRef.current = navigationKey;
    scrollEvidenceRegionIntoView(runtime, page, selected);
  }, [page, runtime, selectedEvidenceRegionId]);

  useEffect(() => {
    if (
      !runtime ||
      !selectedEvidenceRegionId ||
      !evidenceOverlay.locateRegion ||
      page?.regions.some((region) => region.id === selectedEvidenceRegionId)
    ) {
      return;
    }
    const abort = new AbortController();
    void evidenceOverlay
      .locateRegion(selectedEvidenceRegionId, abort.signal)
      .then((location) => {
        if (
          !abort.signal.aborted &&
          location &&
          location.pageNumber !== currentPage
        ) {
          scrollToEvidencePage(runtime, location.pageNumber);
        }
      })
      .catch(() => undefined);
    return () => abort.abort();
  }, [currentPage, evidenceOverlay, page, runtime, selectedEvidenceRegionId]);

  if (pageState.status === 'error') {
    return (
      <div
        className="work-pdf-evidence-status"
        role="status"
        title={pageState.message}
      >
        当前页定位不可用
      </div>
    );
  }
  if (pageState.status === 'ready' && !pageState.page) {
    return (
      <div className="work-pdf-evidence-status" role="status">
        当前页暂无安全定位
      </div>
    );
  }
  if (!page || regions.length === 0) return null;

  return (
    <fieldset className="work-pdf-evidence-overlay" aria-label="解析证据定位">
      {regions.map(({ region, style }) => {
        const selected = region.id === selectedEvidenceRegionId;
        return (
          <button
            key={region.id}
            type="button"
            className="work-pdf-evidence-region"
            data-selected={selected || undefined}
            style={style}
            aria-label={region.label ?? `解析区域 ${region.id}`}
            aria-pressed={selected}
            title={region.label}
            onClick={() => onEvidenceRegionSelect?.(region)}
          />
        );
      })}
    </fieldset>
  );
}

function readEvidenceRuntime(registry: PluginRegistry): EvidenceRuntime {
  const documentManager = registry
    .getPlugin<DocumentManagerPlugin>(DocumentManagerPlugin.id)
    ?.provides();
  const scroll = registry.getPlugin<ScrollPlugin>(ScrollPlugin.id)?.provides();
  const viewport = registry
    .getPlugin<ViewportPlugin>(ViewportPlugin.id)
    ?.provides();
  if (!documentManager || !scroll || !viewport) {
    throw new Error('Required PDF evidence capabilities are unavailable.');
  }
  return {
    documentManager,
    scroll,
    viewport,
    zoom: registry.getPlugin<ZoomPlugin>(ZoomPlugin.id)?.provides() ?? null,
  };
}

function validateOverlayIdentity(
  overlay: PdfEvidenceOverlay,
  sourceKey: string | undefined,
): string | null {
  if (!sourceKey || sourceKey !== overlay.sourceSha256) {
    return 'The evidence source identity does not match the PDF source.';
  }
  if (
    !isSha256(overlay.sourceSha256) ||
    !isSha256(overlay.renderProfileSha256)
  ) {
    return 'The evidence overlay contains an invalid SHA-256 identity.';
  }
  if (overlay.coordinateBasis !== PDF_EVIDENCE_COORDINATE_BASIS) {
    return 'The evidence overlay uses an unsupported coordinate basis.';
  }
  return null;
}

export function validatePdfEvidencePage(
  page: PdfEvidencePage,
  overlay: PdfEvidenceOverlay,
  runtime: EvidenceRuntime,
  requestedPage: number,
): void {
  if (
    page.pageNumber !== requestedPage ||
    page.sourceSha256 !== overlay.sourceSha256 ||
    page.renderProfileSha256 !== overlay.renderProfileSha256 ||
    page.coordinateBasis !== PDF_EVIDENCE_COORDINATE_BASIS ||
    page.canvasKind !== 'source'
  ) {
    throw new Error('The evidence page identity is invalid.');
  }
  if (
    !boundedId(page.canvasId) ||
    !positiveFinite(page.canvasWidth) ||
    !positiveFinite(page.canvasHeight)
  ) {
    throw new Error('The evidence canvas identity or dimensions are invalid.');
  }
  const documentPage =
    runtime.documentManager.getActiveDocument()?.pages[requestedPage - 1];
  if (!documentPage) throw new Error('The PDF page is unavailable.');
  const expectedRotation = documentPage.rotation * 90;
  if (page.rotationDegrees !== expectedRotation) {
    throw new Error(
      'The evidence canvas rotation does not match the PDF page.',
    );
  }
  const rotatedWidth =
    documentPage.rotation % 2
      ? documentPage.size.height
      : documentPage.size.width;
  const rotatedHeight =
    documentPage.rotation % 2
      ? documentPage.size.width
      : documentPage.size.height;
  const expectedRatio = rotatedWidth / rotatedHeight;
  const evidenceRatio = page.canvasWidth / page.canvasHeight;
  if (
    Math.abs(evidenceRatio / expectedRatio - 1) > PAGE_ASPECT_RATIO_TOLERANCE
  ) {
    throw new Error(
      'The evidence canvas aspect ratio does not match the PDF page.',
    );
  }
  if (page.regions.length > MAX_EVIDENCE_REGIONS_PER_PAGE) {
    throw new Error('The evidence page exceeds the bounded region limit.');
  }
  const ids = new Set<string>();
  for (const region of page.regions) {
    if (!boundedId(region.id) || ids.has(region.id)) {
      throw new Error(
        'The evidence page contains an invalid or duplicate region identity.',
      );
    }
    ids.add(region.id);
    if (
      region.targetIds.length > MAX_REGION_TARGETS ||
      region.sourceRegionIds.length === 0 ||
      region.sourceRegionIds.length > MAX_EVIDENCE_REGIONS_PER_PAGE ||
      region.targetIds.some((id) => !boundedId(id)) ||
      region.sourceRegionIds.some((id) => !boundedId(id)) ||
      new Set(region.targetIds).size !== region.targetIds.length ||
      new Set(region.sourceRegionIds).size !== region.sourceRegionIds.length
    ) {
      throw new Error('The evidence region identity receipts are invalid.');
    }
    validateBounds(region.bounds);
  }
}

function validateBounds(bounds: PdfEvidenceBounds): void {
  const values = [bounds.left, bounds.top, bounds.right, bounds.bottom];
  if (
    values.some(
      (value) =>
        !Number.isInteger(value) ||
        value < 0 ||
        value > PDF_EVIDENCE_COORDINATE_BASIS,
    ) ||
    bounds.right <= bounds.left ||
    bounds.bottom <= bounds.top
  ) {
    throw new Error('The evidence region bounds are invalid.');
  }
}

function positionEvidenceRegions(
  page: PdfEvidencePage,
  runtime: EvidenceRuntime,
  _revision: number,
) {
  const documentId = runtime.documentManager.getActiveDocumentId();
  const documentPage =
    runtime.documentManager.getActiveDocument()?.pages[page.pageNumber - 1];
  if (!documentId || !documentPage) return [];
  const scroll = runtime.scroll.forDocument(documentId);
  const viewport = runtime.viewport.forDocument(documentId);
  const metrics = viewport.getMetrics();
  const viewportGap = runtime.viewport.getViewportGap();
  const zoom =
    runtime.zoom?.forDocument(documentId).getState().currentZoomLevel ?? 1;
  const contentWidth = scroll.getLayout().totalContentSize.width;
  const horizontalCenterOffset = pdfEvidenceHorizontalCenterOffset(
    metrics.clientWidth,
    viewportGap,
    contentWidth,
    zoom,
  );
  const width =
    documentPage.rotation % 2
      ? documentPage.size.height
      : documentPage.size.width;
  const height =
    documentPage.rotation % 2
      ? documentPage.size.width
      : documentPage.size.height;
  return page.regions.flatMap((region) => {
    const rect = pdfEvidenceBoundsToPageRect(region.bounds, width, height);
    const positioned = scroll.getRectPositionForPage(
      page.pageNumber - 1,
      rect,
      undefined,
      0,
    );
    if (!positioned) return [];
    return [
      {
        region,
        style: {
          height: positioned.size.height,
          left:
            positioned.origin.x +
            viewportGap -
            metrics.scrollLeft +
            horizontalCenterOffset +
            metrics.clientLeft,
          top:
            positioned.origin.y +
            viewportGap -
            metrics.scrollTop +
            metrics.clientTop,
          width: positioned.size.width,
        },
      },
    ];
  });
}

function scrollEvidenceRegionIntoView(
  runtime: EvidenceRuntime,
  page: PdfEvidencePage,
  region: PdfEvidenceRegion,
): void {
  const documentId = runtime.documentManager.getActiveDocumentId();
  const documentPage =
    runtime.documentManager.getActiveDocument()?.pages[page.pageNumber - 1];
  if (!documentId || !documentPage) return;
  const width =
    documentPage.rotation % 2
      ? documentPage.size.height
      : documentPage.size.width;
  const height =
    documentPage.rotation % 2
      ? documentPage.size.width
      : documentPage.size.height;
  const rect = pdfEvidenceBoundsToPageRect(region.bounds, width, height);
  const rotatedCenter = {
    x: rect.origin.x + rect.size.width / 2,
    y: rect.origin.y + rect.size.height / 2,
  };
  const pageCoordinates = pdfEvidencePointToUnrotatedPage(
    rotatedCenter,
    documentPage.size.width,
    documentPage.size.height,
    documentPage.rotation,
  );
  runtime.scroll.forDocument(documentId).scrollToPage({
    alignX: 50,
    alignY: 32,
    behavior: 'auto',
    pageCoordinates,
    pageNumber: page.pageNumber,
  });
}

function scrollToEvidencePage(
  runtime: EvidenceRuntime,
  pageNumber: number,
): void {
  const documentId = runtime.documentManager.getActiveDocumentId();
  if (!documentId || !Number.isInteger(pageNumber) || pageNumber < 1) return;
  const scroll = runtime.scroll.forDocument(documentId);
  if (pageNumber > scroll.getTotalPages()) return;
  scroll.scrollToPage({
    behavior: 'auto',
    pageNumber,
  });
}

export function pdfEvidenceBoundsToPageRect(
  bounds: PdfEvidenceBounds,
  width: number,
  height: number,
) {
  return {
    origin: {
      x: (bounds.left / PDF_EVIDENCE_COORDINATE_BASIS) * width,
      y: (bounds.top / PDF_EVIDENCE_COORDINATE_BASIS) * height,
    },
    size: {
      height:
        ((bounds.bottom - bounds.top) / PDF_EVIDENCE_COORDINATE_BASIS) * height,
      width:
        ((bounds.right - bounds.left) / PDF_EVIDENCE_COORDINATE_BASIS) * width,
    },
  };
}

export function pdfEvidencePointToUnrotatedPage(
  point: { x: number; y: number },
  pageWidth: number,
  pageHeight: number,
  rotation: 0 | 1 | 2 | 3,
): { x: number; y: number } {
  switch (rotation) {
    case 1:
      return { x: point.y, y: pageHeight - point.x };
    case 2:
      return { x: pageWidth - point.x, y: pageHeight - point.y };
    case 3:
      return { x: pageWidth - point.y, y: point.x };
    default:
      return point;
  }
}

export function pdfEvidenceHorizontalCenterOffset(
  viewportClientWidth: number,
  viewportGap: number,
  contentWidth: number,
  zoom: number,
): number {
  const availableWidth = viewportClientWidth - viewportGap * 2;
  const renderedContentWidth = contentWidth * zoom;
  return Math.max(0, (availableWidth - renderedContentWidth) / 2);
}

function boundedId(value: string): boolean {
  return value.length > 0 && value.length <= MAX_ID_LENGTH;
}

function positiveFinite(value: number): boolean {
  return Number.isFinite(value) && value > 0;
}

function isSha256(value: string): boolean {
  return /^[a-f0-9]{64}$/.test(value);
}

function evidenceErrorMessage(error: unknown): string {
  return error instanceof Error && error.message
    ? error.message
    : 'Unable to load PDF evidence.';
}
