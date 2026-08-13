import { useEffect, useRef, useState } from 'react';
import {
  mountWorkLiveDocumentCapture,
  type WorkLiveDocumentCapturePage,
} from '../work-document-page-capture';

export interface WorkDocumentPageThumbnailSource {
  element: HTMLElement;
  pageCount: number;
  pageGap: number;
  pageHeight: number;
  pageWidth: number;
  pages?: readonly WorkLiveDocumentCapturePage[];
  revision: string;
}

export interface DocumentPageCanvasRenderOptions {
  backgroundColor: string;
  fontReadyTimeout: 0;
  height: number;
  logging: false;
  scale: number;
  useCORS: true;
  width: number;
  windowHeight: number;
  windowWidth: number;
}

export type DocumentPageCanvasRenderer = (
  element: HTMLElement,
  options: DocumentPageCanvasRenderOptions,
) => Promise<HTMLCanvasElement>;

export type DocumentPageThumbnailRenderer = (
  source: WorkDocumentPageThumbnailSource,
  pageIndex: number,
  targetPixelWidth: number,
) => Promise<string>;

type ThumbnailState = 'failed' | 'idle' | 'loading' | 'ready' | 'refreshing';

let thumbnailCaptureQueue: Promise<unknown> = Promise.resolve();

export function DocumentPageThumbnail({
  backgroundColor,
  fallbackText,
  orientation,
  pageIndex,
  priority = false,
  source,
  renderThumbnail = renderDocumentPageThumbnail,
}: {
  backgroundColor?: string;
  fallbackText: string;
  orientation: 'portrait' | 'landscape';
  pageIndex: number;
  priority?: boolean;
  source?: WorkDocumentPageThumbnailSource;
  renderThumbnail?: DocumentPageThumbnailRenderer;
}) {
  const rootRef = useRef<HTMLSpanElement>(null);
  const nearViewport = useNearThumbnailViewport(rootRef);
  const captureActive = priority || nearViewport;
  const sourceMutation = useThumbnailSourceMutation(source, captureActive);
  const [image, setImage] = useState<string | null>(null);
  const [state, setState] = useState<ThumbnailState>(
    source ? 'loading' : 'idle',
  );
  const sourcePage = source ? thumbnailSourcePage(source, pageIndex) : null;

  useEffect(() => {
    if (!source || !captureActive) {
      setImage(null);
      setState('idle');
      return;
    }
    if (pageIndex < 0 || pageIndex >= source.pageCount) {
      setImage(null);
      setState('failed');
      return;
    }

    let cancelled = false;
    setState((current) => (current === 'ready' ? 'refreshing' : 'loading'));
    const targetPixelWidth = orientation === 'landscape' ? 340 : 276;
    const task = async () => {
      if (cancelled) return;
      const nextImage = await renderThumbnail(
        source,
        pageIndex,
        targetPixelWidth,
      );
      if (cancelled) return;
      setImage(nextImage);
      setState('ready');
    };
    const queued = thumbnailCaptureQueue.then(task, task);
    thumbnailCaptureQueue = queued.catch(() => undefined);
    void queued.catch(() => {
      if (cancelled) return;
      setImage(null);
      setState('failed');
    });
    return () => {
      cancelled = true;
    };
  }, [
    captureActive,
    orientation,
    pageIndex,
    renderThumbnail,
    source?.element,
    source?.pageCount,
    source?.pageGap,
    source?.pageHeight,
    source?.pageWidth,
    sourcePage?.height,
    sourcePage?.left,
    sourcePage?.top,
    sourcePage?.width,
    source?.revision,
    sourceMutation,
  ]);

  return (
    <span
      ref={rootRef}
      className={`work-document-page-thumbnail ${orientation}`}
      data-testid="document-page-thumbnail"
      data-thumbnail-source={source ? 'live-page' : 'text-fallback'}
      data-thumbnail-state={state}
      aria-hidden="true"
      style={{
        aspectRatio: source
          ? `${sourcePage?.width ?? source.pageWidth} / ${
              sourcePage?.height ?? source.pageHeight
            }`
          : undefined,
        backgroundColor,
      }}
    >
      {image ? (
        <img
          alt=""
          data-document-page-raster={pageIndex + 1}
          draggable={false}
          src={image}
        />
      ) : (
        <span className="work-document-page-thumbnail-fallback">
          {fallbackText || '空白页'}
        </span>
      )}
    </span>
  );
}

export async function renderDocumentPageThumbnail(
  source: WorkDocumentPageThumbnailSource,
  pageIndex: number,
  targetPixelWidth: number,
  renderCanvas: DocumentPageCanvasRenderer = renderDocumentPageCanvas,
): Promise<string> {
  if (
    !Number.isSafeInteger(pageIndex) ||
    pageIndex < 0 ||
    pageIndex >= source.pageCount
  ) {
    throw new Error('The requested document thumbnail page is out of range.');
  }
  const capture = mountWorkLiveDocumentCapture(source, pageIndex);
  try {
    // The live source is admitted only after pagination is ready. Waiting for
    // a document-wide FontFaceSet or a requestAnimationFrame can hang on
    // unrelated faces or a background agent tab. The mounted clone has all
    // dimensions set synchronously, and html2canvas performs its own layout
    // read before rendering.
    const page = thumbnailSourcePage(source, pageIndex);
    const scale = Math.max(0.1, Math.min(1, targetPixelWidth / page.width));
    const canvas = await renderCanvas(capture.viewport, {
      backgroundColor: capture.backgroundColor,
      fontReadyTimeout: 0,
      height: page.height,
      logging: false,
      scale,
      useCORS: true,
      width: page.width,
      windowHeight: Math.ceil(page.height),
      windowWidth: Math.ceil(page.width),
    });
    return canvas.toDataURL('image/png');
  } finally {
    capture.host.remove();
  }
}

function thumbnailSourcePage(
  source: WorkDocumentPageThumbnailSource,
  pageIndex: number,
): WorkLiveDocumentCapturePage {
  return (
    source.pages?.[pageIndex] ?? {
      height: source.pageHeight,
      left: 0,
      top: pageIndex * (source.pageHeight + source.pageGap),
      width: source.pageWidth,
    }
  );
}

async function renderDocumentPageCanvas(
  element: HTMLElement,
  options: DocumentPageCanvasRenderOptions,
): Promise<HTMLCanvasElement> {
  const { default: html2canvas } = await import('html2canvas');
  return html2canvas(element, options);
}

function useNearThumbnailViewport(
  rootRef: React.RefObject<HTMLElement | null>,
) {
  const [nearViewport, setNearViewport] = useState(
    () => typeof IntersectionObserver === 'undefined',
  );

  useEffect(() => {
    const element = rootRef.current;
    if (!element) return;
    if (typeof IntersectionObserver === 'undefined') {
      setNearViewport(true);
      return;
    }
    const root = element.closest('.work-document-page-navigation');
    const observer = new IntersectionObserver(
      (entries) => {
        const entry = entries[0];
        if (entry) setNearViewport(entry.isIntersecting);
      },
      {
        root,
        rootMargin: '280px 0px',
      },
    );
    observer.observe(element);
    return () => observer.disconnect();
  }, [rootRef]);

  return nearViewport;
}

function useThumbnailSourceMutation(
  source: WorkDocumentPageThumbnailSource | undefined,
  active: boolean,
): number {
  const [revision, setRevision] = useState(0);

  useEffect(() => {
    if (!source || !active || typeof MutationObserver === 'undefined') return;
    let timer = 0;
    const observer = new MutationObserver(() => {
      window.clearTimeout(timer);
      timer = window.setTimeout(
        () => setRevision((current) => current + 1),
        160,
      );
    });
    observer.observe(source.element, {
      attributes: true,
      characterData: true,
      childList: true,
      subtree: true,
    });
    return () => {
      observer.disconnect();
      window.clearTimeout(timer);
    };
  }, [active, source?.element, source?.revision]);

  return revision;
}
