import { afterEach, expect, test } from '@rstest/core';
import { act, render, screen, waitFor } from '@testing-library/react';
import {
  DocumentPageThumbnail,
  renderDocumentPageThumbnail,
  type WorkDocumentPageThumbnailSource,
} from '../src/internal/features/work/editors/document-page-thumbnail';

const originalIntersectionObserver = globalThis.IntersectionObserver;

afterEach(() => {
  document.body.replaceChildren();
  Object.defineProperty(globalThis, 'IntersectionObserver', {
    configurable: true,
    value: originalIntersectionObserver,
    writable: true,
  });
});

test('rasterizes one exact physical page from the live paginated surface', async () => {
  const source = createThumbnailSource();
  let capturedFontReadyTimeout: number | undefined;
  let capturedViewport: HTMLElement | null = null;
  let capturedScale = 0;

  const image = await renderDocumentPageThumbnail(
    source,
    1,
    276,
    async (viewport, options) => {
      capturedFontReadyTimeout = options.fontReadyTimeout;
      capturedViewport = viewport;
      capturedScale = options.scale;
      const canvas = document.createElement('canvas');
      Object.defineProperty(canvas, 'toDataURL', {
        value: () => 'data:image/png;base64,page-two',
      });
      return canvas;
    },
  );

  expect(image).toBe('data:image/png;base64,page-two');
  expect(capturedFontReadyTimeout).toBe(0);
  expect(capturedScale).toBeCloseTo(276 / source.pageWidth);
  expect(capturedViewport).not.toBeNull();
  expect(capturedViewport?.style.height).toBe(`${source.pageHeight}px`);
  expect(capturedViewport?.style.width).toBe(`${source.pageWidth}px`);
  const snapshot = capturedViewport?.querySelector<HTMLElement>(
    '.work-document-live-pdf-snapshot',
  );
  expect(snapshot?.style.top).toBe(`-${source.pageHeight + source.pageGap}px`);
  expect(snapshot?.querySelector('h1')).toHaveTextContent('项目方案');
  expect(
    snapshot?.querySelector('.ProseMirror')?.hasAttribute('contenteditable'),
  ).toBe(false);
  expect(document.querySelector('.work-pdf-export-surface')).toBeNull();
});

test('captures a mixed-size page at its own viewport and surface offset', async () => {
  const source = createThumbnailSource();
  source.pages = [
    { height: 400, left: 100, top: 0, width: 300 },
    { height: 200, left: 0, top: 420, width: 500 },
    { height: 600, left: 50, top: 650, width: 400 },
  ];
  let capturedViewport: HTMLElement | null = null;
  let capturedScale = 0;

  await renderDocumentPageThumbnail(
    source,
    1,
    276,
    async (viewport, options) => {
      capturedViewport = viewport;
      capturedScale = options.scale;
      const canvas = document.createElement('canvas');
      Object.defineProperty(canvas, 'toDataURL', {
        value: () => 'data:image/png;base64,mixed-page-two',
      });
      return canvas;
    },
  );

  expect(capturedScale).toBeCloseTo(276 / 500);
  expect(capturedViewport?.style.height).toBe('200px');
  expect(capturedViewport?.style.width).toBe('500px');
  const snapshot = capturedViewport?.querySelector<HTMLElement>(
    '.work-document-live-pdf-snapshot',
  );
  expect(snapshot?.style.left).toBe('0px');
  expect(snapshot?.style.top).toBe('-420px');
  expect(snapshot?.style.width).toBe('500px');
});

test('does not wait for unrelated page-wide fonts after pagination is ready', async () => {
  const fontsDescriptor = Object.getOwnPropertyDescriptor(document, 'fonts');
  const originalRequestAnimationFrame = globalThis.requestAnimationFrame;
  Object.defineProperty(document, 'fonts', {
    configurable: true,
    value: { ready: new Promise(() => undefined) },
  });
  globalThis.requestAnimationFrame = ((callback: FrameRequestCallback) => {
    callback(0);
    return 1;
  }) as typeof requestAnimationFrame;
  let rendererCalled = false;

  try {
    const rendering = renderDocumentPageThumbnail(
      createThumbnailSource(),
      0,
      276,
      async () => {
        rendererCalled = true;
        const canvas = document.createElement('canvas');
        Object.defineProperty(canvas, 'toDataURL', {
          value: () => 'data:image/png;base64,ready-layout',
        });
        return canvas;
      },
    );
    for (let turn = 0; turn < 4; turn += 1) await Promise.resolve();
    expect(rendererCalled).toBe(true);
    await expect(rendering).resolves.toBe('data:image/png;base64,ready-layout');
  } finally {
    globalThis.requestAnimationFrame = originalRequestAnimationFrame;
    if (fontsDescriptor) {
      Object.defineProperty(document, 'fonts', fontsDescriptor);
    } else {
      Reflect.deleteProperty(document, 'fonts');
    }
  }
});

test('does not depend on animation frames that background agent tabs can suspend', async () => {
  const originalRequestAnimationFrame = globalThis.requestAnimationFrame;
  let pendingFrame: FrameRequestCallback | undefined;
  globalThis.requestAnimationFrame = ((callback: FrameRequestCallback) => {
    pendingFrame = callback;
    return 1;
  }) as typeof requestAnimationFrame;
  let rendererCalled = false;
  let rendering: Promise<string> | undefined;

  try {
    rendering = renderDocumentPageThumbnail(
      createThumbnailSource(),
      0,
      276,
      async () => {
        rendererCalled = true;
        const canvas = document.createElement('canvas');
        Object.defineProperty(canvas, 'toDataURL', {
          value: () => 'data:image/png;base64,background-tab',
        });
        return canvas;
      },
    );
    for (let turn = 0; turn < 4; turn += 1) await Promise.resolve();
    expect(rendererCalled).toBe(true);
    await expect(rendering).resolves.toBe(
      'data:image/png;base64,background-tab',
    );
  } finally {
    pendingFrame?.(0);
    await rendering?.catch(() => undefined);
    globalThis.requestAnimationFrame = originalRequestAnimationFrame;
  }
});

test('replaces the text fallback with a non-interactive page raster', async () => {
  const source = createThumbnailSource();
  render(
    <DocumentPageThumbnail
      backgroundColor="#ffffff"
      fallbackText="项目方案 背景与目标"
      orientation="portrait"
      pageIndex={0}
      priority
      source={source}
      renderThumbnail={async () => 'data:image/png;base64,page-one'}
    />,
  );

  const thumbnail = screen.getByTestId('document-page-thumbnail');
  expect(thumbnail).toHaveAttribute('data-thumbnail-state', 'loading');
  expect(thumbnail.style.aspectRatio).toBe(
    `${source.pageWidth} / ${source.pageHeight}`,
  );
  expect(thumbnail).toHaveTextContent('项目方案 背景与目标');

  await waitFor(() =>
    expect(thumbnail).toHaveAttribute('data-thumbnail-state', 'ready'),
  );
  const image = thumbnail.querySelector('img');
  expect(image).toHaveAttribute('src', 'data:image/png;base64,page-one');
  expect(image).toHaveAttribute('alt', '');
  expect(thumbnail).not.toHaveTextContent('项目方案 背景与目标');
});

test('captures only thumbnails admitted by the bounded viewport observer', async () => {
  const observers: TestIntersectionObserver[] = [];
  class TestIntersectionObserver {
    readonly callback: IntersectionObserverCallback;
    target: Element | null = null;

    constructor(callback: IntersectionObserverCallback) {
      this.callback = callback;
      observers.push(this);
    }

    disconnect() {}
    observe(target: Element) {
      this.target = target;
    }
    unobserve() {}
    takeRecords(): IntersectionObserverEntry[] {
      return [];
    }
    readonly root = null;
    readonly rootMargin = '0px';
    readonly thresholds = [0];

    trigger(isIntersecting: boolean) {
      if (!this.target) return;
      this.callback(
        [
          {
            isIntersecting,
            target: this.target,
          } as IntersectionObserverEntry,
        ],
        this as unknown as IntersectionObserver,
      );
    }
  }
  Object.defineProperty(globalThis, 'IntersectionObserver', {
    configurable: true,
    value: TestIntersectionObserver,
    writable: true,
  });

  let captures = 0;
  const source = { ...createThumbnailSource(), pageCount: 20 };
  const pages = Array.from({ length: 20 }, (_, pageIndex) => ({
    id: `page-${pageIndex + 1}`,
    pageIndex,
  }));
  render(
    pages.map(({ id, pageIndex }) => (
      <DocumentPageThumbnail
        key={id}
        fallbackText={`第 ${pageIndex + 1} 页`}
        orientation="portrait"
        pageIndex={pageIndex}
        source={source}
        renderThumbnail={async () => {
          captures += 1;
          return `data:image/png;base64,page-${pageIndex + 1}`;
        }}
      />
    )),
  );

  expect(observers).toHaveLength(20);
  expect(captures).toBe(0);

  act(() => {
    observers[7]?.trigger(true);
    observers[8]?.trigger(true);
    observers[9]?.trigger(true);
  });
  await waitFor(() => expect(captures).toBe(3));

  act(() => observers[8]?.trigger(false));
  await waitFor(() =>
    expect(screen.getAllByTestId('document-page-thumbnail')[8]).toHaveAttribute(
      'data-thumbnail-state',
      'idle',
    ),
  );
  expect(captures).toBe(3);
});

test('drops queued captures when a thumbnail leaves the mounted window', async () => {
  let releaseFirstCapture: (() => void) | undefined;
  let secondPageCaptures = 0;
  const firstCapture = new Promise<void>((resolve) => {
    releaseFirstCapture = resolve;
  });
  const source = createThumbnailSource();
  const renderThumbnail = async (
    _source: WorkDocumentPageThumbnailSource,
    pageIndex: number,
  ) => {
    if (pageIndex === 0) await firstCapture;
    if (pageIndex === 1) secondPageCaptures += 1;
    return `data:image/png;base64,page-${pageIndex + 1}`;
  };
  const view = render(
    <>
      <DocumentPageThumbnail
        fallbackText="第 1 页"
        orientation="portrait"
        pageIndex={0}
        priority
        source={source}
        renderThumbnail={renderThumbnail}
      />
      <DocumentPageThumbnail
        fallbackText="第 2 页"
        orientation="portrait"
        pageIndex={1}
        priority
        source={source}
        renderThumbnail={renderThumbnail}
      />
    </>,
  );

  view.rerender(
    <DocumentPageThumbnail
      fallbackText="第 1 页"
      orientation="portrait"
      pageIndex={0}
      priority
      source={source}
      renderThumbnail={renderThumbnail}
    />,
  );
  await act(async () => releaseFirstCapture?.());
  await waitFor(() =>
    expect(screen.getByTestId('document-page-thumbnail')).toHaveAttribute(
      'data-thumbnail-state',
      'ready',
    ),
  );
  expect(secondPageCaptures).toBe(0);
});

function createThumbnailSource(): WorkDocumentPageThumbnailSource {
  const element = document.createElement('article');
  element.className = 'work-document-page paginated';
  element.innerHTML = `
    <div class="work-document-page-stack"></div>
    <section class="work-document-editable">
      <div class="ProseMirror ProseMirror-focused" contenteditable="true">
        <h1>项目方案</h1>
        <p>正文内容</p>
      </div>
    </section>
  `;
  document.body.append(element);
  return {
    element,
    pageCount: 3,
    pageGap: 28,
    pageHeight: 1123,
    pageWidth: 794,
    revision: 'revision-1',
  };
}
