export interface WorkLiveDocumentCapturePage {
  height: number;
  left: number;
  top: number;
  width: number;
}

export interface WorkLiveDocumentCaptureSurface {
  element: HTMLElement;
  pageCount: number;
  pageGap: number;
  pageHeight: number;
  pageWidth: number;
  pages?: readonly WorkLiveDocumentCapturePage[];
}

export interface WorkLiveDocumentCapture {
  backgroundColor: string;
  host: HTMLDivElement;
  snapshot: HTMLElement;
  viewport: HTMLDivElement;
}

export function mountWorkLiveDocumentCapture(
  surface: WorkLiveDocumentCaptureSurface,
  firstPageIndex = 0,
  captureHeight?: number,
): WorkLiveDocumentCapture {
  const host = document.createElement('div');
  host.className = 'work-pdf-export-surface';
  host.setAttribute('aria-hidden', 'true');

  const viewport = document.createElement('div');
  viewport.className = 'work-document-live-pdf-viewport';
  const backgroundColor =
    getComputedStyle(surface.element).backgroundColor || '#ffffff';
  Object.assign(viewport.style, {
    backgroundColor,
    overflow: 'hidden',
    position: 'relative',
  });

  const snapshot = surface.element.cloneNode(true) as HTMLElement;
  snapshot.removeAttribute('data-work-pdf-live-document');
  snapshot.removeAttribute('aria-label');
  snapshot.classList.remove(
    'page-chrome-editing',
    'work-document-preview-page',
  );
  snapshot.classList.add(
    'document',
    'work-document-live-pdf-snapshot',
    'work-pdf-export-page',
  );
  snapshot.dataset.documentCommentAppearance = 'plain';
  snapshot.setAttribute('aria-hidden', 'true');
  const pages = capturePages(surface);
  const surfaceHeight = Math.max(
    1,
    ...pages.map((page) => page.top + page.height),
  );
  const surfaceWidth = Math.max(
    1,
    ...pages.map((page) => page.left + page.width),
  );
  Object.assign(snapshot.style, {
    borderColor: 'transparent',
    boxShadow: 'none',
    height: `${surfaceHeight}px`,
    margin: '0',
    minHeight: `${surfaceHeight}px`,
    position: 'absolute',
    width: `${surfaceWidth}px`,
  });
  normalizeWorkLiveDocumentSnapshot(snapshot);

  viewport.append(snapshot);
  host.append(viewport);
  document.body.append(host);
  const capture = { backgroundColor, host, snapshot, viewport };
  positionWorkLiveDocumentCapture(
    capture,
    surface,
    firstPageIndex,
    captureHeight,
  );
  return capture;
}

export function positionWorkLiveDocumentCapture(
  capture: WorkLiveDocumentCapture,
  surface: WorkLiveDocumentCaptureSurface,
  pageIndex: number,
  captureHeight?: number,
): WorkLiveDocumentCapturePage {
  const pages = capturePages(surface);
  const page = pages[pageIndex];
  if (!page)
    throw new Error('The requested live document page is unavailable.');
  Object.assign(capture.viewport.style, {
    height: `${captureHeight ?? page.height}px`,
    width: `${page.width}px`,
  });
  Object.assign(capture.snapshot.style, {
    left: `-${page.left}px`,
    top: `-${page.top}px`,
  });
  return page;
}

export function normalizeWorkLiveDocumentSnapshot(snapshot: HTMLElement): void {
  snapshot
    .querySelectorAll('.work-document-selection-toolbar')
    .forEach((element) => {
      element.remove();
    });
  for (const editor of snapshot.querySelectorAll<HTMLElement>(
    '.work-document-page-chrome-inline-editor',
  )) {
    const content = editor.querySelector<HTMLElement>(
      '[data-document-page-chrome-engine="tiptap"]',
    );
    const replacement = document.createElement('div');
    replacement.className = 'work-document-page-chrome-html';
    replacement.innerHTML = content?.innerHTML ?? '';
    editor.replaceWith(replacement);
  }
  for (const element of snapshot.querySelectorAll<HTMLElement>('*')) {
    element.removeAttribute('contenteditable');
    element.removeAttribute('id');
    element.removeAttribute('tabindex');
    element.classList.remove(
      'ProseMirror-focused',
      'ProseMirror-selectednode',
      'is-editor-empty',
      'selectedCell',
    );
  }
}

function capturePages(
  surface: WorkLiveDocumentCaptureSurface,
): WorkLiveDocumentCapturePage[] {
  if (surface.pages?.length === surface.pageCount) {
    return surface.pages.map((page) => ({ ...page }));
  }
  return Array.from({ length: surface.pageCount }, (_, pageIndex) => ({
    height: surface.pageHeight,
    left: 0,
    top: pageIndex * (surface.pageHeight + surface.pageGap),
    width: surface.pageWidth,
  }));
}
