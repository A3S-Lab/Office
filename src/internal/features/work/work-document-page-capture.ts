export interface WorkLiveDocumentCaptureSurface {
  element: HTMLElement;
  pageCount: number;
  pageGap: number;
  pageHeight: number;
  pageWidth: number;
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
  captureHeight = surface.pageHeight,
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
    height: `${captureHeight}px`,
    overflow: 'hidden',
    position: 'relative',
    width: `${surface.pageWidth}px`,
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
  const surfaceHeight =
    surface.pageCount * surface.pageHeight +
    Math.max(0, surface.pageCount - 1) * surface.pageGap;
  Object.assign(snapshot.style, {
    borderColor: 'transparent',
    boxShadow: 'none',
    height: `${surfaceHeight}px`,
    left: '0',
    margin: '0',
    minHeight: `${surfaceHeight}px`,
    position: 'absolute',
    top: `-${firstPageIndex * (surface.pageHeight + surface.pageGap)}px`,
    width: `${surface.pageWidth}px`,
  });
  normalizeWorkLiveDocumentSnapshot(snapshot);

  viewport.append(snapshot);
  host.append(viewport);
  document.body.append(host);
  return { backgroundColor, host, snapshot, viewport };
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
