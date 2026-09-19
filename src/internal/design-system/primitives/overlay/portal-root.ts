export function officeOverlayPortalRoot(
  ownerDocument: Document,
  ...anchors: Array<Element | null | undefined>
): HTMLElement {
  return resolveOfficeOverlayPortalRoot(ownerDocument, anchors);
}

export function officeFloatingPortalRoot(
  ownerDocument: Document,
  ...anchors: Array<Element | null | undefined>
): HTMLElement {
  const root = resolveOfficeOverlayPortalRoot(ownerDocument, anchors);
  if (!shouldEscapeClippedPortal(root)) return root;
  return floatingThemeRoot(ownerDocument, root);
}

function resolveOfficeOverlayPortalRoot(
  ownerDocument: Document,
  anchors: Array<Element | null | undefined>,
): HTMLElement {
  for (const anchor of anchors) {
    const portalRoot = closestOfficeOverlayRoot(anchor);
    if (portalRoot) return portalRoot;
  }
  const activeElement = ownerDocument.activeElement;
  if (activeElement instanceof Element) {
    const portalRoot = closestOfficeOverlayRoot(activeElement);
    if (portalRoot) return portalRoot;
  }
  return ownerDocument.body;
}

function closestOfficeOverlayRoot(anchor: Element | null | undefined) {
  return (
    anchor?.closest<HTMLElement>('[role="dialog"][aria-modal="true"]') ??
    anchor?.closest<HTMLElement>('[data-a3s-office]') ??
    null
  );
}

function shouldEscapeClippedPortal(root: HTMLElement): boolean {
  if (root.getAttribute('role') === 'dialog') return false;
  const ownerDocument = root.ownerDocument;
  // `.a3s-office` is overflow:hidden by design. Only a host outside that root
  // moves floating overlays onto a theme-preserving root. Dialogs stay put.
  let node = root.parentElement;
  while (
    node &&
    node !== ownerDocument.body &&
    node !== ownerDocument.documentElement
  ) {
    if (clipsFixedOverlay(node)) return true;
    node = node.parentElement;
  }
  return false;
}

function floatingThemeRoot(
  ownerDocument: Document,
  source: HTMLElement,
): HTMLElement {
  const existing = ownerDocument.body.querySelector<HTMLElement>(
    ':scope > .a3s-office-floating-root',
  );
  const root = existing ?? ownerDocument.createElement('div');
  root.className = 'a3s-office a3s-office-floating-root';
  root.setAttribute('data-a3s-office', '');
  root.setAttribute(
    'data-theme',
    source.getAttribute('data-theme') ??
      source.closest('[data-theme]')?.getAttribute('data-theme') ??
      'system',
  );
  if (!existing) ownerDocument.body.append(root);
  return root;
}

function clipsFixedOverlay(element: HTMLElement): boolean {
  const view = element.ownerDocument.defaultView;
  if (!view) return false;
  const style = view.getComputedStyle(element);
  const clipsAxis = (value: string) => value !== '' && value !== 'visible';
  const identity = (value: string) => value === '' || value === 'none';
  if (
    clipsAxis(style.overflow) ||
    clipsAxis(style.overflowX) ||
    clipsAxis(style.overflowY)
  ) {
    return true;
  }
  if (
    !identity(style.transform) ||
    !identity(style.filter) ||
    !identity(style.perspective) ||
    !identity(style.backdropFilter)
  ) {
    return true;
  }
  if (
    style.contain.includes('paint') ||
    style.contain === 'strict' ||
    style.contain === 'content'
  ) {
    return true;
  }
  return style.willChange
    .split(',')
    .some((value) =>
      ['transform', 'filter', 'perspective', 'contain'].includes(value.trim()),
    );
}
