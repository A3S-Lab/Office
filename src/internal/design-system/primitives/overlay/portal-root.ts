export function officeOverlayPortalRoot(
  ownerDocument: Document,
  ...anchors: Array<Element | null | undefined>
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
