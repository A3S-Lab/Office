export function officeOverlayPortalRoot(
  ownerDocument: Document,
  ...anchors: Array<Element | null | undefined>
): HTMLElement {
  for (const anchor of anchors) {
    const officeRoot = anchor?.closest<HTMLElement>('[data-a3s-office]');
    if (officeRoot) return officeRoot;
  }
  const activeElement = ownerDocument.activeElement;
  if (activeElement instanceof Element) {
    const officeRoot = activeElement.closest<HTMLElement>('[data-a3s-office]');
    if (officeRoot) return officeRoot;
  }
  return ownerDocument.body;
}
