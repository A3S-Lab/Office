export const WORK_OFFICE_RIBBON_SCROLL_INSET = 34;

export interface WorkOfficeRibbonItemGeometry {
  left: number;
  right: number;
}

interface WorkOfficeRibbonViewport {
  clientWidth: number;
  items: readonly WorkOfficeRibbonItemGeometry[];
  scrollLeft: number;
}

export function calculateRibbonOverflow({
  clientWidth,
  items,
  scrollLeft,
}: WorkOfficeRibbonViewport): { backward: boolean; forward: boolean } {
  const lastItem = items.at(-1);
  if (!lastItem || clientWidth <= 0) {
    return { backward: false, forward: false };
  }
  if (lastItem.right <= clientWidth - 2) {
    return { backward: false, forward: false };
  }
  return {
    backward: scrollLeft > 2,
    forward: lastItem.right > scrollLeft + clientWidth - 2,
  };
}

export function calculateRibbonScrollTarget({
  clientWidth,
  direction,
  items,
  scrollLeft,
}: WorkOfficeRibbonViewport & { direction: -1 | 1 }): number {
  if (!items.length || clientWidth <= 0) return scrollLeft;
  if ((items.at(-1)?.right ?? 0) <= clientWidth - 2) return 0;
  const visibleLeft = scrollLeft + WORK_OFFICE_RIBBON_SCROLL_INSET;
  const visibleRight =
    scrollLeft + clientWidth - WORK_OFFICE_RIBBON_SCROLL_INSET;
  const target =
    direction === 1
      ? items.find(
          (item) =>
            item.left > visibleLeft + 2 && item.right > visibleRight + 2,
        )
      : [...items].reverse().find((item) => item.left < visibleLeft - 2);
  if (!target) return scrollLeft;
  return Math.max(0, target.left - WORK_OFFICE_RIBBON_SCROLL_INSET);
}
