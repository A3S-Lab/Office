import type { OfficeKernelPresentationAlignment } from '../../../kernel/office-kernel-protocol';
import type { WorkSlideElement } from '../work-types';

export type PresentationDistribution = 'horizontal' | 'vertical';

export interface PresentationSelectionBounds {
  bottom: number;
  height: number;
  left: number;
  right: number;
  top: number;
  width: number;
}

export function selectPresentationElement(
  selectedIds: readonly string[],
  elementId: string,
  additive: boolean,
): string[] {
  if (!additive) return [elementId];
  const normalized = [...new Set(selectedIds)];
  return normalized.includes(elementId)
    ? normalized.filter((id) => id !== elementId)
    : [...normalized, elementId];
}

export function selectedPresentationElements(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
): WorkSlideElement[] {
  const selected = new Set(selectedIds);
  return elements.filter((element) => selected.has(element.id));
}

export function presentationSelectionBounds(
  elements: readonly WorkSlideElement[],
): PresentationSelectionBounds | null {
  if (!elements.length) return null;
  const left = Math.min(...elements.map((element) => element.x));
  const top = Math.min(...elements.map((element) => element.y));
  const right = Math.max(
    ...elements.map((element) => element.x + element.width),
  );
  const bottom = Math.max(
    ...elements.map((element) => element.y + element.height),
  );
  return {
    bottom,
    height: bottom - top,
    left,
    right,
    top,
    width: right - left,
  };
}

export function translatePresentationSelection(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
  deltaX: number,
  deltaY: number,
): WorkSlideElement[] {
  const selected = selectedPresentationElements(elements, selectedIds);
  const bounds = presentationSelectionBounds(selected);
  if (!bounds) return [...elements];
  const boundedX = clamp(deltaX, -bounds.left, 100 - bounds.right);
  const boundedY = clamp(deltaY, -bounds.top, 100 - bounds.bottom);
  const selectedSet = new Set(selectedIds);
  return elements.map((element) =>
    selectedSet.has(element.id)
      ? { ...element, x: element.x + boundedX, y: element.y + boundedY }
      : element,
  );
}

export function alignPresentationSelection(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
  alignment: OfficeKernelPresentationAlignment,
): WorkSlideElement[] {
  const selected = selectedPresentationElements(elements, selectedIds);
  const bounds = presentationSelectionBounds(selected);
  if (!bounds || selected.length < 2) return [...elements];
  const selectedSet = new Set(selectedIds);
  return elements.map((element) => {
    if (!selectedSet.has(element.id)) return element;
    if (alignment === 'left') return { ...element, x: bounds.left };
    if (alignment === 'center') {
      return {
        ...element,
        x: bounds.left + (bounds.width - element.width) / 2,
      };
    }
    if (alignment === 'right') {
      return { ...element, x: bounds.right - element.width };
    }
    if (alignment === 'top') return { ...element, y: bounds.top };
    if (alignment === 'middle') {
      return {
        ...element,
        y: bounds.top + (bounds.height - element.height) / 2,
      };
    }
    return { ...element, y: bounds.bottom - element.height };
  });
}

export function distributePresentationSelection(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
  direction: PresentationDistribution,
): WorkSlideElement[] {
  const selected = selectedPresentationElements(elements, selectedIds);
  if (selected.length < 3) return [...elements];
  const sorted = [...selected].sort((left, right) =>
    direction === 'horizontal'
      ? left.x - right.x || left.id.localeCompare(right.id)
      : left.y - right.y || left.id.localeCompare(right.id),
  );
  const start = direction === 'horizontal' ? sorted[0].x : sorted[0].y;
  const last = sorted.at(-1);
  if (!last) return [...elements];
  const end =
    direction === 'horizontal' ? last.x + last.width : last.y + last.height;
  const occupied = sorted.reduce(
    (total, element) =>
      total + (direction === 'horizontal' ? element.width : element.height),
    0,
  );
  const gap = (end - start - occupied) / (sorted.length - 1);
  const positions = new Map<string, number>();
  let cursor = start;
  for (const element of sorted) {
    positions.set(element.id, cursor);
    cursor +=
      (direction === 'horizontal' ? element.width : element.height) + gap;
  }
  return elements.map((element) => {
    const position = positions.get(element.id);
    if (position === undefined) return element;
    return direction === 'horizontal'
      ? { ...element, x: position }
      : { ...element, y: position };
  });
}

export function reorderPresentationSelection(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
  direction: -1 | 1,
): WorkSlideElement[] {
  const reordered = [...elements];
  const selected = new Set(selectedIds);
  if (direction > 0) {
    for (let index = reordered.length - 2; index >= 0; index -= 1) {
      if (
        selected.has(reordered[index].id) &&
        !selected.has(reordered[index + 1].id)
      ) {
        [reordered[index], reordered[index + 1]] = [
          reordered[index + 1],
          reordered[index],
        ];
      }
    }
    return reordered;
  }
  for (let index = 1; index < reordered.length; index += 1) {
    if (
      selected.has(reordered[index].id) &&
      !selected.has(reordered[index - 1].id)
    ) {
      [reordered[index], reordered[index - 1]] = [
        reordered[index - 1],
        reordered[index],
      ];
    }
  }
  return reordered;
}

export function presentationElementCanEditContent(
  element: WorkSlideElement,
): boolean {
  return (
    presentationElementSupportsTextFormatting(element) ||
    element.type === 'table'
  );
}

export function presentationElementSupportsTextFormatting(
  element: WorkSlideElement,
): boolean {
  return (
    element.type === 'text' ||
    element.type === 'shape' ||
    Boolean(element.text || element.textRuns?.length)
  );
}

function clamp(value: number, minimum: number, maximum: number): number {
  return Math.min(Math.max(value, minimum), maximum);
}
