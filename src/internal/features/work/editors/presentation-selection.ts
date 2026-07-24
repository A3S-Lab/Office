import type { OfficeKernelPresentationAlignment } from '../../../kernel/office-kernel-protocol';
import {
  expandPresentationGroupSelection,
  presentationSelectionUnits,
} from '../work-presentation-groups';
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

export function selectPresentationElementUnit(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
  elementId: string,
  additive: boolean,
): string[] {
  const unit = expandPresentationGroupSelection(elements, [elementId]);
  if (!unit.length) return [...new Set(selectedIds)];
  if (!additive) return unit;
  const normalized = expandPresentationGroupSelection(elements, selectedIds);
  const selected = new Set(normalized);
  const unitSelected = unit.every((id) => selected.has(id));
  if (unitSelected) {
    const unitIds = new Set(unit);
    return normalized.filter((id) => !unitIds.has(id));
  }
  const result = normalized.filter((id) => !unit.includes(id));
  return [...result, ...unit];
}

export function selectedPresentationElements(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
): WorkSlideElement[] {
  const selected = new Set(
    expandPresentationGroupSelection(elements, selectedIds),
  );
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
  const selectedSet = new Set(
    expandPresentationGroupSelection(elements, selectedIds),
  );
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
  const units = presentationSelectionUnits(elements, selectedIds);
  const bounds = presentationSelectionBounds(
    units.flatMap((unit) => unit.elements),
  );
  if (!bounds || units.length < 2) return [...elements];
  const deltas = new Map<string, { x: number; y: number }>();
  for (const unit of units) {
    const unitBounds = presentationSelectionBounds(unit.elements);
    if (!unitBounds) continue;
    let deltaX = 0;
    let deltaY = 0;
    if (alignment === 'left') deltaX = bounds.left - unitBounds.left;
    else if (alignment === 'center') {
      deltaX =
        bounds.left +
        bounds.width / 2 -
        (unitBounds.left + unitBounds.width / 2);
    } else if (alignment === 'right') {
      deltaX = bounds.right - unitBounds.right;
    } else if (alignment === 'top') {
      deltaY = bounds.top - unitBounds.top;
    } else if (alignment === 'middle') {
      deltaY =
        bounds.top +
        bounds.height / 2 -
        (unitBounds.top + unitBounds.height / 2);
    } else {
      deltaY = bounds.bottom - unitBounds.bottom;
    }
    for (const elementId of unit.elementIds) {
      deltas.set(elementId, { x: deltaX, y: deltaY });
    }
  }
  return elements.map((element) => {
    const delta = deltas.get(element.id);
    return delta
      ? { ...element, x: element.x + delta.x, y: element.y + delta.y }
      : element;
  });
}

export function distributePresentationSelection(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
  direction: PresentationDistribution,
): WorkSlideElement[] {
  const units = presentationSelectionUnits(elements, selectedIds)
    .map((unit) => ({
      ...unit,
      bounds: presentationSelectionBounds(unit.elements),
    }))
    .filter(
      (
        unit,
      ): unit is ReturnType<typeof presentationSelectionUnits>[number] & {
        bounds: PresentationSelectionBounds;
      } => Boolean(unit.bounds),
    );
  if (units.length < 3) return [...elements];
  const sorted = [...units].sort((left, right) =>
    direction === 'horizontal'
      ? left.bounds.left - right.bounds.left ||
        left.key.localeCompare(right.key)
      : left.bounds.top - right.bounds.top || left.key.localeCompare(right.key),
  );
  const start =
    direction === 'horizontal' ? sorted[0].bounds.left : sorted[0].bounds.top;
  const last = sorted.at(-1);
  if (!last) return [...elements];
  const end =
    direction === 'horizontal' ? last.bounds.right : last.bounds.bottom;
  const occupied = sorted.reduce(
    (total, unit) =>
      total +
      (direction === 'horizontal' ? unit.bounds.width : unit.bounds.height),
    0,
  );
  const gap = (end - start - occupied) / (sorted.length - 1);
  const deltas = new Map<string, number>();
  let cursor = start;
  for (const unit of sorted) {
    const unitStart =
      direction === 'horizontal' ? unit.bounds.left : unit.bounds.top;
    for (const elementId of unit.elementIds) {
      deltas.set(elementId, cursor - unitStart);
    }
    cursor +=
      (direction === 'horizontal' ? unit.bounds.width : unit.bounds.height) +
      gap;
  }
  return elements.map((element) => {
    const delta = deltas.get(element.id);
    if (delta === undefined) return element;
    return direction === 'horizontal'
      ? { ...element, x: element.x + delta }
      : { ...element, y: element.y + delta };
  });
}

export function reorderPresentationSelection(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
  direction: -1 | 1,
): WorkSlideElement[] {
  const reordered = [...elements];
  const selected = new Set(
    expandPresentationGroupSelection(elements, selectedIds),
  );
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
