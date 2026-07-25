import { createWorkId } from './work-templates';
import type { WorkSlideElement } from './work-types';

export interface PresentationSelectionUnit {
  key: string;
  groupId?: string;
  elements: WorkSlideElement[];
  elementIds: string[];
}

export function expandPresentationGroupSelection(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
): string[] {
  const byId = new Map<string, WorkSlideElement>();
  const membersByGroupId = new Map<string, WorkSlideElement[]>();
  for (const element of elements) {
    byId.set(element.id, element);
    const groupId = topPresentationGroupId(element);
    if (!groupId) continue;
    const members = membersByGroupId.get(groupId) ?? [];
    members.push(element);
    membersByGroupId.set(groupId, members);
  }
  const requested = new Set(selectedIds);
  const result: string[] = [];
  const seen = new Set<string>();
  for (const selectedId of selectedIds) {
    const selected = byId.get(selectedId);
    if (!selected) continue;
    const groupId = topPresentationGroupId(selected);
    const members = groupId
      ? (membersByGroupId.get(groupId) ?? [selected])
      : [selected];
    if (!members.every((member) => requested.has(member.id))) {
      for (const member of members) {
        if (member.id === selectedId || seen.has(member.id)) continue;
        seen.add(member.id);
        result.push(member.id);
      }
    }
    if (!seen.has(selectedId)) {
      seen.add(selectedId);
      result.push(selectedId);
    }
  }
  return result;
}

export function presentationSelectionUnits(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
): PresentationSelectionUnit[] {
  const expanded = new Set(
    expandPresentationGroupSelection(elements, selectedIds),
  );
  const units = new Map<string, PresentationSelectionUnit>();
  for (const element of elements) {
    if (!expanded.has(element.id)) continue;
    const groupId = topPresentationGroupId(element);
    const key = groupId ? `group:${groupId}` : `element:${element.id}`;
    const unit = units.get(key) ?? {
      key,
      groupId,
      elements: [],
      elementIds: [],
    };
    unit.elements.push(element);
    unit.elementIds.push(element.id);
    units.set(key, unit);
  }
  return [...units.values()];
}

export function presentationSelectionUnitCount(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
): number {
  return presentationSelectionUnits(elements, selectedIds).length;
}

export function canGroupPresentationElements(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
): boolean {
  return presentationSelectionUnitCount(elements, selectedIds) >= 2;
}

export function canUngroupPresentationElements(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
): boolean {
  const selected = new Set(
    expandPresentationGroupSelection(elements, selectedIds),
  );
  return elements.some(
    (element) =>
      selected.has(element.id) && presentationGroupPath(element).length > 0,
  );
}

export function groupPresentationElements(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
  groupId: string,
): WorkSlideElement[] {
  const normalizedGroupId = groupId.trim();
  if (
    !normalizedGroupId ||
    !canGroupPresentationElements(elements, selectedIds)
  ) {
    return [...elements];
  }
  const selected = new Set(
    expandPresentationGroupSelection(elements, selectedIds),
  );
  return elements.map((element) =>
    selected.has(element.id)
      ? {
          ...element,
          groupIds: [normalizedGroupId, ...presentationGroupPath(element)],
        }
      : element,
  );
}

export function ungroupPresentationElements(
  elements: readonly WorkSlideElement[],
  selectedIds: readonly string[],
): WorkSlideElement[] {
  const selected = new Set(
    expandPresentationGroupSelection(elements, selectedIds),
  );
  const selectedGroups = new Set(
    elements
      .filter((element) => selected.has(element.id))
      .map(topPresentationGroupId)
      .filter((groupId): groupId is string => Boolean(groupId)),
  );
  if (!selectedGroups.size) return [...elements];
  return elements.map((element) => {
    const path = presentationGroupPath(element);
    if (!path.length || !selectedGroups.has(path[0])) return element;
    const remaining = path.slice(1);
    return {
      ...element,
      groupIds: remaining.length ? remaining : undefined,
    };
  });
}

export function remapPresentationGroupPaths(
  elements: readonly WorkSlideElement[],
): WorkSlideElement[] {
  const replacements = new Map<string, string>();
  return elements.map((element) => {
    const path = presentationGroupPath(element);
    if (!path.length) return element;
    return {
      ...element,
      groupIds: path.map((groupId) => {
        const replacement =
          replacements.get(groupId) ?? createWorkId('element-group');
        replacements.set(groupId, replacement);
        return replacement;
      }),
    };
  });
}

export function topPresentationGroupId(
  element: WorkSlideElement,
): string | undefined {
  return presentationGroupPath(element)[0];
}

export function presentationGroupPath(element: WorkSlideElement): string[] {
  const seen = new Set<string>();
  return (element.groupIds ?? [])
    .map((groupId) => groupId.trim())
    .filter((normalized) => {
      if (!normalized || seen.has(normalized)) return false;
      seen.add(normalized);
      return true;
    });
}
