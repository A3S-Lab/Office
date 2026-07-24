import { useCallback, useEffect, useState } from 'react';
import type { WorkSlideElement } from '../work-types';
import { selectPresentationElement } from './presentation-selection';

export interface PresentationSelectionController {
  clear: () => void;
  edit: (elementId: string) => void;
  editingElementId: string | null;
  exitEditing: () => void;
  replace: (elementIds: readonly string[]) => void;
  select: (elementId: string | null, additive?: boolean) => void;
  selectedElementIds: string[];
}

export function usePresentationSelection(
  elements: readonly WorkSlideElement[],
): PresentationSelectionController {
  const [selectedElementIds, setSelectedElementIds] = useState<string[]>([]);
  const [editingElementId, setEditingElementId] = useState<string | null>(null);

  const clear = useCallback(() => {
    setSelectedElementIds([]);
    setEditingElementId(null);
  }, []);

  const edit = useCallback((elementId: string) => {
    setSelectedElementIds([elementId]);
    setEditingElementId(elementId);
  }, []);

  const exitEditing = useCallback(() => setEditingElementId(null), []);

  const replace = useCallback((elementIds: readonly string[]) => {
    setSelectedElementIds([...new Set(elementIds)]);
    setEditingElementId(null);
  }, []);

  const select = useCallback(
    (elementId: string | null, additive = false) => {
      if (!elementId) {
        clear();
        return;
      }
      setSelectedElementIds((current) =>
        selectPresentationElement(current, elementId, additive),
      );
      setEditingElementId(null);
    },
    [clear],
  );

  useEffect(() => {
    const validIds = new Set(elements.map((element) => element.id));
    setSelectedElementIds((current) => {
      const next = current.filter((id) => validIds.has(id));
      return sameStringArray(current, next) ? current : next;
    });
    setEditingElementId((current) =>
      current && validIds.has(current) ? current : null,
    );
  }, [elements]);

  return {
    clear,
    edit,
    editingElementId,
    exitEditing,
    replace,
    select,
    selectedElementIds,
  };
}

function sameStringArray(left: readonly string[], right: readonly string[]) {
  return (
    left.length === right.length &&
    left.every((value, index) => value === right[index])
  );
}
