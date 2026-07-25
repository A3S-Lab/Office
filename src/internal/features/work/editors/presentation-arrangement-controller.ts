import type { OfficeKernelPresentationAlignment } from '../../../kernel/office-kernel-protocol';
import type { WorkPresentationContent, WorkSlideElement } from '../work-types';
import { presentationSelectionUnits } from '../work-presentation-groups';
import type { PresentationDesignMode } from './presentation-editor-types';
import { updatePresentationElements } from './presentation-editor-operations';
import {
  alignPresentationSelection,
  distributePresentationSelection,
  presentationSelectionBounds,
  reorderPresentationSelection,
  translatePresentationSelection,
  type PresentationDistribution,
} from './presentation-selection';
import type { PresentationGeometryController } from './use-presentation-geometry';

export interface PresentationArrangementController {
  align: (alignment: OfficeKernelPresentationAlignment) => Promise<void>;
  distribute: (direction: PresentationDistribution) => void;
  reorder: (direction: -1 | 1) => void;
}

export function createPresentationArrangementController({
  getContent,
  geometry,
  mode,
  onChange,
  selectedElements,
  selectedElementIds,
  targetId,
}: {
  getContent: () => WorkPresentationContent;
  geometry: PresentationGeometryController;
  mode: PresentationDesignMode;
  onChange: (content: WorkPresentationContent) => void;
  selectedElements: readonly WorkSlideElement[];
  selectedElementIds: readonly string[];
  targetId: string | undefined;
}): PresentationArrangementController {
  return {
    align: async (alignment) => {
      const selectedElement = selectedElements.at(-1);
      if (!selectedElement || !targetId) return;
      const selectionUnits = presentationSelectionUnits(
        selectedElements,
        selectedElementIds,
      );
      if (selectionUnits.length > 1) {
        updatePresentationElements(
          getContent(),
          mode,
          targetId,
          (elements) =>
            alignPresentationSelection(elements, selectedElementIds, alignment),
          onChange,
        );
        return;
      }
      const selectionBounds = presentationSelectionBounds(selectedElements);
      if (!selectionBounds) return;
      const alignmentTarget: WorkSlideElement =
        selectedElements.length > 1
          ? {
              ...selectedElement,
              x: selectionBounds.left,
              y: selectionBounds.top,
              width: selectionBounds.width,
              height: selectionBounds.height,
            }
          : selectedElement;
      const aligned = await geometry.alignElement(alignmentTarget, alignment);
      if (!aligned) return;
      updatePresentationElements(
        getContent(),
        mode,
        targetId,
        (elements) =>
          translatePresentationSelection(
            elements,
            selectedElementIds,
            aligned.x - selectionBounds.left,
            aligned.y - selectionBounds.top,
          ),
        onChange,
      );
    },
    distribute: (direction) => {
      if (
        presentationSelectionUnits(selectedElements, selectedElementIds)
          .length < 3 ||
        !targetId
      ) {
        return;
      }
      updatePresentationElements(
        getContent(),
        mode,
        targetId,
        (elements) =>
          distributePresentationSelection(
            elements,
            selectedElementIds,
            direction,
          ),
        onChange,
      );
    },
    reorder: (direction) => {
      if (!selectedElements.length || !targetId) return;
      updatePresentationElements(
        getContent(),
        mode,
        targetId,
        (elements) =>
          reorderPresentationSelection(elements, selectedElementIds, direction),
        onChange,
      );
    },
  };
}
