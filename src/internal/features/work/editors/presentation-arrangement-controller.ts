import type { OfficeKernelPresentationAlignment } from '../../../kernel/office-kernel-protocol';
import type { WorkPresentationContent, WorkSlideElement } from '../work-types';
import type { PresentationDesignMode } from './presentation-design-panel';
import { updatePresentationElements } from './presentation-editor-operations';
import {
  alignPresentationSelection,
  distributePresentationSelection,
  reorderPresentationSelection,
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
      if (selectedElements.length > 1) {
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
      const aligned = await geometry.alignElement(selectedElement, alignment);
      if (!aligned) return;
      updatePresentationElements(
        getContent(),
        mode,
        targetId,
        (elements) =>
          elements.map((element) =>
            element.id === selectedElement.id
              ? { ...element, x: aligned.x, y: aligned.y }
              : element,
          ),
        onChange,
      );
    },
    distribute: (direction) => {
      if (selectedElements.length < 3 || !targetId) return;
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
