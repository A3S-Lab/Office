import { useCallback, type MutableRefObject } from 'react';
import { createPresentationChartElement } from '../work-presentation-charts';
import {
  canGroupPresentationElements,
  canUngroupPresentationElements,
  groupPresentationElements,
  ungroupPresentationElements,
} from '../work-presentation-groups';
import { createWorkId } from '../work-templates';
import type {
  WorkPresentationContent,
  WorkSlide,
  WorkSlideElement,
} from '../work-types';
import type { PresentationDesignMode } from './presentation-editor-types';
import type { OfficeTableDimensions } from './office-table-dimensions';
import {
  newPresentationElement,
  newPresentationImageElement,
  newPresentationTableElement,
  structuredCopy,
  updatePresentationElements,
} from './presentation-editor-operations';
import type { PresentationTextValue } from './presentation-text-editor';
import { applyPresentationElementFormattingPatch } from './presentation-text-formatting';

export interface PresentationElementCommands {
  addChart: () => void;
  addElement: (type: 'shape' | 'text') => void;
  addImage: (file: File) => Promise<void>;
  addTable: (dimensions: OfficeTableDimensions) => void;
  groupSelection: () => boolean;
  instantiatePlaceholder: (definition: WorkSlideElement) => void;
  ungroupSelection: () => boolean;
  updateElement: (patch: Partial<WorkSlideElement>) => void;
  updateTextElement: (elementId: string, value: PresentationTextValue) => void;
}

export function usePresentationElementCommands({
  activeElements,
  activeTargetId,
  content,
  contentRef,
  designMode,
  onChange,
  onEditElement,
  onSelectElements,
  selectedElementIds,
  selectedElements,
  selectedSlide,
}: {
  activeElements: readonly WorkSlideElement[];
  activeTargetId: string | undefined;
  content: WorkPresentationContent;
  contentRef: MutableRefObject<WorkPresentationContent>;
  designMode: PresentationDesignMode;
  onChange: (content: WorkPresentationContent) => void;
  onEditElement: (id: string) => void;
  onSelectElements: (ids: readonly string[]) => void;
  selectedElementIds: readonly string[];
  selectedElements: readonly WorkSlideElement[];
  selectedSlide: WorkSlide;
}): PresentationElementCommands {
  const groupSelection = useCallback((): boolean => {
    if (
      !activeTargetId ||
      !canGroupPresentationElements(activeElements, selectedElementIds)
    ) {
      return false;
    }
    updatePresentationElements(
      contentRef.current,
      designMode,
      activeTargetId,
      (elements) =>
        groupPresentationElements(
          elements,
          selectedElementIds,
          createWorkId('element-group'),
        ),
      (next) => {
        contentRef.current = next;
        onChange(next);
      },
    );
    return true;
  }, [
    activeElements,
    activeTargetId,
    contentRef,
    designMode,
    onChange,
    selectedElementIds,
  ]);

  const ungroupSelection = useCallback((): boolean => {
    if (
      !activeTargetId ||
      !canUngroupPresentationElements(activeElements, selectedElementIds)
    ) {
      return false;
    }
    updatePresentationElements(
      contentRef.current,
      designMode,
      activeTargetId,
      (elements) => ungroupPresentationElements(elements, selectedElementIds),
      (next) => {
        contentRef.current = next;
        onChange(next);
      },
    );
    return true;
  }, [
    activeElements,
    activeTargetId,
    contentRef,
    designMode,
    onChange,
    selectedElementIds,
  ]);

  const updateElement = useCallback(
    (patch: Partial<WorkSlideElement>) => {
      if (!selectedElements.length || !activeTargetId) return;
      const selectedIds = new Set(
        selectedElements.map((element) => element.id),
      );
      updatePresentationElements(
        contentRef.current,
        designMode,
        activeTargetId,
        (elements) =>
          elements.map((element) =>
            selectedIds.has(element.id)
              ? applyPresentationElementFormattingPatch(element, patch)
              : element,
          ),
        (next) => {
          contentRef.current = next;
          onChange(next);
        },
      );
    },
    [activeTargetId, contentRef, designMode, onChange, selectedElements],
  );

  const updateTextElement = useCallback(
    (elementId: string, value: PresentationTextValue) => {
      if (!activeTargetId) return;
      updatePresentationElements(
        contentRef.current,
        designMode,
        activeTargetId,
        (elements) =>
          elements.map((element) =>
            element.id === elementId ? { ...element, ...value } : element,
          ),
        (next) => {
          contentRef.current = next;
          onChange(next);
        },
      );
    },
    [activeTargetId, contentRef, designMode, onChange],
  );

  const addElement = useCallback(
    (type: 'shape' | 'text') => {
      const element = newPresentationElement(type);
      if (!activeTargetId) return;
      updatePresentationElements(
        content,
        designMode,
        activeTargetId,
        (elements) => [...elements, element],
        onChange,
      );
      if (type === 'text') onEditElement(element.id);
      else onSelectElements([element.id]);
    },
    [
      activeTargetId,
      content,
      designMode,
      onChange,
      onEditElement,
      onSelectElements,
    ],
  );

  const addTable = useCallback(
    (dimensions: OfficeTableDimensions) => {
      const element = newPresentationTableElement(
        dimensions.rows,
        dimensions.columns,
      );
      updatePresentationElements(
        content,
        designMode,
        activeTargetId ?? selectedSlide.id,
        (elements) => [...elements, element],
        onChange,
      );
      onSelectElements([element.id]);
    },
    [
      activeTargetId,
      content,
      designMode,
      onChange,
      onSelectElements,
      selectedSlide.id,
    ],
  );

  const addChart = useCallback(() => {
    const element = createPresentationChartElement();
    updatePresentationElements(
      content,
      'slide',
      selectedSlide.id,
      (elements) => [...elements, element],
      onChange,
    );
    onSelectElements([element.id]);
  }, [content, onChange, onSelectElements, selectedSlide.id]);

  const addImage = useCallback(
    async (file: File) => {
      const element = await newPresentationImageElement(file);
      if (!activeTargetId) return;
      updatePresentationElements(
        content,
        designMode,
        activeTargetId,
        (elements) => [...elements, element],
        onChange,
      );
      onSelectElements([element.id]);
    },
    [activeTargetId, content, designMode, onChange, onSelectElements],
  );

  const instantiatePlaceholder = useCallback(
    (definition: WorkSlideElement) => {
      if (!activeTargetId || designMode !== 'slide') return;
      const element: WorkSlideElement = {
        ...structuredCopy(definition),
        id: createWorkId('element'),
        text: '',
        textRuns: undefined,
      };
      updatePresentationElements(
        content,
        'slide',
        activeTargetId,
        (elements) => [...elements, element],
        onChange,
      );
      onEditElement(element.id);
    },
    [activeTargetId, content, designMode, onChange, onEditElement],
  );

  return {
    addChart,
    addElement,
    addImage,
    addTable,
    groupSelection,
    instantiatePlaceholder,
    ungroupSelection,
    updateElement,
    updateTextElement,
  };
}
