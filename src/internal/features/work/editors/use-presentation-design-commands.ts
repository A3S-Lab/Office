import { useCallback } from 'react';
import { applyPresentationLayout } from '../work-presentation-layouts';
import { remapPresentationGroupPaths } from '../work-presentation-groups';
import { createWorkId } from '../work-templates';
import type {
  WorkPresentationContent,
  WorkPresentationLayout,
  WorkPresentationMaster,
  WorkSlide,
  WorkSlideElement,
} from '../work-types';
import type { PresentationDesignMode } from './presentation-editor-types';
import {
  structuredCopy,
  updatePresentationElements,
  updateSlide,
} from './presentation-editor-operations';

export interface PresentationDesignCommands {
  addPlaceholder: (type: 'body' | 'title') => void;
  applyLayout: (layoutId: string) => void;
  close: () => void;
  createLayout: (copyCurrent: boolean) => void;
  deleteLayout: () => void;
  edit: (mode: PresentationDesignMode) => void;
  renameLayout: (name: string) => void;
  renameMaster: (name: string) => void;
  setActiveBackground: (background: string) => void;
  setLayoutBackground: (background: string | undefined) => void;
  setMasterBackground: (background: string) => void;
  toggleDesignPanel: () => void;
  toggleLayoutBackground: (enabled: boolean) => void;
}

export function usePresentationDesignCommands({
  activeElements,
  activeTargetId,
  designContent,
  designMode,
  designOpen,
  onChange,
  onClearSelection,
  onCloseComments,
  onEditElement,
  onSetDesignMode,
  onSetDesignOpen,
  selectedLayout,
  selectedMaster,
  selectedSlide,
}: {
  activeElements: readonly WorkSlideElement[];
  activeTargetId: string | undefined;
  designContent: WorkPresentationContent;
  designMode: PresentationDesignMode;
  designOpen: boolean;
  onChange: (content: WorkPresentationContent) => void;
  onClearSelection: () => void;
  onCloseComments: () => void;
  onEditElement: (id: string) => void;
  onSetDesignMode: (mode: PresentationDesignMode) => void;
  onSetDesignOpen: (open: boolean) => void;
  selectedLayout: WorkPresentationLayout | undefined;
  selectedMaster: WorkPresentationMaster | undefined;
  selectedSlide: WorkSlide;
}): PresentationDesignCommands {
  const updateLayout = useCallback(
    (
      layoutId: string,
      update: (layout: WorkPresentationLayout) => WorkPresentationLayout,
    ) => {
      onChange({
        ...designContent,
        layouts: designContent.layouts?.map((layout) =>
          layout.id === layoutId ? update(structuredCopy(layout)) : layout,
        ),
      });
    },
    [designContent, onChange],
  );

  const updateMaster = useCallback(
    (
      masterId: string,
      update: (master: WorkPresentationMaster) => WorkPresentationMaster,
    ) => {
      onChange({
        ...designContent,
        masters: designContent.masters?.map((master) =>
          master.id === masterId ? update(structuredCopy(master)) : master,
        ),
      });
    },
    [designContent, onChange],
  );

  const toggleDesignPanel = useCallback(() => {
    if (designOpen) {
      onSetDesignOpen(false);
      onSetDesignMode('slide');
      onClearSelection();
      return;
    }
    onChange(designContent);
    onSetDesignOpen(true);
    onCloseComments();
  }, [
    designContent,
    designOpen,
    onChange,
    onClearSelection,
    onCloseComments,
    onSetDesignMode,
    onSetDesignOpen,
  ]);

  const applyLayout = useCallback(
    (layoutId: string) => {
      onChange(
        applyPresentationLayout(designContent, selectedSlide.id, layoutId),
      );
      onSetDesignMode('slide');
      onClearSelection();
    },
    [
      designContent,
      onChange,
      onClearSelection,
      onSetDesignMode,
      selectedSlide.id,
    ],
  );

  const toggleLayoutBackground = useCallback(
    (enabled: boolean) => {
      updateSlide(
        designContent,
        selectedSlide.id,
        (slide) => ({ ...slide, useLayoutBackground: enabled }),
        onChange,
      );
    },
    [designContent, onChange, selectedSlide.id],
  );

  const edit = useCallback(
    (mode: PresentationDesignMode) => {
      onSetDesignMode(mode);
      onClearSelection();
    },
    [onClearSelection, onSetDesignMode],
  );

  const close = useCallback(() => {
    onSetDesignOpen(false);
    onSetDesignMode('slide');
    onClearSelection();
  }, [onClearSelection, onSetDesignMode, onSetDesignOpen]);

  const setActiveBackground = useCallback(
    (background: string) => {
      if (designMode === 'layout' && selectedLayout) {
        updateLayout(selectedLayout.id, (layout) => ({
          ...layout,
          background,
        }));
        return;
      }
      if (designMode === 'master' && selectedMaster) {
        updateMaster(selectedMaster.id, (master) => ({
          ...master,
          background,
        }));
        return;
      }
      updateSlide(
        designContent,
        selectedSlide.id,
        (slide) => ({ ...slide, background, useLayoutBackground: false }),
        onChange,
      );
    },
    [
      designContent,
      designMode,
      onChange,
      selectedLayout,
      selectedMaster,
      selectedSlide.id,
      updateLayout,
      updateMaster,
    ],
  );

  const createLayout = useCallback(
    (copyCurrent: boolean) => {
      if (!selectedMaster || !selectedLayout) return;
      const id = createWorkId('layout');
      const layout: WorkPresentationLayout = copyCurrent
        ? {
            ...structuredCopy(selectedLayout),
            id,
            name: `${selectedLayout.name} 副本`,
            elements: remapPresentationGroupPaths(selectedLayout.elements).map(
              (element) => ({
                ...structuredCopy(element),
                id: createWorkId('element'),
              }),
            ),
          }
        : {
            id,
            name: `自定义布局 ${(designContent.layouts?.length ?? 0) + 1}`,
            masterId: selectedMaster.id,
            elements: [],
          };
      const next = applyPresentationLayout(
        {
          ...designContent,
          layouts: [...(designContent.layouts ?? []), layout],
        },
        selectedSlide.id,
        id,
      );
      onChange(next);
      onSetDesignMode('layout');
      onClearSelection();
    },
    [
      designContent,
      onChange,
      onClearSelection,
      onSetDesignMode,
      selectedLayout,
      selectedMaster,
      selectedSlide.id,
    ],
  );

  const deleteLayout = useCallback(() => {
    if (!selectedLayout || (designContent.layouts?.length ?? 0) < 2) return;
    const fallback = designContent.layouts?.find(
      (layout) => layout.id !== selectedLayout.id,
    );
    if (!fallback) return;
    onChange({
      ...designContent,
      layouts: designContent.layouts?.filter(
        (layout) => layout.id !== selectedLayout.id,
      ),
      slides: designContent.slides.map((slide) =>
        slide.layoutId === selectedLayout.id
          ? { ...slide, layoutId: fallback.id }
          : slide,
      ),
    });
    onSetDesignMode('slide');
    onClearSelection();
  }, [
    designContent,
    onChange,
    onClearSelection,
    onSetDesignMode,
    selectedLayout,
  ]);

  const addPlaceholder = useCallback(
    (type: 'title' | 'body') => {
      if (designMode === 'slide' || !activeTargetId) return;
      const count = activeElements.filter(
        (element) => element.placeholder?.type === type,
      ).length;
      const prompt = type === 'title' ? '单击添加标题' : '单击添加内容';
      const element: WorkSlideElement = {
        id: createWorkId('element'),
        type: 'text',
        x: type === 'title' ? 8 : 10,
        y: type === 'title' ? 9 : 24,
        width: type === 'title' ? 84 : 80,
        height: type === 'title' ? 12 : 58,
        text: prompt,
        fontSize: type === 'title' ? 30 : 20,
        color: '#172033',
        fill: 'transparent',
        bold: type === 'title',
        align: 'left',
        placeholder: {
          key: count ? `type:${type}:${count + 1}` : `type:${type}`,
          type,
          prompt,
        },
      };
      updatePresentationElements(
        designContent,
        designMode,
        activeTargetId,
        (elements) => [...elements, element],
        onChange,
      );
      onEditElement(element.id);
    },
    [
      activeElements,
      activeTargetId,
      designContent,
      designMode,
      onChange,
      onEditElement,
    ],
  );

  const renameLayout = useCallback(
    (name: string) => {
      if (!selectedLayout) return;
      updateLayout(selectedLayout.id, (layout) => ({ ...layout, name }));
    },
    [selectedLayout, updateLayout],
  );

  const renameMaster = useCallback(
    (name: string) => {
      if (!selectedMaster) return;
      updateMaster(selectedMaster.id, (master) => ({ ...master, name }));
    },
    [selectedMaster, updateMaster],
  );

  const setLayoutBackground = useCallback(
    (background: string | undefined) => {
      if (!selectedLayout) return;
      updateLayout(selectedLayout.id, (layout) => ({
        ...layout,
        background,
      }));
    },
    [selectedLayout, updateLayout],
  );

  const setMasterBackground = useCallback(
    (background: string) => {
      if (!selectedMaster) return;
      updateMaster(selectedMaster.id, (master) => ({
        ...master,
        background,
      }));
    },
    [selectedMaster, updateMaster],
  );

  return {
    addPlaceholder,
    applyLayout,
    close,
    createLayout,
    deleteLayout,
    edit,
    renameLayout,
    renameMaster,
    setActiveBackground,
    setLayoutBackground,
    setMasterBackground,
    toggleDesignPanel,
    toggleLayoutBackground,
  };
}
