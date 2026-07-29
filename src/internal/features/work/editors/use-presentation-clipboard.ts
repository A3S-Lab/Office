import { useCallback } from 'react';
import { showToast } from '../../../state/app-state';
import {
  clonePresentationElementForPaste,
  clonePresentationElementsForPaste,
  clonePresentationSlideForPaste,
  copyPresentationElements,
  copyPresentationSlide,
  PRESENTATION_OBJECT_OFFSET_STEP,
  takePresentationClipboard,
} from '../work-presentation-clipboard';
import { withPresentationDesign } from '../work-presentation-layouts';
import type {
  WorkPresentationContent,
  WorkSlide,
  WorkSlideElement,
} from '../work-types';
import type { PresentationDesignMode } from './presentation-editor-types';
import { translatePresentationSelection } from './presentation-selection';
import {
  applyPresentationElementFormattingPatch,
  presentationElementToolbarState,
} from './presentation-text-formatting';

export function usePresentationClipboard({
  content,
  mode,
  targetId,
  selectedSlide,
  selectedElements,
  onChange,
  onSelectSlide,
  onSelectElements,
}: {
  content: WorkPresentationContent;
  mode: PresentationDesignMode;
  targetId: string | undefined;
  selectedSlide: WorkSlide | undefined;
  selectedElements: readonly WorkSlideElement[];
  onChange: (content: WorkPresentationContent) => void;
  onSelectSlide: (id: string) => void;
  onSelectElements: (ids: readonly string[]) => void;
}) {
  const selectedElement = selectedElements.at(-1) ?? null;
  const copySelection = useCallback((): boolean => {
    if (selectedElements.length) {
      copyPresentationElements(selectedElements);
      showToast(
        selectedElements.length > 1
          ? `已复制 ${selectedElements.length} 个对象`
          : '已复制演示元素',
        'success',
      );
      return true;
    }
    if (mode !== 'slide' || !selectedSlide) return false;
    copyPresentationSlide(selectedSlide);
    showToast('已复制幻灯片', 'success');
    return true;
  }, [mode, selectedElements, selectedSlide]);

  const deleteSelectedElement = useCallback((): boolean => {
    if (!selectedElements.length || !targetId) return false;
    const selectedIds = new Set(selectedElements.map((element) => element.id));
    const next = updateTargetElements(content, mode, targetId, (elements) =>
      elements.filter((element) => !selectedIds.has(element.id)),
    );
    if (!next) return false;
    onChange(next);
    onSelectElements([]);
    return true;
  }, [content, mode, onChange, onSelectElements, selectedElements, targetId]);

  const cutSelection = useCallback((): boolean => {
    if (selectedElements.length) {
      copyPresentationElements(selectedElements);
      if (!deleteSelectedElement()) return false;
      showToast(
        selectedElements.length > 1
          ? `已剪切 ${selectedElements.length} 个对象`
          : '已剪切演示元素',
        'success',
      );
      return true;
    }
    if (mode !== 'slide' || !selectedSlide) return false;
    if (content.slides.length === 1) {
      showToast('演示文稿至少需要保留一张幻灯片。', 'info');
      return true;
    }
    copyPresentationSlide(selectedSlide);
    const index = content.slides.findIndex(
      (slide) => slide.id === selectedSlide.id,
    );
    const slides = content.slides.filter(
      (slide) => slide.id !== selectedSlide.id,
    );
    onChange({ ...content, slides });
    onSelectSlide(slides[Math.min(index, slides.length - 1)].id);
    onSelectElements([]);
    showToast('已剪切幻灯片', 'success');
    return true;
  }, [
    content,
    deleteSelectedElement,
    mode,
    onChange,
    onSelectElements,
    onSelectSlide,
    selectedElements,
    selectedSlide,
  ]);

  const pasteSelection = useCallback((): boolean => {
    const clipboard = takePresentationClipboard();
    if (!clipboard) {
      showToast('没有可粘贴的演示内容。', 'info');
      return true;
    }
    if (
      clipboard.payload.kind === 'element' ||
      clipboard.payload.kind === 'elements'
    ) {
      if (!targetId) return false;
      const pasted =
        clipboard.payload.kind === 'element'
          ? [
              clonePresentationElementForPaste(
                clipboard.payload.element,
                clipboard.offset,
              ),
            ]
          : clonePresentationElementsForPaste(
              clipboard.payload.elements,
              clipboard.offset,
            );
      const next = updateTargetElements(content, mode, targetId, (elements) => [
        ...elements,
        ...pasted,
      ]);
      if (!next) return false;
      onChange(next);
      onSelectElements(pasted.map((element) => element.id));
      showToast(
        pasted.length > 1 ? `已粘贴 ${pasted.length} 个对象` : '已粘贴演示元素',
        'success',
      );
      return true;
    }
    if (mode !== 'slide' || !selectedSlide) {
      showToast('请返回幻灯片编辑后粘贴整张幻灯片。', 'info');
      return true;
    }
    const pasted = clonePresentationSlideForPaste(
      clipboard.payload.slide,
      content.slides.map((slide) => slide.name),
    );
    const index = content.slides.findIndex(
      (slide) => slide.id === selectedSlide.id,
    );
    const slides = [...content.slides];
    slides.splice(index + 1, 0, pasted);
    onChange({ ...content, slides });
    onSelectSlide(pasted.id);
    onSelectElements([]);
    showToast('已粘贴幻灯片', 'success');
    return true;
  }, [
    content,
    mode,
    onChange,
    onSelectElements,
    onSelectSlide,
    selectedSlide,
    targetId,
  ]);

  const duplicateSelection = useCallback((): boolean => {
    if (selectedElements.length && targetId) {
      const copies = clonePresentationElementsForPaste(
        selectedElements,
        PRESENTATION_OBJECT_OFFSET_STEP,
      );
      const next = updateTargetElements(content, mode, targetId, (elements) => [
        ...elements,
        ...copies,
      ]);
      if (!next) return false;
      onChange(next);
      onSelectElements(copies.map((element) => element.id));
      showToast(
        copies.length > 1 ? `已复制 ${copies.length} 个对象` : '已复制演示元素',
        'success',
      );
      return true;
    }
    if (mode !== 'slide' || !selectedSlide) return false;
    const copy = clonePresentationSlideForPaste(
      selectedSlide,
      content.slides.map((slide) => slide.name),
    );
    const index = content.slides.findIndex(
      (slide) => slide.id === selectedSlide.id,
    );
    const slides = [...content.slides];
    slides.splice(index + 1, 0, copy);
    onChange({ ...content, slides });
    onSelectSlide(copy.id);
    onSelectElements([]);
    showToast('已复制幻灯片', 'success');
    return true;
  }, [
    content,
    mode,
    onChange,
    onSelectElements,
    onSelectSlide,
    selectedElements,
    selectedSlide,
    targetId,
  ]);

  const nudgeSelection = useCallback(
    (key: string, distance: number): boolean => {
      if (!selectedElements.length || !targetId) return false;
      const horizontal =
        key === 'ArrowLeft' ? -distance : key === 'ArrowRight' ? distance : 0;
      const vertical =
        key === 'ArrowUp' ? -distance : key === 'ArrowDown' ? distance : 0;
      if (!horizontal && !vertical) return false;
      const next = updateTargetElements(content, mode, targetId, (elements) =>
        translatePresentationSelection(
          elements,
          selectedElements.map((element) => element.id),
          horizontal,
          vertical,
        ),
      );
      if (!next) return false;
      onChange(next);
      return true;
    },
    [content, mode, onChange, selectedElements, targetId],
  );

  const toggleFormatting = useCallback(
    (attribute: 'bold' | 'italic' | 'underline'): boolean => {
      if (!selectedElement || !selectedElements.length || !targetId)
        return false;
      const value =
        !presentationElementToolbarState(selectedElement)[attribute];
      const selectedIds = new Set(
        selectedElements.map((element) => element.id),
      );
      const next = updateTargetElements(content, mode, targetId, (elements) =>
        elements.map((element) =>
          selectedIds.has(element.id)
            ? applyPresentationElementFormattingPatch(
                element,
                attribute === 'bold'
                  ? { bold: value }
                  : attribute === 'italic'
                    ? { italic: value }
                    : { underline: value },
              )
            : element,
        ),
      );
      if (!next) return false;
      onChange(next);
      return true;
    },
    [content, mode, onChange, selectedElement, selectedElements, targetId],
  );

  const toggleBold = useCallback(
    () => toggleFormatting('bold'),
    [toggleFormatting],
  );
  const toggleItalic = useCallback(
    () => toggleFormatting('italic'),
    [toggleFormatting],
  );
  const toggleUnderline = useCallback(
    () => toggleFormatting('underline'),
    [toggleFormatting],
  );

  return {
    copySelection,
    cutSelection,
    deleteSelection: deleteSelectedElement,
    duplicateSelection,
    nudgeSelection,
    pasteSelection,
    toggleBold,
    toggleItalic,
    toggleUnderline,
  };
}

function updateTargetElements(
  content: WorkPresentationContent,
  mode: PresentationDesignMode,
  targetId: string,
  update: (elements: WorkSlideElement[]) => WorkSlideElement[],
): WorkPresentationContent | null {
  if (mode === 'slide') {
    if (!content.slides.some((slide) => slide.id === targetId)) return null;
    return {
      ...content,
      slides: content.slides.map((slide) =>
        slide.id === targetId
          ? { ...slide, elements: update(structuredCopy(slide.elements)) }
          : slide,
      ),
    };
  }
  const normalized = withPresentationDesign(content);
  if (mode === 'layout') {
    if (!normalized.layouts?.some((layout) => layout.id === targetId))
      return null;
    return {
      ...normalized,
      layouts: normalized.layouts.map((layout) =>
        layout.id === targetId
          ? { ...layout, elements: update(structuredCopy(layout.elements)) }
          : layout,
      ),
    };
  }
  if (!normalized.masters?.some((master) => master.id === targetId))
    return null;
  return {
    ...normalized,
    masters: normalized.masters.map((master) =>
      master.id === targetId
        ? { ...master, elements: update(structuredCopy(master.elements)) }
        : master,
    ),
  };
}

function structuredCopy<T>(value: T): T {
  if (typeof structuredClone === 'function') return structuredClone(value);
  return JSON.parse(JSON.stringify(value)) as T;
}
