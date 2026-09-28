import { useCallback } from 'react';
import { officeMessage, resolveOfficeMessages } from '../../../i18n/office-locale';
import { showToast } from '../../../state/app-state';
import {
  clonePresentationElementsAndAnimationsForPaste,
  clonePresentationSlideForPaste,
  copyPresentationElements,
  copyPresentationSlide,
  PRESENTATION_OBJECT_OFFSET_STEP,
  takePresentationClipboard,
} from '../work-presentation-clipboard';
import { withPresentationDesign } from '../work-presentation-layouts';
import {
  removeWorkSlideAnimationsForElements,
  WORK_SLIDE_ANIMATION_LIMIT,
} from '../work-presentation-animation';
import type {
  WorkPresentationContent,
  WorkSlide,
  WorkSlideAnimation,
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
  const messages = resolveOfficeMessages();
  const selectedElement = selectedElements.at(-1) ?? null;
  const copySelection = useCallback((): boolean => {
    if (selectedElements.length) {
      copyPresentationElements(
        selectedElements,
        mode === 'slide' ? selectedSlide?.animations : undefined,
      );
      showToast(
        selectedElements.length > 1
          ? officeMessage(messages, 'presentation.clipboard.copiedObjects', { count: String(selectedElements.length) })
          : officeMessage(messages, 'presentation.clipboard.copiedElement'),
        'success',
      );
      return true;
    }
    if (mode !== 'slide' || !selectedSlide) return false;
    copyPresentationSlide(selectedSlide);
    showToast(officeMessage(messages, 'presentation.clipboard.copiedSlide'), 'success');
    return true;
  }, [mode, selectedElements, selectedSlide]);

  const deleteSelectedElement = useCallback((): boolean => {
    if (!selectedElements.length || !targetId) return false;
    const selectedIds = new Set(selectedElements.map((element) => element.id));
    const next = updateTargetElements(
      content,
      mode,
      targetId,
      (elements) => elements.filter((element) => !selectedIds.has(element.id)),
      selectedIds,
    );
    if (!next) return false;
    onChange(next);
    onSelectElements([]);
    return true;
  }, [content, mode, onChange, onSelectElements, selectedElements, targetId]);

  const cutSelection = useCallback((): boolean => {
    if (selectedElements.length) {
      copyPresentationElements(
        selectedElements,
        mode === 'slide' ? selectedSlide?.animations : undefined,
      );
      if (!deleteSelectedElement()) return false;
      showToast(
        selectedElements.length > 1
          ? officeMessage(messages, 'presentation.clipboard.cutObjects', { count: String(selectedElements.length) })
          : officeMessage(messages, 'presentation.clipboard.cutElement'),
        'success',
      );
      return true;
    }
    if (mode !== 'slide' || !selectedSlide) return false;
    if (content.slides.length === 1) {
      showToast(officeMessage(messages, 'presentation.clipboard.keepOneSlide'), 'info');
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
    showToast(officeMessage(messages, 'presentation.clipboard.cutSlide'), 'success');
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
      showToast(officeMessage(messages, 'presentation.clipboard.empty'), 'info');
      return true;
    }
    if (clipboard.payload.kind === 'elements') {
      if (!targetId) return false;
      const pasted = clonePresentationElementsAndAnimationsForPaste(
        clipboard.payload.elements,
        clipboard.payload.animations ?? [],
        clipboard.offset,
      );
      if (
        mode === 'slide' &&
        !canAppendSlideAnimations(content, targetId, pasted.animations)
      ) {
        showToast(officeMessage(messages, 'presentation.clipboard.animationLimit'), 'info');
        return true;
      }
      const next = updateTargetElements(
        content,
        mode,
        targetId,
        (elements) => [...elements, ...pasted.elements],
        undefined,
        pasted.animations,
      );
      if (!next) return false;
      onChange(next);
      onSelectElements(pasted.elements.map((element) => element.id));
      showToast(
        pasted.elements.length > 1
          ? officeMessage(messages, 'presentation.clipboard.pastedObjects', { count: String(pasted.elements.length) })
          : officeMessage(messages, 'presentation.clipboard.pastedElement'),
        'success',
      );
      return true;
    }
    if (mode !== 'slide' || !selectedSlide) {
      showToast(officeMessage(messages, 'presentation.clipboard.pasteSlideNeedsSlideMode'), 'info');
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
    showToast(officeMessage(messages, 'presentation.clipboard.pastedSlide'), 'success');
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
      const copies = clonePresentationElementsAndAnimationsForPaste(
        selectedElements,
        mode === 'slide' ? (selectedSlide?.animations ?? []) : [],
        PRESENTATION_OBJECT_OFFSET_STEP,
      );
      if (
        mode === 'slide' &&
        !canAppendSlideAnimations(content, targetId, copies.animations)
      ) {
        showToast(officeMessage(messages, 'presentation.clipboard.animationLimit'), 'info');
        return true;
      }
      const next = updateTargetElements(
        content,
        mode,
        targetId,
        (elements) => [...elements, ...copies.elements],
        undefined,
        copies.animations,
      );
      if (!next) return false;
      onChange(next);
      onSelectElements(copies.elements.map((element) => element.id));
      showToast(
        copies.elements.length > 1
          ? officeMessage(messages, 'presentation.clipboard.copiedObjects', { count: String(copies.elements.length) })
          : officeMessage(messages, 'presentation.clipboard.copiedElement'),
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
    showToast(officeMessage(messages, 'presentation.clipboard.copiedSlide'), 'success');
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
  removedElementIds?: ReadonlySet<string>,
  addedAnimations?: readonly WorkSlideAnimation[],
): WorkPresentationContent | null {
  if (mode === 'slide') {
    if (!content.slides.some((slide) => slide.id === targetId)) return null;
    return {
      ...content,
      slides: content.slides.map((slide) =>
        slide.id === targetId
          ? appendWorkSlideAnimations(
              removeWorkSlideAnimationsForElements(
                { ...slide, elements: update(structuredCopy(slide.elements)) },
                removedElementIds ?? new Set(),
              ),
              addedAnimations,
            )
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

function canAppendSlideAnimations(
  content: WorkPresentationContent,
  slideId: string,
  animations: readonly WorkSlideAnimation[] | undefined,
): boolean {
  if (!animations?.length) return true;
  const slide = content.slides.find((candidate) => candidate.id === slideId);
  return (
    Boolean(slide) &&
    (slide?.animations?.length ?? 0) + animations.length <=
      WORK_SLIDE_ANIMATION_LIMIT
  );
}

function appendWorkSlideAnimations(
  slide: WorkSlide,
  animations: readonly WorkSlideAnimation[] | undefined,
): WorkSlide {
  if (!animations?.length) return slide;
  return {
    ...slide,
    animations: [...(slide.animations ?? []), ...structuredCopy(animations)],
  };
}

function structuredCopy<T>(value: T): T {
  if (typeof structuredClone === 'function') return structuredClone(value);
  return JSON.parse(JSON.stringify(value)) as T;
}
