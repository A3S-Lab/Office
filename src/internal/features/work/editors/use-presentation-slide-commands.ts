import { useCallback } from 'react';
import { clonePresentationSlideForPaste } from '../work-presentation-clipboard';
import { workSlideTransitionsEqual } from '../work-presentation-transition';
import type {
  WorkPresentationContent,
  WorkSlide,
  WorkSlideTransition,
} from '../work-types';
import {
  newSlide,
  structuredCopy,
  updateSlide,
} from './presentation-editor-operations';

export interface PresentationSlideCommands {
  addSlide: () => void;
  applyTransitionToAll: (
    transition: WorkSlideTransition | undefined,
  ) => boolean;
  canApplyTransitionToAll: (
    transition: WorkSlideTransition | undefined,
  ) => boolean;
  deleteSlide: () => boolean;
  deleteSlideById: (slideId: string) => boolean;
  duplicateSlide: () => void;
  setTransition: (transition: WorkSlideTransition | undefined) => boolean;
  updateNotes: (notes: string) => void;
}

export function usePresentationSlideCommands({
  content,
  onChange,
  onClearSelection,
  onSelectSlide,
  selectedSlide,
}: {
  content: WorkPresentationContent;
  onChange: (content: WorkPresentationContent) => void;
  onClearSelection: () => void;
  onSelectSlide: (id: string) => void;
  selectedSlide: WorkSlide;
}): PresentationSlideCommands {
  const addSlide = useCallback(() => {
    const slide = newSlide(content.slides.length + 1);
    onChange({ ...content, slides: [...content.slides, slide] });
    onSelectSlide(slide.id);
    onClearSelection();
  }, [content, onChange, onClearSelection, onSelectSlide]);

  const duplicateSlide = useCallback(() => {
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
    onClearSelection();
  }, [content, onChange, onClearSelection, onSelectSlide, selectedSlide]);

  const deleteSlideById = useCallback(
    (slideId: string): boolean => {
      if (content.slides.length === 1) return false;
      const index = content.slides.findIndex((slide) => slide.id === slideId);
      if (index < 0) return false;
      const slides = content.slides.filter((slide) => slide.id !== slideId);
      onChange({ ...content, slides });
      onSelectSlide(slides[Math.min(index, slides.length - 1)].id);
      onClearSelection();
      return true;
    },
    [content, onChange, onClearSelection, onSelectSlide],
  );

  const deleteSlide = useCallback(
    () => deleteSlideById(selectedSlide.id),
    [deleteSlideById, selectedSlide.id],
  );

  const canApplyTransitionToAll = useCallback(
    (transition: WorkSlideTransition | undefined) =>
      content.slides.some(
        (slide) => !workSlideTransitionsEqual(slide.transition, transition),
      ),
    [content.slides],
  );

  const applyTransitionToAll = useCallback(
    (transition: WorkSlideTransition | undefined): boolean => {
      if (!canApplyTransitionToAll(transition)) return false;
      onChange({
        ...content,
        slides: content.slides.map((slide) => ({
          ...slide,
          transition: transition ? structuredCopy(transition) : undefined,
        })),
      });
      return true;
    },
    [canApplyTransitionToAll, content, onChange],
  );

  const setTransition = useCallback(
    (transition: WorkSlideTransition | undefined) => {
      if (workSlideTransitionsEqual(selectedSlide.transition, transition)) {
        return false;
      }
      updateSlide(
        content,
        selectedSlide.id,
        (slide) => ({ ...slide, transition }),
        onChange,
      );
      return true;
    },
    [content, onChange, selectedSlide.id, selectedSlide.transition],
  );

  const updateNotes = useCallback(
    (notes: string) => {
      updateSlide(
        content,
        selectedSlide.id,
        (slide) => ({ ...slide, notes }),
        onChange,
      );
    },
    [content, onChange, selectedSlide.id],
  );

  return {
    addSlide,
    applyTransitionToAll,
    canApplyTransitionToAll,
    deleteSlide,
    deleteSlideById,
    duplicateSlide,
    setTransition,
    updateNotes,
  };
}
