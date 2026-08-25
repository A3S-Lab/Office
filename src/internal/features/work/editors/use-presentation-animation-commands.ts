import { useCallback } from 'react';
import {
  createWorkSlideAnimation,
  normalizeWorkSlideAnimation,
  WORK_SLIDE_ANIMATION_LIMIT,
  workSlideAnimationForElement,
  workSlideAnimationIndex,
} from '../work-presentation-animation';
import type {
  WorkPresentationContent,
  WorkSlide,
  WorkSlideAnimation,
  WorkSlideAnimationEffect,
} from '../work-types';
import { updateSlide } from './presentation-editor-operations';

export interface PresentationAnimationCommands {
  canMoveEntranceAnimation: (direction: -1 | 1) => boolean;
  canSetEntranceAnimation: boolean;
  canUpdateEntranceAnimation: boolean;
  moveEntranceAnimation: (direction: -1 | 1) => boolean;
  setEntranceAnimation: (
    effect: WorkSlideAnimationEffect | undefined,
  ) => boolean;
  updateEntranceAnimation: (patch: Partial<WorkSlideAnimation>) => boolean;
}

export function usePresentationAnimationCommands({
  content,
  onChange,
  selectedElementId,
  selectedSlide,
}: {
  content: WorkPresentationContent;
  onChange: (content: WorkPresentationContent) => void;
  selectedElementId: string | undefined;
  selectedSlide: WorkSlide;
}): PresentationAnimationCommands {
  const selectedElementExists = Boolean(
    selectedElementId &&
      selectedSlide.elements.some(
        (element) => element.id === selectedElementId,
      ),
  );
  const selectedAnimation = workSlideAnimationForElement(
    selectedSlide,
    selectedElementId,
  );

  const setEntranceAnimation = useCallback(
    (effect: WorkSlideAnimationEffect | undefined): boolean => {
      if (!selectedElementId || !selectedElementExists) return false;
      const current = workSlideAnimationForElement(
        selectedSlide,
        selectedElementId,
      );
      if (!effect) {
        if (!current) return false;
        updateSlide(
          content,
          selectedSlide.id,
          (slide) => {
            const animations = slide.animations?.filter(
              (animation) => animation.id !== current.id,
            );
            return {
              ...slide,
              animations: animations?.length ? animations : undefined,
            };
          },
          onChange,
        );
        return true;
      }
      if (current?.effect === effect) return false;
      if (
        !current &&
        (selectedSlide.animations?.length ?? 0) >= WORK_SLIDE_ANIMATION_LIMIT
      ) {
        return false;
      }
      updateSlide(
        content,
        selectedSlide.id,
        (slide) => {
          const next = createWorkSlideAnimation(
            selectedElementId,
            effect,
            current,
          );
          const animations = [...(slide.animations ?? [])];
          const index = current
            ? animations.findIndex((animation) => animation.id === current.id)
            : -1;
          if (index >= 0) animations[index] = next;
          else animations.push(next);
          return { ...slide, animations };
        },
        onChange,
      );
      return true;
    },
    [
      content,
      onChange,
      selectedElementExists,
      selectedElementId,
      selectedSlide,
    ],
  );

  const updateEntranceAnimation = useCallback(
    (patch: Partial<WorkSlideAnimation>): boolean => {
      if (!selectedAnimation) return false;
      const next = normalizeWorkSlideAnimation({
        ...selectedAnimation,
        ...patch,
        id: selectedAnimation.id,
        elementId: selectedAnimation.elementId,
      });
      if (animationsEqual(selectedAnimation, next)) return false;
      updateSlide(
        content,
        selectedSlide.id,
        (slide) => ({
          ...slide,
          animations: slide.animations?.map((animation) =>
            animation.id === selectedAnimation.id ? next : animation,
          ),
        }),
        onChange,
      );
      return true;
    },
    [content, onChange, selectedAnimation, selectedSlide.id],
  );

  const canMoveEntranceAnimation = useCallback(
    (direction: -1 | 1): boolean => {
      const index = workSlideAnimationIndex(
        selectedSlide,
        selectedAnimation?.id,
      );
      const target = index + direction;
      return (
        index >= 0 &&
        target >= 0 &&
        target < (selectedSlide.animations?.length ?? 0)
      );
    },
    [selectedAnimation?.id, selectedSlide],
  );

  const moveEntranceAnimation = useCallback(
    (direction: -1 | 1): boolean => {
      if (!selectedAnimation || !canMoveEntranceAnimation(direction)) {
        return false;
      }
      updateSlide(
        content,
        selectedSlide.id,
        (slide) => {
          const animations = [...(slide.animations ?? [])];
          const index = animations.findIndex(
            (animation) => animation.id === selectedAnimation.id,
          );
          const target = index + direction;
          [animations[index], animations[target]] = [
            animations[target],
            animations[index],
          ];
          return { ...slide, animations };
        },
        onChange,
      );
      return true;
    },
    [
      canMoveEntranceAnimation,
      content,
      onChange,
      selectedAnimation,
      selectedSlide.id,
    ],
  );

  return {
    canMoveEntranceAnimation,
    canSetEntranceAnimation: selectedElementExists,
    canUpdateEntranceAnimation: Boolean(selectedAnimation),
    moveEntranceAnimation,
    setEntranceAnimation,
    updateEntranceAnimation,
  };
}

function animationsEqual(
  left: WorkSlideAnimation,
  right: WorkSlideAnimation,
): boolean {
  return (
    left.id === right.id &&
    left.elementId === right.elementId &&
    left.effect === right.effect &&
    left.trigger === right.trigger &&
    left.durationMs === right.durationMs &&
    left.delayMs === right.delayMs &&
    left.direction === right.direction
  );
}
