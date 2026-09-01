import { useCallback } from 'react';
import {
  createWorkSlideAnimation,
  normalizeWorkSlideAnimation,
  WORK_SLIDE_ANIMATION_LIMIT,
  workSlideAnimationEffectMatchesClass,
  workSlideAnimationForElement,
  workSlideAnimationIndex,
  workSlideAnimationSequenceIssue,
} from '../work-presentation-animation';
import type {
  WorkPresentationContent,
  WorkSlide,
  WorkSlideAnimation,
  WorkSlideAnimationClass,
  WorkSlideAnimationEffect,
} from '../work-types';
import { updateSlide } from './presentation-editor-operations';

export interface PresentationAnimationCommands {
  canMoveAnimation: (
    animationClass: WorkSlideAnimationClass,
    direction: -1 | 1,
  ) => boolean;
  canSetAnimation: boolean;
  canUpdateAnimation: (
    animationClass: WorkSlideAnimationClass,
    patch: Partial<WorkSlideAnimation>,
  ) => boolean;
  moveAnimation: (
    animationClass: WorkSlideAnimationClass,
    direction: -1 | 1,
  ) => boolean;
  setAnimation: (
    animationClass: WorkSlideAnimationClass,
    effect: WorkSlideAnimationEffect | undefined,
  ) => boolean;
  updateAnimation: (
    animationClass: WorkSlideAnimationClass,
    patch: Partial<WorkSlideAnimation>,
  ) => boolean;
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
  const setAnimation = useCallback(
    (
      animationClass: WorkSlideAnimationClass,
      effect: WorkSlideAnimationEffect | undefined,
    ): boolean => {
      if (!selectedElementId || !selectedElementExists) return false;
      const current = workSlideAnimationForElement(
        selectedSlide,
        selectedElementId,
        animationClass,
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
      if (!workSlideAnimationEffectMatchesClass(effect, animationClass)) {
        return false;
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

  const canUpdateAnimation = useCallback(
    (
      animationClass: WorkSlideAnimationClass,
      patch: Partial<WorkSlideAnimation>,
    ): boolean => {
      const selectedAnimation = workSlideAnimationForElement(
        selectedSlide,
        selectedElementId,
        animationClass,
      );
      if (!selectedAnimation) return false;
      const next = normalizeWorkSlideAnimation({
        ...selectedAnimation,
        ...patch,
        id: selectedAnimation.id,
        elementId: selectedAnimation.elementId,
        effect: selectedAnimation.effect,
      });
      const animations = selectedSlide.animations?.map((animation) =>
        animation.id === selectedAnimation.id ? next : animation,
      );
      return workSlideAnimationSequenceIssue(animations) === undefined;
    },
    [selectedElementId, selectedSlide],
  );

  const updateAnimation = useCallback(
    (
      animationClass: WorkSlideAnimationClass,
      patch: Partial<WorkSlideAnimation>,
    ): boolean => {
      const selectedAnimation = workSlideAnimationForElement(
        selectedSlide,
        selectedElementId,
        animationClass,
      );
      if (!selectedAnimation || !canUpdateAnimation(animationClass, patch)) {
        return false;
      }
      const next = normalizeWorkSlideAnimation({
        ...selectedAnimation,
        ...patch,
        id: selectedAnimation.id,
        elementId: selectedAnimation.elementId,
        effect: selectedAnimation.effect,
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
    [canUpdateAnimation, content, onChange, selectedElementId, selectedSlide],
  );

  const canMoveAnimation = useCallback(
    (animationClass: WorkSlideAnimationClass, direction: -1 | 1): boolean => {
      const selectedAnimation = workSlideAnimationForElement(
        selectedSlide,
        selectedElementId,
        animationClass,
      );
      const index = workSlideAnimationIndex(
        selectedSlide,
        selectedAnimation?.id,
      );
      const target = index + direction;
      if (
        index < 0 ||
        target < 0 ||
        target >= (selectedSlide.animations?.length ?? 0)
      ) {
        return false;
      }
      const animations = [...(selectedSlide.animations ?? [])];
      [animations[index], animations[target]] = [
        animations[target],
        animations[index],
      ];
      return workSlideAnimationSequenceIssue(animations) === undefined;
    },
    [selectedElementId, selectedSlide],
  );

  const moveAnimation = useCallback(
    (animationClass: WorkSlideAnimationClass, direction: -1 | 1): boolean => {
      const selectedAnimation = workSlideAnimationForElement(
        selectedSlide,
        selectedElementId,
        animationClass,
      );
      if (!selectedAnimation || !canMoveAnimation(animationClass, direction)) {
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
    [canMoveAnimation, content, onChange, selectedElementId, selectedSlide],
  );

  return {
    canMoveAnimation,
    canSetAnimation: selectedElementExists,
    canUpdateAnimation,
    moveAnimation,
    setAnimation,
    updateAnimation,
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
