import { createWorkId } from './work-templates';
import {
  WORK_SLIDE_ANIMATION_MAX_DELAY_MS,
  WORK_SLIDE_ANIMATION_MAX_DURATION_MS,
  WORK_SLIDE_ANIMATION_MIN_DURATION_MS,
} from './work-presentation-animation-constraints';
import type {
  WorkSlide,
  WorkSlideAnimation,
  WorkSlideAnimationClass,
  WorkSlideAnimationDirection,
  WorkSlideAnimationEffect,
} from './work-types';

export {
  WORK_SLIDE_ANIMATION_LIMIT,
  WORK_SLIDE_ANIMATION_MAX_DELAY_MS,
  WORK_SLIDE_ANIMATION_MAX_DURATION_MS,
  WORK_SLIDE_ANIMATION_MIN_DURATION_MS,
} from './work-presentation-animation-constraints';

export interface WorkSlideAnimationCueItem {
  animation: WorkSlideAnimation;
  endOffsetMs: number;
  startOffsetMs: number;
}

export interface WorkSlideAnimationCue {
  automatic: boolean;
  items: WorkSlideAnimationCueItem[];
  totalDurationMs: number;
}

export type WorkSlideAnimationSequenceIssue =
  | 'duplicate-class'
  | 'overlapping-target';

export function createWorkSlideAnimation(
  elementId: string,
  effect: WorkSlideAnimationEffect,
  previous?: WorkSlideAnimation,
): WorkSlideAnimation {
  return normalizeWorkSlideAnimation({
    id: previous?.id ?? createWorkId('slide-animation'),
    elementId,
    effect,
    trigger: previous?.trigger ?? 'on-click',
    durationMs: previous?.durationMs ?? 500,
    delayMs: previous?.delayMs ?? 0,
    direction: isFlyAnimationEffect(effect)
      ? (previous?.direction ?? 'left')
      : undefined,
  });
}

export function normalizeWorkSlideAnimation(
  animation: WorkSlideAnimation,
): WorkSlideAnimation {
  return {
    ...animation,
    durationMs: boundedInteger(
      animation.durationMs,
      WORK_SLIDE_ANIMATION_MIN_DURATION_MS,
      WORK_SLIDE_ANIMATION_MAX_DURATION_MS,
    ),
    delayMs: boundedInteger(
      animation.delayMs,
      0,
      WORK_SLIDE_ANIMATION_MAX_DELAY_MS,
    ),
    direction: isFlyAnimationEffect(animation.effect)
      ? normalizeWorkSlideAnimationDirection(animation.direction)
      : undefined,
  };
}

export function workSlideAnimationClass(
  effect: WorkSlideAnimationEffect,
): WorkSlideAnimationClass {
  return effect === 'disappear' ||
    effect === 'fade-out' ||
    effect === 'fly-out' ||
    effect === 'zoom-out'
    ? 'exit'
    : 'entrance';
}

export function workSlideAnimationEffectMatchesClass(
  effect: WorkSlideAnimationEffect,
  animationClass: WorkSlideAnimationClass,
): boolean {
  return workSlideAnimationClass(effect) === animationClass;
}

export function workSlideAnimationForElement(
  slide: WorkSlide,
  elementId: string | undefined,
  animationClass: WorkSlideAnimationClass,
): WorkSlideAnimation | undefined {
  if (!elementId) return undefined;
  return slide.animations?.find(
    (animation) =>
      animation.elementId === elementId &&
      workSlideAnimationClass(animation.effect) === animationClass,
  );
}

export function workSlideAnimationIndex(
  slide: WorkSlide,
  animationId: string | undefined,
): number {
  if (!animationId) return -1;
  return (
    slide.animations?.findIndex((animation) => animation.id === animationId) ??
    -1
  );
}

export function workSlideAnimationCues(
  animations: readonly WorkSlideAnimation[] | undefined,
): WorkSlideAnimationCue[] {
  const cues: WorkSlideAnimationCue[] = [];
  let current: WorkSlideAnimationCue | undefined;
  for (const raw of animations ?? []) {
    const animation = normalizeWorkSlideAnimation(raw);
    if (!current || animation.trigger === 'on-click') {
      current = {
        automatic: animation.trigger !== 'on-click',
        items: [],
        totalDurationMs: 0,
      };
      cues.push(current);
    }
    const previous = current.items.at(-1);
    const baseStart =
      animation.trigger === 'after-previous'
        ? (previous?.endOffsetMs ?? 0)
        : animation.trigger === 'with-previous'
          ? (previous?.startOffsetMs ?? 0)
          : 0;
    const startOffsetMs = baseStart + animation.delayMs;
    const endOffsetMs = startOffsetMs + animation.durationMs;
    current.items.push({ animation, startOffsetMs, endOffsetMs });
    current.totalDurationMs = Math.max(current.totalDurationMs, endOffsetMs);
  }
  return cues;
}

export function initialWorkSlideAnimationCueIndex(
  cues: readonly WorkSlideAnimationCue[],
): number {
  return cues[0]?.automatic ? 0 : -1;
}

export function workSlideAnimationSequenceIssue(
  animations: readonly WorkSlideAnimation[] | undefined,
): WorkSlideAnimationSequenceIssue | undefined {
  const classTargets = new Set<string>();
  for (const animation of animations ?? []) {
    const key = `${animation.elementId}\u0000${workSlideAnimationClass(animation.effect)}`;
    if (classTargets.has(key)) return 'duplicate-class';
    classTargets.add(key);
  }

  for (const cue of workSlideAnimationCues(animations)) {
    const rangesByTarget = new Map<
      string,
      Array<{ endOffsetMs: number; startOffsetMs: number }>
    >();
    for (const item of cue.items) {
      const ranges = rangesByTarget.get(item.animation.elementId) ?? [];
      if (
        ranges.some(
          (range) =>
            item.startOffsetMs < range.endOffsetMs &&
            item.endOffsetMs > range.startOffsetMs,
        )
      ) {
        return 'overlapping-target';
      }
      ranges.push(item);
      rangesByTarget.set(item.animation.elementId, ranges);
    }
  }
  return undefined;
}

export function removeWorkSlideAnimationsForElements(
  slide: WorkSlide,
  elementIds: ReadonlySet<string>,
): WorkSlide {
  const animations = slide.animations?.filter(
    (animation) => !elementIds.has(animation.elementId),
  );
  if (animations?.length === slide.animations?.length) return slide;
  return {
    ...slide,
    animations: animations?.length ? animations : undefined,
  };
}

export function remapWorkSlideAnimations(
  animations: readonly WorkSlideAnimation[] | undefined,
  elementIds: ReadonlyMap<string, string>,
): WorkSlideAnimation[] | undefined {
  const remapped = (animations ?? []).flatMap((animation) => {
    const elementId = elementIds.get(animation.elementId);
    return elementId
      ? [
          {
            ...animation,
            id: createWorkId('slide-animation'),
            elementId,
          },
        ]
      : [];
  });
  return remapped.length ? remapped : undefined;
}

function normalizeWorkSlideAnimationDirection(
  direction: WorkSlideAnimationDirection | undefined,
): WorkSlideAnimationDirection {
  return direction === 'right' || direction === 'up' || direction === 'down'
    ? direction
    : 'left';
}

function isFlyAnimationEffect(effect: WorkSlideAnimationEffect): boolean {
  return effect === 'fly-in' || effect === 'fly-out';
}

function boundedInteger(value: number, minimum: number, maximum: number) {
  const finite = Number.isFinite(value) ? Math.round(value) : minimum;
  return Math.min(maximum, Math.max(minimum, finite));
}
