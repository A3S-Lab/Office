import { expect, test } from '@rstest/core';
import {
  initialWorkSlideAnimationCueIndex,
  normalizeWorkSlideAnimation,
  remapWorkSlideAnimations,
  removeWorkSlideAnimationsForElements,
  workSlideAnimationCues,
} from '../src/internal/features/work/work-presentation-animation';
import { createWorkArtifact } from '../src/internal/features/work/work-templates';
import type {
  WorkSlide,
  WorkSlideAnimation,
} from '../src/internal/features/work/work-types';

test('builds explicit click cues with with-previous and after-previous timing', () => {
  const cues = workSlideAnimationCues([
    animation('a', 'one', 'on-click', 600, 100),
    animation('b', 'two', 'with-previous', 400, 50),
    animation('c', 'three', 'after-previous', 300, 75),
    animation('d', 'four', 'on-click', 200, 0),
  ]);

  expect(cues).toHaveLength(2);
  expect(cues[0]).toMatchObject({ automatic: false, totalDurationMs: 925 });
  expect(
    cues[0].items.map(({ startOffsetMs, endOffsetMs }) => [
      startOffsetMs,
      endOffsetMs,
    ]),
  ).toEqual([
    [100, 700],
    [150, 550],
    [625, 925],
  ]);
  expect(cues[1]).toMatchObject({ automatic: false, totalDurationMs: 200 });
  expect(initialWorkSlideAnimationCueIndex(cues)).toBe(-1);
});

test('starts a leading automatic cue without consuming a click', () => {
  const cues = workSlideAnimationCues([
    animation('a', 'one', 'after-previous', 500, 250),
    animation('b', 'two', 'with-previous', 300, 0),
    animation('c', 'three', 'on-click', 300, 0),
  ]);

  expect(cues.map(({ automatic }) => automatic)).toEqual([true, false]);
  expect(initialWorkSlideAnimationCueIndex(cues)).toBe(0);
});

test('normalizes bounded timing and effect-specific direction', () => {
  expect(
    normalizeWorkSlideAnimation({
      ...animation('a', 'one', 'on-click', -1, 100_000),
      effect: 'fade',
      direction: 'right',
    }),
  ).toMatchObject({ durationMs: 100, delayMs: 60_000, direction: undefined });
  expect(
    normalizeWorkSlideAnimation({
      ...animation('b', 'two', 'on-click', 100_000, -20),
      effect: 'fly-in',
      direction: undefined,
    }),
  ).toMatchObject({ durationMs: 60_000, delayMs: 0, direction: 'left' });
});

test('removes deleted targets and remaps slide-copy animation identities', () => {
  const source: WorkSlide = {
    id: 'slide',
    name: 'Animated',
    background: '#fff',
    elements: [],
    animations: [
      animation('a', 'one', 'on-click', 500, 0),
      animation('b', 'two', 'on-click', 500, 0),
    ],
  };

  expect(
    removeWorkSlideAnimationsForElements(source, new Set(['one'])).animations,
  ).toEqual([expect.objectContaining({ id: 'b', elementId: 'two' })]);
  const remapped = remapWorkSlideAnimations(
    source.animations,
    new Map([
      ['one', 'copy-one'],
      ['two', 'copy-two'],
    ]),
  );
  expect(remapped?.map(({ elementId }) => elementId)).toEqual([
    'copy-one',
    'copy-two',
  ]);
  expect(remapped?.map(({ id }) => id)).not.toEqual(['a', 'b']);
});

test('publishes an editable Playground template for every supported entrance effect and trigger', () => {
  const artifact = createWorkArtifact('animated-deck');
  if (artifact.content.type !== 'presentation') {
    throw new Error('Expected the animated Presentation template.');
  }
  const animations = artifact.content.slides[0]?.animations ?? [];

  expect(artifact.title).toBe('入场动画示例');
  expect(animations.map((animation) => animation.effect)).toEqual([
    'appear',
    'fade',
    'fly-in',
    'zoom',
  ]);
  expect(new Set(animations.map((animation) => animation.trigger))).toEqual(
    new Set(['on-click', 'with-previous', 'after-previous']),
  );
});

function animation(
  id: string,
  elementId: string,
  trigger: WorkSlideAnimation['trigger'],
  durationMs: number,
  delayMs: number,
): WorkSlideAnimation {
  return {
    id,
    elementId,
    effect: 'fade',
    trigger,
    durationMs,
    delayMs,
  };
}
