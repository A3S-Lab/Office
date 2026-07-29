import { expect, test } from '@rstest/core';
import { act, renderHook } from '@testing-library/react';
import { usePresentationSlideCommands } from '../src/internal/features/work/editors/use-presentation-slide-commands';
import type {
  WorkPresentationContent,
  WorkSlide,
  WorkSlideTransition,
} from '../src/internal/features/work/work-types';

const fadeTransition: WorkSlideTransition = {
  type: 'fade',
  speed: 'medium',
  advanceOnClick: true,
  advanceAfterMs: 5000,
};

test('applies an exact transition only when at least one slide would change', () => {
  const initial = presentation([
    slide('slide-1', fadeTransition),
    slide('slide-2'),
    slide('slide-3', { ...fadeTransition, speed: 'slow' }),
  ]);
  const changes: WorkPresentationContent[] = [];
  const { result, rerender } = renderHook(
    ({ content, selectedSlide }) =>
      usePresentationSlideCommands({
        content,
        selectedSlide,
        onChange: (next) => changes.push(next),
        onClearSelection: () => undefined,
        onSelectSlide: () => undefined,
      }),
    {
      initialProps: {
        content: initial,
        selectedSlide: initial.slides[0],
      },
    },
  );

  expect(result.current.canApplyTransitionToAll(fadeTransition)).toBe(true);
  act(() =>
    expect(result.current.applyTransitionToAll(fadeTransition)).toBe(true),
  );
  expect(changes).toHaveLength(1);
  expect(changes[0].slides.map((item) => item.transition)).toEqual([
    fadeTransition,
    fadeTransition,
    fadeTransition,
  ]);
  expect(changes[0].slides[0].transition).not.toBe(fadeTransition);
  expect(changes[0].slides[0].transition).not.toBe(
    changes[0].slides[1].transition,
  );

  const applied = changes[0];
  rerender({ content: applied, selectedSlide: applied.slides[0] });
  expect(result.current.canApplyTransitionToAll(fadeTransition)).toBe(false);
  act(() =>
    expect(result.current.applyTransitionToAll(fadeTransition)).toBe(false),
  );
  expect(changes).toHaveLength(1);
});

test('does not emit duplicate selected-slide transition changes', () => {
  const initial = presentation([
    slide('slide-1', fadeTransition),
    slide('slide-2'),
  ]);
  const changes: WorkPresentationContent[] = [];
  const { result } = renderHook(() =>
    usePresentationSlideCommands({
      content: initial,
      selectedSlide: initial.slides[0],
      onChange: (next) => changes.push(next),
      onClearSelection: () => undefined,
      onSelectSlide: () => undefined,
    }),
  );

  act(() => expect(result.current.setTransition(fadeTransition)).toBe(false));
  expect(changes).toEqual([]);

  const changed = { ...fadeTransition, advanceAfterMs: 7750 };
  act(() => expect(result.current.setTransition(changed)).toBe(true));
  expect(changes).toHaveLength(1);
  expect(changes[0].slides[0].transition).toEqual(changed);
  expect(changes[0].slides[1].transition).toBeUndefined();
});

test('can clear transitions from every slide without creating a no-op update', () => {
  const initial = presentation([
    slide('slide-1', fadeTransition),
    slide('slide-2', fadeTransition),
  ]);
  const changes: WorkPresentationContent[] = [];
  const { result, rerender } = renderHook(
    ({ content }) =>
      usePresentationSlideCommands({
        content,
        selectedSlide: content.slides[0],
        onChange: (next) => changes.push(next),
        onClearSelection: () => undefined,
        onSelectSlide: () => undefined,
      }),
    { initialProps: { content: initial } },
  );

  act(() => expect(result.current.applyTransitionToAll(undefined)).toBe(true));
  expect(changes[0].slides.every((item) => !item.transition)).toBe(true);

  rerender({ content: changes[0] });
  act(() => expect(result.current.applyTransitionToAll(undefined)).toBe(false));
  expect(changes).toHaveLength(1);
});

function presentation(slides: WorkSlide[]): WorkPresentationContent {
  return { type: 'presentation', slides };
}

function slide(id: string, transition?: WorkSlideTransition): WorkSlide {
  return {
    id,
    name: id,
    background: '#ffffff',
    elements: [],
    transition,
  };
}
