import { expect, test } from '@rstest/core';
import { act, renderHook } from '@testing-library/react';
import { usePresentationAnimationCommands } from '../src/internal/features/work/editors/use-presentation-animation-commands';
import type {
  WorkPresentationContent,
  WorkSlide,
  WorkSlideElement,
} from '../src/internal/features/work/work-types';

test('creates, updates, reorders, and removes one entrance and exit animation per object', () => {
  const initial = presentation();
  const changes: WorkPresentationContent[] = [];
  const { result, rerender } = renderHook(
    ({ content, elementId }) =>
      usePresentationAnimationCommands({
        content,
        selectedSlide: content.slides[0],
        selectedElementId: elementId,
        onChange: (next) => changes.push(next),
      }),
    { initialProps: { content: initial, elementId: 'one' } },
  );

  act(() => expect(result.current.setAnimation('entrance', 'fade')).toBe(true));
  expect(changes).toHaveLength(1);
  expect(changes[0].slides[0].animations).toEqual([
    expect.objectContaining({ elementId: 'one', effect: 'fade' }),
  ]);

  rerender({ content: changes[0], elementId: 'one' });
  act(() => expect(result.current.setAnimation('exit', 'fade-out')).toBe(true));
  expect(changes[1].slides[0].animations).toEqual([
    expect.objectContaining({ elementId: 'one', effect: 'fade' }),
    expect.objectContaining({ elementId: 'one', effect: 'fade-out' }),
  ]);

  rerender({ content: changes[1], elementId: 'one' });
  act(() =>
    expect(
      result.current.updateAnimation('exit', {
        trigger: 'after-previous',
        durationMs: 775,
        delayMs: 250,
      }),
    ).toBe(true),
  );
  expect(changes[2].slides[0].animations?.[1]).toMatchObject({
    effect: 'fade-out',
    trigger: 'after-previous',
    durationMs: 775,
    delayMs: 250,
  });

  rerender({ content: changes[2], elementId: 'two' });
  act(() =>
    expect(result.current.setAnimation('entrance', 'fly-in')).toBe(true),
  );
  const withTwo = changes[3];
  expect(
    withTwo.slides[0].animations?.map(({ elementId }) => elementId),
  ).toEqual(['one', 'one', 'two']);

  rerender({ content: withTwo, elementId: 'two' });
  act(() =>
    expect(
      result.current.updateAnimation('entrance', {
        trigger: 'with-previous',
        durationMs: 775,
        delayMs: 250,
        direction: 'right',
      }),
    ).toBe(true),
  );
  expect(changes[4].slides[0].animations?.[2]).toMatchObject({
    trigger: 'with-previous',
    durationMs: 775,
    delayMs: 250,
    direction: 'right',
  });

  rerender({ content: changes[4], elementId: 'two' });
  act(() => expect(result.current.moveAnimation('entrance', -1)).toBe(true));
  expect(
    changes[5].slides[0].animations?.map(({ elementId }) => elementId),
  ).toEqual(['one', 'two', 'one']);

  rerender({ content: changes[5], elementId: 'one' });
  act(() => expect(result.current.setAnimation('exit', undefined)).toBe(true));
  expect(changes[6].slides[0].animations).toEqual([
    expect.objectContaining({ elementId: 'one', effect: 'fade' }),
    expect.objectContaining({ elementId: 'two', effect: 'fly-in' }),
  ]);
});

test('rejects an overlapping animation on the same object and cue', () => {
  const initial = presentation();
  initial.slides[0].animations = [
    {
      id: 'entrance',
      elementId: 'one',
      effect: 'fade',
      trigger: 'on-click',
      durationMs: 500,
      delayMs: 0,
    },
    {
      id: 'exit',
      elementId: 'one',
      effect: 'fade-out',
      trigger: 'after-previous',
      durationMs: 300,
      delayMs: 0,
    },
  ];
  const changes: WorkPresentationContent[] = [];
  const { result } = renderHook(() =>
    usePresentationAnimationCommands({
      content: initial,
      selectedSlide: initial.slides[0],
      selectedElementId: 'one',
      onChange: (next) => changes.push(next),
    }),
  );

  expect(
    result.current.canUpdateAnimation('exit', { trigger: 'with-previous' }),
  ).toBe(false);
  act(() =>
    expect(
      result.current.updateAnimation('exit', { trigger: 'with-previous' }),
    ).toBe(false),
  );
  expect(changes).toEqual([]);
});

test('does not publish no-op or unavailable animation intents', () => {
  const initial = presentation();
  const changes: WorkPresentationContent[] = [];
  const { result, rerender } = renderHook(
    ({ elementId }) =>
      usePresentationAnimationCommands({
        content: initial,
        selectedSlide: initial.slides[0],
        selectedElementId: elementId,
        onChange: (next) => changes.push(next),
      }),
    { initialProps: { elementId: undefined as string | undefined } },
  );

  act(() =>
    expect(result.current.setAnimation('entrance', 'appear')).toBe(false),
  );
  rerender({ elementId: 'missing' });
  act(() =>
    expect(result.current.setAnimation('entrance', 'appear')).toBe(false),
  );
  act(() =>
    expect(result.current.updateAnimation('entrance', { delayMs: 10 })).toBe(
      false,
    ),
  );
  act(() => expect(result.current.moveAnimation('entrance', 1)).toBe(false));
  expect(changes).toEqual([]);
});

function presentation(): WorkPresentationContent {
  const elements: WorkSlideElement[] = [element('one'), element('two')];
  const slide: WorkSlide = {
    id: 'slide',
    name: 'Animated',
    background: '#ffffff',
    elements,
  };
  return { type: 'presentation', slides: [slide] };
}

function element(id: string): WorkSlideElement {
  return {
    id,
    type: 'text',
    x: 10,
    y: 10,
    width: 20,
    height: 10,
    text: id,
    fontSize: 20,
    color: '#111111',
    fill: 'transparent',
    bold: false,
    align: 'left',
  };
}
