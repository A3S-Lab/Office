import { expect, test } from '@rstest/core';
import { act, renderHook } from '@testing-library/react';
import { usePresentationAnimationCommands } from '../src/internal/features/work/editors/use-presentation-animation-commands';
import type {
  WorkPresentationContent,
  WorkSlide,
  WorkSlideElement,
} from '../src/internal/features/work/work-types';

test('creates, updates, reorders, and removes one entrance animation per object', () => {
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

  act(() => expect(result.current.setEntranceAnimation('fade')).toBe(true));
  expect(changes).toHaveLength(1);
  expect(changes[0].slides[0].animations).toEqual([
    expect.objectContaining({ elementId: 'one', effect: 'fade' }),
  ]);

  rerender({ content: changes[0], elementId: 'two' });
  act(() => expect(result.current.setEntranceAnimation('fly-in')).toBe(true));
  const withTwo = changes[1];
  expect(
    withTwo.slides[0].animations?.map(({ elementId }) => elementId),
  ).toEqual(['one', 'two']);

  rerender({ content: withTwo, elementId: 'two' });
  act(() =>
    expect(
      result.current.updateEntranceAnimation({
        trigger: 'with-previous',
        durationMs: 775,
        delayMs: 250,
        direction: 'right',
      }),
    ).toBe(true),
  );
  expect(changes[2].slides[0].animations?.[1]).toMatchObject({
    trigger: 'with-previous',
    durationMs: 775,
    delayMs: 250,
    direction: 'right',
  });

  rerender({ content: changes[2], elementId: 'two' });
  act(() => expect(result.current.moveEntranceAnimation(-1)).toBe(true));
  expect(
    changes[3].slides[0].animations?.map(({ elementId }) => elementId),
  ).toEqual(['two', 'one']);

  rerender({ content: changes[3], elementId: 'two' });
  act(() => expect(result.current.setEntranceAnimation(undefined)).toBe(true));
  expect(changes[4].slides[0].animations).toEqual([
    expect.objectContaining({ elementId: 'one' }),
  ]);
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

  act(() => expect(result.current.setEntranceAnimation('appear')).toBe(false));
  rerender({ elementId: 'missing' });
  act(() => expect(result.current.setEntranceAnimation('appear')).toBe(false));
  act(() =>
    expect(result.current.updateEntranceAnimation({ delayMs: 10 })).toBe(false),
  );
  act(() => expect(result.current.moveEntranceAnimation(1)).toBe(false));
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
