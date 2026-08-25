import { expect, test } from '@rstest/core';
import { act, renderHook } from '@testing-library/react';
import { usePresentationClipboard } from '../src/internal/features/work/editors/use-presentation-clipboard';
import type {
  WorkPresentationContent,
  WorkSlideAnimation,
  WorkSlideElement,
} from '../src/internal/features/work/work-types';

test('duplicates an animated object with independent element and animation identities', () => {
  const content = animatedPresentation();
  const changes: WorkPresentationContent[] = [];
  const { result } = renderHook(() =>
    usePresentationClipboard({
      content,
      mode: 'slide',
      targetId: 'slide',
      selectedSlide: content.slides[0],
      selectedElements: [content.slides[0].elements[0]],
      onChange: (next) => changes.push(next),
      onSelectSlide: () => undefined,
      onSelectElements: () => undefined,
    }),
  );

  act(() => expect(result.current.duplicateSelection()).toBe(true));
  const next = changes[0].slides[0];
  expect(next.elements).toHaveLength(3);
  expect(next.animations).toHaveLength(2);
  expect(next.animations?.[1]).toMatchObject({
    elementId: next.elements[2].id,
    effect: 'fade',
  });
  expect(next.animations?.[1].id).not.toBe(next.animations?.[0].id);
});

test('deletes every animation owned by the deleted object', () => {
  const content = animatedPresentation();
  const changes: WorkPresentationContent[] = [];
  const { result } = renderHook(() =>
    usePresentationClipboard({
      content,
      mode: 'slide',
      targetId: 'slide',
      selectedSlide: content.slides[0],
      selectedElements: [content.slides[0].elements[0]],
      onChange: (next) => changes.push(next),
      onSelectSlide: () => undefined,
      onSelectElements: () => undefined,
    }),
  );

  act(() => expect(result.current.deleteSelection()).toBe(true));
  expect(changes[0].slides[0].elements.map(({ id }) => id)).toEqual(['two']);
  expect(changes[0].slides[0].animations).toBeUndefined();
});

test('does not publish a duplicate that would exceed the slide animation limit', () => {
  const elements = Array.from({ length: 256 }, (_, index) =>
    element(`element-${index}`),
  );
  const animations: WorkSlideAnimation[] = elements.map((item, index) => ({
    id: `animation-${index}`,
    elementId: item.id,
    effect: 'fade',
    trigger: 'on-click',
    durationMs: 500,
    delayMs: 0,
  }));
  const content: WorkPresentationContent = {
    type: 'presentation',
    slides: [{ id: 'slide', name: 'At limit', elements, animations }],
  };
  const changes: WorkPresentationContent[] = [];
  const { result } = renderHook(() =>
    usePresentationClipboard({
      content,
      mode: 'slide',
      targetId: 'slide',
      selectedSlide: content.slides[0],
      selectedElements: [elements[0]],
      onChange: (next) => changes.push(next),
      onSelectSlide: () => undefined,
      onSelectElements: () => undefined,
    }),
  );

  act(() => expect(result.current.duplicateSelection()).toBe(true));
  expect(changes).toEqual([]);
});

function animatedPresentation(): WorkPresentationContent {
  return {
    type: 'presentation',
    slides: [
      {
        id: 'slide',
        name: 'Animated',
        elements: [element('one'), element('two')],
        animations: [
          {
            id: 'animation-one',
            elementId: 'one',
            effect: 'fade',
            trigger: 'on-click',
            durationMs: 500,
            delayMs: 0,
          },
        ],
      },
    ],
  };
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
