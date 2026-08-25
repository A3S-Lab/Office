import { expect, test } from '@rstest/core';
import {
  clearPresentationClipboard,
  clonePresentationElementsForPaste,
  clonePresentationElementsAndAnimationsForPaste,
  clonePresentationSlideForPaste,
  copyPresentationElements,
  takePresentationClipboard,
} from '../src/internal/features/work/work-presentation-clipboard';
import type { WorkSlideElement } from '../src/internal/features/work/work-types';

test('copies and offsets a presentation object set without changing its layout', () => {
  const elements = [
    presentationElement('first', 10, 20),
    presentationElement('second', 40, 50),
  ];
  clearPresentationClipboard();
  copyPresentationElements(elements);

  const clipboard = takePresentationClipboard();
  expect(clipboard?.payload).toMatchObject({
    kind: 'elements',
    elements: [{ id: 'first' }, { id: 'second' }],
  });
  expect(clipboard?.offset).toBe(5);
  const copies = clonePresentationElementsForPaste(
    elements,
    clipboard?.offset ?? 0,
  );
  expect(copies).toMatchObject([
    { x: 15, y: 25, placeholder: undefined },
    { x: 45, y: 55, placeholder: undefined },
  ]);
  expect(copies[0].id).not.toBe(elements[0].id);
  expect(copies[1].id).not.toBe(elements[1].id);
  expect(copies[1].x - copies[0].x).toBe(30);
  expect(copies[1].y - copies[0].y).toBe(30);
});

test('keeps repeated presentation pastes visibly offset and bounded', () => {
  clearPresentationClipboard();
  copyPresentationElements([presentationElement('source', 10, 20)]);

  expect(
    Array.from({ length: 6 }, () => takePresentationClipboard()?.offset),
  ).toEqual([5, 10, 15, 20, 20, 20]);
});

test('clones presentation group paths without linking the copy to its source', () => {
  const elements = [
    {
      ...presentationElement('first', 10, 20),
      groupIds: ['outer', 'inner'],
    },
    {
      ...presentationElement('second', 40, 50),
      groupIds: ['outer', 'inner'],
    },
    {
      ...presentationElement('third', 60, 20),
      groupIds: ['outer'],
    },
  ];

  const copies = clonePresentationElementsForPaste(elements, 2);
  expect(copies[0].groupIds?.[0]).toBe(copies[1].groupIds?.[0]);
  expect(copies[0].groupIds?.[0]).toBe(copies[2].groupIds?.[0]);
  expect(copies[0].groupIds?.[1]).toBe(copies[1].groupIds?.[1]);
  expect(copies[0].groupIds).not.toEqual(elements[0].groupIds);

  const slideCopy = clonePresentationSlideForPaste({
    id: 'slide',
    name: 'Grouped slide',
    background: '#ffffff',
    elements,
  });
  expect(slideCopy.elements[0].groupIds?.[0]).toBe(
    slideCopy.elements[2].groupIds?.[0],
  );
  expect(slideCopy.elements[0].groupIds).not.toEqual(elements[0].groupIds);
});

test('numbers repeated slide copies without growing the source label', () => {
  const source = {
    id: 'slide',
    name: '封面 副本',
    background: '#ffffff',
    elements: [],
  };
  const copy = clonePresentationSlideForPaste(source, [
    '封面',
    '封面 副本',
    '封面 副本 2',
  ]);

  expect(copy.name).toBe('封面 副本 3');
});

test('remaps slide animation identities and targets with copied objects', () => {
  const elements = [
    presentationElement('first', 10, 20),
    presentationElement('second', 40, 50),
  ];
  const copy = clonePresentationSlideForPaste({
    id: 'animated-slide',
    name: 'Animated',
    background: '#ffffff',
    elements,
    animations: [
      {
        id: 'animation-first',
        elementId: 'first',
        effect: 'fade',
        trigger: 'on-click',
        durationMs: 500,
        delayMs: 0,
      },
      {
        id: 'animation-second',
        elementId: 'second',
        effect: 'zoom',
        trigger: 'after-previous',
        durationMs: 700,
        delayMs: 100,
      },
    ],
  });

  expect(copy.animations?.map(({ elementId }) => elementId)).toEqual(
    copy.elements.map(({ id }) => id),
  );
  expect(copy.animations?.map(({ id }) => id)).not.toEqual([
    'animation-first',
    'animation-second',
  ]);
});

test('remaps selected object animations without copying unrelated cues', () => {
  const elements = [
    presentationElement('first', 10, 20),
    presentationElement('second', 40, 50),
  ];
  const paste = clonePresentationElementsAndAnimationsForPaste(
    [elements[0]],
    [
      {
        id: 'animation-first',
        elementId: 'first',
        effect: 'fade',
        trigger: 'on-click',
        durationMs: 500,
        delayMs: 0,
      },
      {
        id: 'animation-second',
        elementId: 'second',
        effect: 'zoom',
        trigger: 'on-click',
        durationMs: 500,
        delayMs: 0,
      },
    ],
    5,
  );

  expect(paste.elements).toHaveLength(1);
  expect(paste.animations).toEqual([
    expect.objectContaining({
      elementId: paste.elements[0].id,
      effect: 'fade',
    }),
  ]);
  expect(paste.animations?.[0].id).not.toBe('animation-first');
});

function presentationElement(
  id: string,
  x: number,
  y: number,
): WorkSlideElement {
  return {
    id,
    type: 'shape',
    x,
    y,
    width: 20,
    height: 20,
    text: id,
    fontSize: 14,
    color: '#172033',
    fill: '#dce6fb',
    bold: false,
    align: 'center',
    placeholder: {
      key: id,
      type: 'body',
    },
  };
}
