import { expect, test } from '@rstest/core';
import {
  clearPresentationClipboard,
  clonePresentationElementsForPaste,
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
  const copies = clonePresentationElementsForPaste(elements, 2);
  expect(copies).toMatchObject([
    { x: 12, y: 22, placeholder: undefined },
    { x: 42, y: 52, placeholder: undefined },
  ]);
  expect(copies[0].id).not.toBe(elements[0].id);
  expect(copies[1].id).not.toBe(elements[1].id);
  expect(copies[1].x - copies[0].x).toBe(30);
  expect(copies[1].y - copies[0].y).toBe(30);
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
