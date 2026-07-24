import { expect, test } from '@rstest/core';
import type { WorkSlideElement } from '../src/internal/features/work/work-types';
import {
  alignPresentationSelection,
  distributePresentationSelection,
  reorderPresentationSelection,
  selectPresentationElement,
  translatePresentationSelection,
} from '../src/internal/features/work/editors/presentation-selection';

test('keeps an ordered additive presentation selection', () => {
  expect(selectPresentationElement([], 'title', false)).toEqual(['title']);
  expect(selectPresentationElement(['title'], 'accent', true)).toEqual([
    'title',
    'accent',
  ]);
  expect(selectPresentationElement(['title', 'accent'], 'title', true)).toEqual(
    ['accent'],
  );
  expect(selectPresentationElement(['title', 'accent'], 'body', false)).toEqual(
    ['body'],
  );
});

test('translates a selection as one bounded object set', () => {
  const elements = [
    presentationElement('left', 70, 5, 10, 10),
    presentationElement('right', 85, 25, 10, 10),
    presentationElement('outside', 20, 20, 5, 5),
  ];

  expect(
    translatePresentationSelection(elements, ['left', 'right'], 20, -20),
  ).toMatchObject([
    { id: 'left', x: 75, y: 0 },
    { id: 'right', x: 90, y: 20 },
    { id: 'outside', x: 20, y: 20 },
  ]);
});

test('aligns multiple objects to their shared selection bounds', () => {
  const elements = [
    presentationElement('first', 10, 20, 10, 10),
    presentationElement('second', 40, 40, 20, 20),
  ];

  expect(
    alignPresentationSelection(elements, ['first', 'second'], 'center'),
  ).toMatchObject([
    { id: 'first', x: 30, y: 20 },
    { id: 'second', x: 25, y: 40 },
  ]);
  expect(
    alignPresentationSelection(elements, ['first', 'second'], 'top'),
  ).toMatchObject([
    { id: 'first', x: 10, y: 20 },
    { id: 'second', x: 40, y: 20 },
  ]);
});

test('distributes three or more selected objects with equal gaps', () => {
  const elements = [
    presentationElement('first', 0, 0, 10, 10),
    presentationElement('second', 30, 25, 10, 10),
    presentationElement('third', 90, 50, 10, 10),
  ];

  expect(
    distributePresentationSelection(
      elements,
      ['first', 'second', 'third'],
      'horizontal',
    ),
  ).toMatchObject([
    { id: 'first', x: 0 },
    { id: 'second', x: 45 },
    { id: 'third', x: 90 },
  ]);
  expect(
    distributePresentationSelection(
      elements,
      ['first', 'second', 'third'],
      'vertical',
    ),
  ).toMatchObject([
    { id: 'first', y: 0 },
    { id: 'second', y: 25 },
    { id: 'third', y: 50 },
  ]);
});

test('moves a selected layer set without changing its internal order', () => {
  const elements = [
    presentationElement('first', 0, 0, 10, 10),
    presentationElement('second', 10, 0, 10, 10),
    presentationElement('third', 20, 0, 10, 10),
    presentationElement('fourth', 30, 0, 10, 10),
  ];

  expect(
    reorderPresentationSelection(elements, ['second', 'third'], 1).map(
      (element) => element.id,
    ),
  ).toEqual(['first', 'fourth', 'second', 'third']);
  expect(
    reorderPresentationSelection(elements, ['second', 'third'], -1).map(
      (element) => element.id,
    ),
  ).toEqual(['second', 'third', 'first', 'fourth']);
});

function presentationElement(
  id: string,
  x: number,
  y: number,
  width: number,
  height: number,
): WorkSlideElement {
  return {
    id,
    type: 'shape',
    x,
    y,
    width,
    height,
    text: id,
    fontSize: 14,
    color: '#172033',
    fill: '#dce6fb',
    bold: false,
    align: 'center',
  };
}
