import { expect, test } from '@rstest/core';
import {
  canGroupPresentationElements,
  canUngroupPresentationElements,
  expandPresentationGroupSelection,
  groupPresentationElements,
  presentationSelectionUnitCount,
  ungroupPresentationElements,
} from '../src/internal/features/work/work-presentation-groups';
import type { WorkSlideElement } from '../src/internal/features/work/work-types';
import {
  alignPresentationSelection,
  distributePresentationSelection,
  selectPresentationElementUnit,
} from '../src/internal/features/work/editors/presentation-selection';

test('selects a persistent presentation group as one ordered unit', () => {
  const elements = [
    presentationElement('first', 0, ['group']),
    presentationElement('second', 10, ['group']),
    presentationElement('outside', 40),
  ];

  const selected = selectPresentationElementUnit(elements, [], 'first', false);
  expect(new Set(selected)).toEqual(new Set(['first', 'second']));
  expect(selected.at(-1)).toBe('first');

  const additive = selectPresentationElementUnit(
    elements,
    selected,
    'outside',
    true,
  );
  expect(additive).toEqual(['second', 'first', 'outside']);
  expect(
    selectPresentationElementUnit(elements, additive, 'second', true),
  ).toEqual(['outside']);
});

test('creates nested group paths and removes one group level at a time', () => {
  const elements = [
    presentationElement('first', 0, ['inner']),
    presentationElement('second', 10, ['inner']),
    presentationElement('outside', 40),
  ];
  const selectedIds = ['first', 'second', 'outside'];

  expect(canGroupPresentationElements(elements, selectedIds)).toBe(true);
  const grouped = groupPresentationElements(elements, selectedIds, 'outer');
  expect(grouped.map((element) => element.groupIds)).toEqual([
    ['outer', 'inner'],
    ['outer', 'inner'],
    ['outer'],
  ]);
  expect(presentationSelectionUnitCount(grouped, selectedIds)).toBe(1);
  expect(canUngroupPresentationElements(grouped, ['first'])).toBe(true);
  expect(expandPresentationGroupSelection(grouped, ['first'])).toEqual([
    'second',
    'outside',
    'first',
  ]);

  const ungrouped = ungroupPresentationElements(grouped, ['first']);
  expect(ungrouped.map((element) => element.groupIds)).toEqual([
    ['inner'],
    ['inner'],
    undefined,
  ]);
});

test('aligns presentation groups without changing member geometry', () => {
  const elements = [
    presentationElement('group-left', 0, ['group']),
    presentationElement('group-right', 10, ['group']),
    presentationElement('outside', 50),
  ];

  const aligned = alignPresentationSelection(
    elements,
    ['group-left', 'group-right', 'outside'],
    'right',
  );
  expect(aligned).toMatchObject([
    { id: 'group-left', x: 40 },
    { id: 'group-right', x: 50 },
    { id: 'outside', x: 50 },
  ]);
  expect(aligned[1].x - aligned[0].x).toBe(10);
});

test('distributes presentation groups as logical units', () => {
  const elements = [
    presentationElement('left', 0),
    presentationElement('group-left', 20, ['group']),
    presentationElement('group-right', 30, ['group']),
    presentationElement('right', 90),
  ];

  const distributed = distributePresentationSelection(
    elements,
    elements.map((element) => element.id),
    'horizontal',
  );

  expect(distributed).toMatchObject([
    { id: 'left', x: 0 },
    { id: 'group-left', x: 40 },
    { id: 'group-right', x: 50 },
    { id: 'right', x: 90 },
  ]);
  expect(distributed[2].x - distributed[1].x).toBe(10);
});

function presentationElement(
  id: string,
  x: number,
  groupIds?: string[],
): WorkSlideElement {
  return {
    id,
    type: 'shape',
    x,
    y: 10,
    width: 10,
    height: 10,
    text: id,
    fontSize: 14,
    color: '#172033',
    fill: '#dce6fb',
    bold: false,
    align: 'center',
    groupIds,
  };
}
