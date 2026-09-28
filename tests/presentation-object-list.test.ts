import { expect, test } from '@rstest/core';
import {
  presentationObjectLabel,
  presentationObjectListItems,
} from '../src/internal/features/work/editors/presentation-object-list';
import type { WorkSlideElement } from '../src/internal/features/work/work-types';

function element(
  patch: Partial<WorkSlideElement> & Pick<WorkSlideElement, 'id' | 'type'>,
): WorkSlideElement {
  return {
    x: 0,
    y: 0,
    width: 10,
    height: 10,
    text: '',
    fontSize: 18,
    color: '#000',
    fill: '#fff',
    bold: false,
    align: 'left',
    ...patch,
  };
}

test('prefers alt text, then body text, then typed fallback for object labels', () => {
  expect(
    presentationObjectLabel(
      element({ id: 'a', type: 'image', altText: 'Logo', text: 'ignored' }),
    ),
  ).toBe('Logo');
  expect(
    presentationObjectLabel(
      element({ id: 'b', type: 'text', text: 'Quarterly plan' }),
    ),
  ).toBe('Quarterly plan');
  expect(presentationObjectLabel(element({ id: 'c', type: 'chart' }))).toBe(
    'Chart',
  );
});

test('builds a stable object list for the current slide', () => {
  const items = presentationObjectListItems([
    element({ id: 't1', type: 'text', text: 'Title' }),
    element({ id: 's1', type: 'shape' }),
  ]);
  expect(items).toEqual([
    { id: 't1', type: 'text', label: 'Title' },
    { id: 's1', type: 'shape', label: 'Shape' },
  ]);
});
