import { expect, test } from '@rstest/core';
import { newPresentationTableElement } from '../src/internal/features/work/editors/presentation-editor-operations';

test('creates presentation tables with the requested dimensions', () => {
  const element = newPresentationTableElement(4, 5);

  expect(element.type).toBe('table');
  expect(element.table?.headerRows).toBe(1);
  expect(element.table?.rows).toHaveLength(4);
  expect(element.table?.rows.every((row) => row.length === 5)).toBe(true);
  expect(element.table?.rows[0]).toEqual([
    '标题 1',
    '标题 2',
    '标题 3',
    '标题 4',
    '标题 5',
  ]);
});
