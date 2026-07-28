import { expect, test } from '@rstest/core';
import {
  calculateRibbonOverflow,
  calculateRibbonScrollTarget,
} from '../src/internal/features/work/editors/work-office-ribbon-overflow';

const compactSpreadsheetGroups = [
  { left: 5, right: 124 },
  { left: 124, right: 464 },
  { left: 464, right: 688 },
  { left: 688, right: 886 },
  { left: 886, right: 958 },
];

test('pages compact ribbons to complete groups with room for navigation', () => {
  expect(
    calculateRibbonOverflow({
      clientWidth: 768,
      items: compactSpreadsheetGroups,
      scrollLeft: 0,
    }),
  ).toEqual({ backward: false, forward: true });

  const next = calculateRibbonScrollTarget({
    clientWidth: 768,
    direction: 1,
    items: compactSpreadsheetGroups,
    scrollLeft: 0,
  });
  expect(next).toBe(654);
  expect(
    calculateRibbonOverflow({
      clientWidth: 768,
      items: compactSpreadsheetGroups,
      scrollLeft: next,
    }),
  ).toEqual({ backward: true, forward: false });
  expect(
    calculateRibbonScrollTarget({
      clientWidth: 768,
      direction: -1,
      items: compactSpreadsheetGroups,
      scrollLeft: next,
    }),
  ).toBe(430);
});

test('does not create overflow navigation when every group fits', () => {
  const items = [
    { left: 10, right: 120 },
    { left: 120, right: 280 },
  ];
  expect(
    calculateRibbonOverflow({
      clientWidth: 768,
      items,
      scrollLeft: 0,
    }),
  ).toEqual({ backward: false, forward: false });
  expect(
    calculateRibbonScrollTarget({
      clientWidth: 768,
      direction: 1,
      items,
      scrollLeft: 0,
    }),
  ).toBe(0);
});

test('resets stale compact navigation after the ribbon grows', () => {
  expect(
    calculateRibbonOverflow({
      clientWidth: 1280,
      items: compactSpreadsheetGroups,
      scrollLeft: 654,
    }),
  ).toEqual({ backward: false, forward: false });
  expect(
    calculateRibbonScrollTarget({
      clientWidth: 1280,
      direction: -1,
      items: compactSpreadsheetGroups,
      scrollLeft: 654,
    }),
  ).toBe(0);
});
