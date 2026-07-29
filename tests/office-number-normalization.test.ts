import { expect, test } from '@rstest/core';
import {
  normalizeOptionalOfficeNumber,
  normalizeRequiredOfficeNumber,
} from '../src/internal/features/work/editors/office-number-normalization';

test('normalizes required office measurements without accepting incomplete text', () => {
  expect(normalizeRequiredOfficeNumber('')).toBeNull();
  expect(normalizeRequiredOfficeNumber('-')).toBeNull();
  expect(
    normalizeRequiredOfficeNumber('12.56', {
      decimalPlaces: 1,
      minimum: 5,
      maximum: 60,
    }),
  ).toBe(12.6);
  expect(normalizeRequiredOfficeNumber('99', { minimum: 5, maximum: 60 })).toBe(
    60,
  );
  expect(normalizeRequiredOfficeNumber('3.7', { integer: true })).toBe(4);
});

test('distinguishes optional automatic values from invalid numbers', () => {
  expect(normalizeOptionalOfficeNumber('')).toBeUndefined();
  expect(normalizeOptionalOfficeNumber('12.5')).toBe(12.5);
  expect(normalizeOptionalOfficeNumber('not-a-number')).toBeNull();
  expect(
    normalizeOptionalOfficeNumber('0', {
      isValid: (value) => value > 0,
    }),
  ).toBeNull();
});
