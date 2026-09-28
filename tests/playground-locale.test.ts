import { expect, test } from '@rstest/core';
import {
  resolvePlaygroundLocale,
  resolvePlaygroundVirtualGrid,
} from '../playground/src/playground-locale';

test('resolvePlaygroundLocale defaults to zh-CN and accepts en-US', () => {
  expect(resolvePlaygroundLocale('')).toBe('zh-CN');
  expect(resolvePlaygroundLocale('?locale=en-US')).toBe('en-US');
  expect(resolvePlaygroundLocale('?locale=zh-CN')).toBe('zh-CN');
  expect(resolvePlaygroundLocale('?locale=fr-FR')).toBe('zh-CN');
});

test('resolvePlaygroundVirtualGrid is opt-in via query', () => {
  expect(resolvePlaygroundVirtualGrid('')).toBe(false);
  expect(resolvePlaygroundVirtualGrid('?virtualGrid=1')).toBe(true);
  expect(resolvePlaygroundVirtualGrid('?virtualGrid=true')).toBe(true);
  expect(resolvePlaygroundVirtualGrid('?virtualGrid=0')).toBe(false);
});
