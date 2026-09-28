import { expect, test } from '@rstest/core';
import {
  officeMessage,
  resolveOfficeLocale,
  resolveOfficeMessages,
} from '../src/internal/i18n/office-locale';

test('resolveOfficeLocale defaults to zh-CN and accepts en-US', () => {
  expect(resolveOfficeLocale(undefined)).toBe('zh-CN');
  expect(resolveOfficeLocale('en-US')).toBe('en-US');
  expect(resolveOfficeLocale('fr-FR')).toBe('zh-CN');
});

test('resolveOfficeMessages merges overrides on the selected locale', () => {
  const catalog = resolveOfficeMessages({
    locale: 'en-US',
    messages: {
      'editor.loading.document': 'Custom opener',
    },
  });
  expect(catalog['editor.loading.document']).toBe('Custom opener');
  expect(catalog['editor.loading.pdf']).toBe('Opening PDF');
  expect(
    officeMessage(catalog, 'editor.error.failed', { title: 'Writer' }),
  ).toBe('Writer failed to load.');
});

test('zh-CN catalog keeps Chinese loading titles as the default product voice', () => {
  const catalog = resolveOfficeMessages({ locale: 'zh-CN' });
  expect(catalog['editor.loading.document']).toBe('正在打开文字编辑器');
});
