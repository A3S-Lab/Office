import { access, readFile } from 'node:fs/promises';
import path from 'node:path';
import { expect, test } from '@rstest/core';
import {
  DOCUMENTATION_DEFAULT_LANGUAGE,
  DOCUMENTATION_DEFAULT_VERSION,
  DOCUMENTATION_LOCALES,
  DOCUMENTATION_REQUIRED_ROUTES,
  DOCUMENTATION_VERSIONS,
} from '../website/documentation-site';

const documentationRoot = path.resolve(import.meta.dirname, '../docs');

test('uses Simplified Chinese and latest as stable documentation defaults', () => {
  expect(DOCUMENTATION_DEFAULT_LANGUAGE).toBe('zh');
  expect(DOCUMENTATION_LOCALES.map(({ lang }) => lang)).toEqual(['zh', 'en']);
  expect(DOCUMENTATION_DEFAULT_VERSION).toBe('latest');
  expect(DOCUMENTATION_VERSIONS).toEqual([
    'latest',
    '0.7.1',
    '0.7.0',
    '0.6.0',
    '0.5.0',
    '0.4.0',
    '0.3.0',
    '0.2.0',
    '0.1.0',
  ]);
});

test('keeps every public route available in every language and version', async () => {
  for (const version of DOCUMENTATION_VERSIONS) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      for (const route of DOCUMENTATION_REQUIRED_ROUTES) {
        await expect(
          access(path.join(documentationRoot, version, lang, route)),
        ).resolves.toBeUndefined();
      }
    }
  }
});

test('keeps published release homepages frozen and visibly versioned', async () => {
  for (const version of [
    '0.7.1',
    '0.7.0',
    '0.6.0',
    '0.5.0',
    '0.4.0',
    '0.3.0',
    '0.2.0',
    '0.1.0',
  ]) {
    const [chinese, english] = await Promise.all([
      readFile(path.join(documentationRoot, `${version}/zh/index.mdx`), 'utf8'),
      readFile(path.join(documentationRoot, `${version}/en/index.mdx`), 'utf8'),
    ]);

    expect(chinese).toContain(`# A3S Office ${version} 文档`);
    expect(chinese).toContain('冻结文档');
    expect(english).toContain(`# A3S Office ${version} documentation`);
    expect(english).toContain('frozen documentation');
  }
});

test('publishes real-time collaboration as a bilingual first-class capability', async () => {
  for (const version of ['latest', '0.7.1']) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [homepage, componentIndex, componentNavigation, collaboration] =
        await Promise.all([
          readFile(path.join(localeRoot, 'index.mdx'), 'utf8'),
          readFile(path.join(localeRoot, 'components/index.mdx'), 'utf8'),
          readFile(path.join(localeRoot, 'components/_meta.json'), 'utf8'),
          readFile(
            path.join(localeRoot, 'components/collaboration.mdx'),
            'utf8',
          ),
        ]);

      expect(homepage).toContain('./components/collaboration.mdx');
      expect(componentIndex).toContain('./collaboration.mdx');
      expect(componentNavigation).toContain('"collaboration"');
      expect(collaboration).toContain('Yjs');
      expect(collaboration).toContain('Awareness');
      expect(collaboration).toContain('Spreadsheet');
      expect(collaboration).toContain('Presentation');
      expect(collaboration).toContain('PDF');
    }
  }
});
