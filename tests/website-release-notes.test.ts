import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { expect, test } from '@rstest/core';
import {
  DOCUMENTATION_LOCALES,
  DOCUMENTATION_VERSIONS,
} from '../website/documentation-site';
import {
  OFFICE_RELEASE_NOTES,
  officeReleaseNotesThroughVersion,
} from '../website/theme/release-notes-data';

const documentationRoot = path.resolve(import.meta.dirname, '../docs');
const repositoryRoot = path.resolve(import.meta.dirname, '..');

test('keeps curated release notes unique, localized, and newest first', () => {
  expect(OFFICE_RELEASE_NOTES.map(({ version }) => version)).toEqual([
    '0.53.0',
    '0.52.0',
    '0.51.0',
    '0.50.0',
    '0.49.0',
    '0.48.1',
    '0.48.0',
    '0.47.0',
    '0.46.0',
    '0.45.0',
    '0.44.0',
    '0.43.0',
    '0.42.0',
    '0.41.0',
    '0.40.0',
    '0.39.0',
    '0.38.1',
    '0.38.0',
    '0.37.5',
    '0.37.4',
    '0.37.3',
    '0.37.2',
    '0.37.1',
    '0.37.0',
    '0.36.0',
    '0.35.0',
    '0.34.0',
    '0.33.0',
    '0.32.0',
    '0.31.0',
    '0.30.0',
  ]);

  expect(new Set(OFFICE_RELEASE_NOTES.map(({ version }) => version)).size).toBe(
    OFFICE_RELEASE_NOTES.length,
  );

  for (const release of OFFICE_RELEASE_NOTES) {
    expect(release.date).toMatch(/^\d{4}-\d{2}-\d{2}$/);
    expect(release.surfaces.length).toBeGreaterThan(0);
    expect(release.highlights).toHaveLength(3);
    expect(release.links.length).toBeGreaterThan(0);

    for (const localized of [
      release.title,
      release.summary,
      ...release.highlights.flatMap(({ title, detail }) => [title, detail]),
      ...release.links.flatMap(({ href, label }) => [href, label]),
    ]) {
      expect(localized.en.trim()).not.toBe('');
      expect(localized.zh.trim()).not.toBe('');
    }
  }
});

test('cuts the visual changelog at the active frozen documentation version', () => {
  expect(
    officeReleaseNotesThroughVersion('latest').map(({ version }) => version),
  ).toEqual(OFFICE_RELEASE_NOTES.map(({ version }) => version));

  expect(
    officeReleaseNotesThroughVersion('0.38.0').map(({ version }) => version),
  ).toEqual([
    '0.38.0',
    '0.37.5',
    '0.37.4',
    '0.37.3',
    '0.37.2',
    '0.37.1',
    '0.37.0',
    '0.36.0',
    '0.35.0',
    '0.34.0',
    '0.33.0',
    '0.32.0',
    '0.31.0',
    '0.30.0',
  ]);
  expect(
    officeReleaseNotesThroughVersion('0.37.2').map(({ version }) => version),
  ).toEqual([
    '0.37.2',
    '0.37.1',
    '0.37.0',
    '0.36.0',
    '0.35.0',
    '0.34.0',
    '0.33.0',
    '0.32.0',
    '0.31.0',
    '0.30.0',
  ]);
  expect(
    officeReleaseNotesThroughVersion('0.36.0').map(({ version }) => version),
  ).toEqual([
    '0.36.0',
    '0.35.0',
    '0.34.0',
    '0.33.0',
    '0.32.0',
    '0.31.0',
    '0.30.0',
  ]);
  expect(officeReleaseNotesThroughVersion('0.29.0')).toEqual([]);
});

test('publishes a bilingual changelog route throughout version navigation', async () => {
  for (const version of DOCUMENTATION_VERSIONS) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [page, metadataSource, navigationSource] = await Promise.all([
        readFile(path.join(localeRoot, 'changelog.mdx'), 'utf8'),
        readFile(path.join(localeRoot, '_meta.json'), 'utf8'),
        readFile(path.join(localeRoot, '_nav.json'), 'utf8'),
      ]);
      const metadata = JSON.parse(metadataSource) as unknown[];
      const navigation = JSON.parse(navigationSource) as Array<{
        activeMatch?: string;
        link?: string;
        text?: string;
      }>;
      const changelogNavigation = navigation.find(
        ({ link }) => link === '/changelog.html',
      );

      expect(page).toContain('<ReleaseNotes />');
      expect(metadata).toContain('changelog');
      expect(changelogNavigation?.text).toBe(
        lang === 'en' ? "What's new" : '更新日志',
      );
      expect(changelogNavigation?.activeMatch).toContain('changelog');
    }
  }
});

test('keeps the changelog semantic, responsive, static, and discoverable', async () => {
  const [
    component,
    styles,
    englishPage,
    chinesePage,
    englishHome,
    chineseHome,
    visualContract,
    a3sTestContract,
    packageManifest,
  ] = await Promise.all([
    readFile(
      path.join(repositoryRoot, 'website/theme/release-notes.tsx'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'website/theme/release-notes.css'),
      'utf8',
    ),
    readFile(path.join(documentationRoot, 'latest/en/changelog.mdx'), 'utf8'),
    readFile(path.join(documentationRoot, 'latest/zh/changelog.mdx'), 'utf8'),
    readFile(path.join(documentationRoot, 'latest/en/index.mdx'), 'utf8'),
    readFile(path.join(documentationRoot, 'latest/zh/index.mdx'), 'utf8'),
    readFile(
      path.join(
        repositoryRoot,
        'visual-tests/docs-changelog.functional.spec.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/office-docs-navigation.acl'),
      'utf8',
    ),
    readFile(path.join(repositoryRoot, 'package.json'), 'utf8'),
  ]);

  expect(component).toContain('useVersion()');
  expect(component).toContain('<ol');
  expect(component).toContain('<article');
  expect(component).toContain('<time dateTime={release.date}>');
  expect(component).not.toContain('useState');
  expect(component).not.toContain('fetch(');
  expect(styles).toContain('@media (max-width: 860px)');
  expect(styles).toContain('@media (max-width: 560px)');
  expect(styles).toContain('var(--ui-line-strong)');
  expect(englishPage).toContain("# What's new");
  expect(chinesePage).toContain('# 更新日志');
  expect(englishHome).toContain("[What's new](./changelog.html)");
  expect(chineseHome).toContain('[更新日志](./changelog.html)');
  expect(visualContract).toContain(
    'documentation changelog stays scannable, localized, and version-aware',
  );
  expect(a3sTestContract).toContain(
    'scenario "scan-version-aware-release-notes"',
  );
  expect(a3sTestContract).toContain('scenario "scan-release-notes-on-phone"');
  expect(packageManifest).toContain('playground:visual:docs-changelog');
});

test('keeps local evidence links aligned with each documentation language', () => {
  const paragraphMarkRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.52.0',
  );
  const shapeRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.53.0',
  );
  const compareRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.51.0',
  );
  const moveRevisionRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.49.0',
  );
  const selectionToolbarRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.48.1',
  );
  const dependentListRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.44.0',
  );
  const textBoxRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.46.0',
  );
  const commonFieldsRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.47.0',
  );
  const contentControlRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.48.0',
  );
  const pictureTransformRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.45.0',
  );
  const customFormulaRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.43.0',
  );
  const validationRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.41.0',
  );
  const numberingRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.40.0',
  );
  const presentationRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.39.0',
  );
  const writerRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.38.0',
  );
  const imeRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.37.2',
  );
  const animationRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.34.0',
  );
  const pdfRelease = OFFICE_RELEASE_NOTES.find(
    ({ version }) => version === '0.33.0',
  );

  expect(selectionToolbarRelease?.links[0]?.href).toEqual({
    en: './components/document.html#selection-toolbar-controls',
    zh: './components/document.html#选择工具栏控件',
  });
  expect(shapeRelease?.links[0]?.href).toEqual({
    en: './components/document.html#built-in-editable-text-boxes',
    zh: './components/document.html#可编辑文本框',
  });
  expect(shapeRelease?.links[2]?.href).toEqual({
    en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.53.0',
    zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.53.0',
  });
  expect(paragraphMarkRelease?.links[0]?.href).toEqual({
    en: './components/document.html#whole-paragraph-mark-revisions',
    zh: './components/document.html#整段段落标记修订',
  });
  expect(paragraphMarkRelease?.links[2]?.href).toEqual({
    en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.52.0',
    zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.52.0',
  });
  expect(compareRelease?.links[0]?.href).toEqual({
    en: './components/document.html#document-compare-and-combine',
    zh: './components/document.html#文档比较与合并',
  });
  expect(compareRelease?.links[1]?.href).toEqual({
    en: './components/collaboration.html#synchronize-move-revisions',
    zh: './components/collaboration.html#同步移动修订',
  });
  expect(compareRelease?.links[2]?.href).toEqual({
    en: 'https://github.com/A3S-Lab/Office/releases/tag/v0.51.0',
    zh: 'https://github.com/A3S-Lab/Office/releases/tag/v0.51.0',
  });
  expect(moveRevisionRelease?.links[0]?.href).toEqual({
    en: './components/document.html#move-revisions',
    zh: './components/document.html#移动修订',
  });
  expect(moveRevisionRelease?.links[1]?.href).toEqual({
    en: './components/collaboration.html#synchronize-move-revisions',
    zh: './components/collaboration.html#同步移动修订',
  });
  expect(dependentListRelease?.links[0]?.href).toEqual({
    en: './components/spreadsheet.html#dependent-dropdown-lists',
    zh: './components/spreadsheet.html#依赖下拉列表',
  });
  expect(textBoxRelease?.links[0]?.href).toEqual({
    en: './components/document.html#built-in-editable-text-boxes',
    zh: './components/document.html#可编辑文本框',
  });
  expect(commonFieldsRelease?.links[0]?.href).toEqual({
    en: './components/document.html#common-live-fields',
    zh: './components/document.html#常用实时字段',
  });
  expect(contentControlRelease?.links[0]?.href).toEqual({
    en: './components/document.html#built-in-content-controls',
    zh: './components/document.html#原生内容控件',
  });
  expect(pictureTransformRelease?.links[0]?.href).toEqual({
    en: './components/document.html#built-in-picture-properties',
    zh: './components/document.html#图片属性',
  });
  expect(customFormulaRelease?.links[0]?.href).toEqual({
    en: './components/spreadsheet.html#formula-conditional-formatting',
    zh: './components/spreadsheet.html#公式条件格式',
  });
  expect(numberingRelease?.links[0]?.href).toEqual({
    en: './components/document.html#ordered-list-numbering-revisions',
    zh: './components/document.html#有序列表编号修订',
  });
  expect(validationRelease?.links[0]?.href).toEqual({
    en: './components/spreadsheet.html#office-style-error-alert-branches',
    zh: './components/spreadsheet.html#与-office-一致的错误警告分支',
  });
  expect(presentationRelease?.links[0]?.href).toEqual({
    en: './components/presentation.html#entrance-and-exit-animations',
    zh: './components/presentation.html#进入与退出动画',
  });
  expect(writerRelease?.links[0]?.href).toEqual({
    en: './components/document.html#native-opentype-typography',
    zh: './components/document.html#原生-opentype-排版',
  });
  expect(imeRelease?.links[1]?.href).toEqual({
    en: './components/markdown.html#visual-editor-ime-behavior',
    zh: './components/markdown.html#可视化编辑器的输入法行为',
  });
  expect(animationRelease?.links[0]?.href).toEqual({
    en: './components/presentation.html#entrance-animations',
    zh: './components/presentation.html#入场动画',
  });
  expect(pdfRelease?.links[0]?.href).toEqual({
    en: './components/pdf.html#page-organization',
    zh: './components/pdf.html#页面组织',
  });
});
