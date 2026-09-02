import { access, readdir, readFile } from 'node:fs/promises';
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
const repositoryRoot = path.resolve(import.meta.dirname, '..');

function documentationComponentHref(
  version: string,
  component: string,
): string {
  const extension =
    version === 'latest' ||
    version === '0.43.0' ||
    version === '0.42.0' ||
    version === '0.41.0' ||
    version === '0.40.0' ||
    version === '0.39.0' ||
    version === '0.38.1' ||
    version === '0.38.0' ||
    version === '0.37.5' ||
    version === '0.37.4' ||
    version === '0.37.3' ||
    version === '0.37.2' ||
    version === '0.37.1' ||
    version === '0.37.0' ||
    version === '0.36.0' ||
    version === '0.34.0' ||
    version === '0.33.0' ||
    version === '0.32.0' ||
    version === '0.31.0' ||
    version === '0.30.0' ||
    version === '0.29.0' ||
    version === '0.28.0' ||
    version === '0.27.0' ||
    version === '0.26.0' ||
    version === '0.25.0' ||
    version === '0.24.0' ||
    version === '0.23.0' ||
    version === '0.22.0' ||
    version === '0.21.0' ||
    version === '0.20.0' ||
    version === '0.19.0' ||
    version === '0.18.0' ||
    version === '0.17.0' ||
    version === '0.16.0' ||
    version === '0.15.0' ||
    version === '0.14.0' ||
    version === '0.13.1' ||
    version === '0.13.0' ||
    version === '0.12.0' ||
    version === '0.11.0' ||
    version === '0.10.0' ||
    version === '0.9.2' ||
    version === '0.9.1' ||
    version === '0.9.0' ||
    version === '0.8.1' ||
    version === '0.8.0' ||
    version === '0.7.3'
      ? 'html'
      : 'mdx';
  return `./components/${component}.${extension}`;
}

test('uses Simplified Chinese and latest as stable documentation defaults', () => {
  expect(DOCUMENTATION_DEFAULT_LANGUAGE).toBe('zh');
  expect(DOCUMENTATION_LOCALES.map(({ lang }) => lang)).toEqual(['zh', 'en']);
  expect(DOCUMENTATION_DEFAULT_VERSION).toBe('latest');
  expect(DOCUMENTATION_VERSIONS).toEqual([
    'latest',
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
    '0.34.0',
    '0.33.0',
    '0.32.0',
    '0.31.0',
    '0.30.0',
    '0.29.0',
    '0.28.0',
    '0.27.0',
    '0.26.0',
    '0.25.0',
    '0.24.0',
    '0.23.0',
    '0.22.0',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
    '0.8.1',
    '0.8.0',
    '0.7.3',
    '0.7.2',
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

test('documents host-owned file action icons and destructive semantics in the current release', async () => {
  const [latestEnglish, latestChinese, releaseEnglish, releaseChinese] =
    await Promise.all([
      readFile(
        path.join(documentationRoot, 'latest/en/components/react.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, 'latest/zh/components/react.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, '0.37.1/en/components/react.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, '0.37.1/zh/components/react.mdx'),
        'utf8',
      ),
    ]);

  for (const document of [latestEnglish, releaseEnglish]) {
    expect(document).toContain('## Host-owned file actions');
    expect(document).toContain('OfficeFileAction');
    expect(document).toContain('danger: true');
    expect(document).toContain('neutral file glyph');
  }
  for (const document of [latestChinese, releaseChinese]) {
    expect(document).toContain('## 宿主持有的文件操作');
    expect(document).toContain('OfficeFileAction');
    expect(document).toContain('danger: true');
    expect(document).toContain('中性文件图标');
  }
});

test('documents controlled rich-text IME settlement in the current release', async () => {
  const [
    latestEnglishDocument,
    latestChineseDocument,
    releaseEnglishDocument,
    releaseChineseDocument,
    latestEnglishMarkdown,
    latestChineseMarkdown,
    releaseEnglishArchitecture,
    releaseChineseArchitecture,
  ] = await Promise.all([
    readFile(
      path.join(documentationRoot, 'latest/en/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.37.2/en/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.37.2/zh/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/components/markdown.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/markdown.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.37.2/en/browser-editor-architecture.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.37.2/zh/browser-editor-architecture.md'),
      'utf8',
    ),
  ]);

  for (const document of [latestEnglishDocument, releaseEnglishDocument]) {
    expect(document).toContain('## IME and controlled updates');
    expect(document).toContain('phonetic pre-edit text');
    expect(document).toContain('defers that replacement');
  }
  for (const document of [latestChineseDocument, releaseChineseDocument]) {
    expect(document).toContain('## 输入法与受控更新');
    expect(document).toContain('拼音等预编辑文字');
    expect(document).toContain('把替换延后');
  }
  expect(latestEnglishMarkdown).toContain('## Visual-editor IME behavior');
  expect(latestEnglishMarkdown).toContain('raw Pinyin');
  expect(latestChineseMarkdown).toContain('## 可视化编辑器的输入法行为');
  expect(releaseEnglishArchitecture).toContain(
    '## Controlled rich-text IME boundary',
  );
  expect(releaseEnglishArchitecture).toContain(
    'creates no timers or extra renders for ordinary keyboard input',
  );
  expect(releaseChineseArchitecture).toContain('## 受控富文本输入法边界');
});

test('builds a product home beside, rather than inside, the versioned docs site', async () => {
  const [
    productHome,
    productTheme,
    productThemeEntry,
    productHomeStyles,
    productConfig,
    docsConfig,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'website/product/index.mdx'), 'utf8'),
    readFile(
      path.join(
        repositoryRoot,
        'website/product-theme/components/HomeLayout.tsx',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'website/product-theme/index.tsx'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'website/product-theme/product-home.css'),
      'utf8',
    ),
    readFile(path.join(repositoryRoot, 'website/rspress.config.ts'), 'utf8'),
    readFile(
      path.join(repositoryRoot, 'website/rspress.docs.config.ts'),
      'utf8',
    ),
  ]);

  expect(productHome).toContain('pageType: home');
  expect(productHome).toContain('sidebar: false');
  expect(productTheme).toContain('className="docs-home office-product-home"');
  expect(productTheme).not.toContain('className="docs-home-assurance"');
  expect(productTheme).not.toContain(
    'className="docs-home-hero__latest docs-home-latest"',
  );
  expect(productThemeEntry).toContain(
    "export { Nav } from '../theme/components/Nav';",
  );
  expect(productThemeEntry).toContain("import '../theme/docs-content.css';");
  expect(productThemeEntry).toContain("import './product-home.css';");
  expect(productHomeStyles).toContain('html .docs-home.office-product-home');
  expect(productHomeStyles).toContain(
    '.office-product-home .docs-home-collaboration',
  );
  expect(productHomeStyles).not.toContain('var(--home-night)');
  expect(productConfig).toContain(
    "root: path.resolve(import.meta.dirname, 'product')",
  );
  expect(productConfig).toContain("{ text: '文档', link: '/docs/' }");
  expect(productConfig).toContain(
    "{ text: 'Playground', link: '/playground/' }",
  );
  expect(docsConfig).toContain(
    "root: path.resolve(import.meta.dirname, '../docs')",
  );
  expect(docsConfig).toContain('const docsBase =');
  expect(docsConfig).toContain('base: docsBase');
});

test('points README documentation links at the independent docs deployment', async () => {
  const readme = await readFile(path.join(repositoryRoot, 'README.md'), 'utf8');

  expect(readme).toContain('https://a3s-lab.github.io/Office/docs/');
  expect(readme).toContain('https://a3s-lab.github.io/Office/docs/components/');
  expect(readme).toContain('https://a3s-lab.github.io/Office/docs/automation/');
  expect(readme).not.toMatch(
    /https:\/\/a3s-lab\.github\.io\/Office\/(?:components|automation|en\/components|0\.\d)/,
  );
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

test('keeps public documentation on the neutral Traditional Office baseline', async () => {
  const versionedDocumentation = (
    await readdir(documentationRoot, { recursive: true })
  )
    .filter((file) => file.endsWith('.md') || file.endsWith('.mdx'))
    .map((file) => path.join(documentationRoot, file));
  const repositoryDocumentation = [
    'CHANGELOG.md',
    'COLLABORATION_ROADMAP.md',
    'PRODUCT.md',
    'README.md',
    'ROADMAP.md',
    'tests/e2e/README.md',
    'visual-tests/README.md',
  ].map((file) => path.join(repositoryRoot, file));

  for (const file of [...repositoryDocumentation, ...versionedDocumentation]) {
    const contents = await readFile(file, 'utf8');
    expect(contents).not.toMatch(/\bwps\b/i);
  }
});

test('keeps the README focused on the five editor surfaces and their explicit boundaries', async () => {
  const readme = await readFile(path.join(repositoryRoot, 'README.md'), 'utf8');

  for (const editor of [
    'Document',
    'Spreadsheet',
    'Presentation',
    'PDF',
    'Markdown',
  ]) {
    expect(readme).toContain(`### ${editor}`);
  }

  expect(readme).toContain('## Five surfaces, one product boundary');
  expect(readme).toContain('## Capabilities and boundaries');
  expect(readme).toContain('[capability roadmap](./ROADMAP.md)');
  expect(readme).not.toContain('Traditional Office baseline');
});

test('routes the concise README and documentation homes to the current release story', async () => {
  const [readme, englishHome, chineseHome] = await Promise.all([
    readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
    readFile(path.join(documentationRoot, 'latest/en/index.mdx'), 'utf8'),
    readFile(path.join(documentationRoot, 'latest/zh/index.mdx'), 'utf8'),
  ]);

  expect(readme).toContain('## Current release');
  expect(readme).toContain('Version `0.43.0`');
  expect(readme).toContain(
    "[What's new](https://a3s-lab.github.io/Office/docs/changelog.html)",
  );
  expect(readme).toContain('[changelog](./CHANGELOG.md)');
  expect(readme).toContain('[capability roadmap](./ROADMAP.md)');
  expect(readme).toContain(
    '[live Playground](https://a3s-lab.github.io/Office/playground/)',
  );

  expect(englishHome).toContain("## What's new on `main` (0.43.0)");
  expect(englishHome).toContain("[What's new](./changelog.html)");
  expect(englishHome).toContain(
    'spreadsheet.html#formula-conditional-formatting',
  );
  expect(englishHome).toContain(
    'spreadsheet.html#office-style-error-alert-branches',
  );
  expect(englishHome).toContain(
    'document.html#ordered-list-numbering-revisions',
  );
  expect(englishHome).toContain(
    'presentation.html#entrance-and-exit-animations',
  );
  expect(englishHome).toContain('document.html#native-opentype-typography');
  expect(englishHome).toContain(
    'spreadsheet.html#xlsx-1904-date-system-retention',
  );

  expect(chineseHome).toContain('## `main` 更新内容（0.43.0）');
  expect(chineseHome).toContain('[更新日志](./changelog.html)');
  expect(chineseHome).toContain('spreadsheet.html#公式条件格式');
  expect(chineseHome).toContain(
    'spreadsheet.html#与-office-一致的错误警告分支',
  );
  expect(chineseHome).toContain('document.html#有序列表编号修订');
  expect(chineseHome).toContain('presentation.html#进入与退出动画');
  expect(chineseHome).toContain('document.html#原生-opentype-排版');
  expect(chineseHome).toContain('spreadsheet.html#xlsx-1904-日期系统保留');
});

test('publishes Writer numbering revisions across implementation, native collaboration, docs, and release evidence', async () => {
  const [
    changelog,
    roadmap,
    product,
    englishDocument,
    chineseDocument,
    englishCollaboration,
    chineseCollaboration,
    releaseEnglishDocument,
    releaseChineseDocument,
    releaseEnglishHome,
    releaseChineseHome,
    tracking,
    importer,
    exporter,
    nativeKinds,
    nativeTests,
    docxTests,
    aclSuite,
    packageManifest,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/components/collaboration.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/collaboration.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.40.0/en/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.40.0/zh/components/document.mdx'),
      'utf8',
    ),
    readFile(path.join(documentationRoot, '0.40.0/en/index.mdx'), 'utf8'),
    readFile(path.join(documentationRoot, '0.40.0/zh/index.mdx'), 'utf8'),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/work-document-numbering-change-tracking.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/work-docx-numbering-change-import.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/work-docx-numbering-change-export.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'crates/core/src/collaboration/types/mutation.rs',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'crates/core/src/collaboration/tests/document_suggestions.rs',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/docx-numbering-changes.test.ts'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/word-numbering-revision.acl'),
      'utf8',
    ),
    readFile(path.join(repositoryRoot, 'package.json'), 'utf8'),
  ]);

  expect(changelog).toContain('## 0.40.0 - 2026-09-02');
  expect(changelog).toContain('Writer ordered-list numbering revisions');
  expect(roadmap).toContain('common single-level `w:numberingChange`');
  expect(product).toContain('The sixty-first Writer milestone');
  expect(releaseEnglishHome).toContain('Released in 0.40.0');
  expect(releaseEnglishHome).toContain(
    'document.html#ordered-list-numbering-revisions',
  );
  expect(releaseChineseHome).toContain('0.40.0 发布');
  expect(releaseChineseHome).toContain('document.html#有序列表编号修订');

  for (const source of [
    englishDocument,
    chineseDocument,
    englishCollaboration,
    chineseCollaboration,
    releaseEnglishDocument,
    releaseChineseDocument,
  ]) {
    expect(source).toContain('w:numberingChange');
    expect(source).toContain('changeKind: "numbering"');
  }
  for (const source of [
    englishDocument,
    chineseDocument,
    releaseEnglishDocument,
    releaseChineseDocument,
  ]) {
    expect(source).toContain('65,536');
  }
  expect(englishDocument).toContain('## Ordered-list numbering revisions');
  expect(chineseDocument).toContain('## 有序列表编号修订');

  expect(tracking).toContain('ReplaceAroundStep');
  expect(importer).toContain('MAX_NUMBERING_CHANGES = 65_536');
  expect(exporter).toContain('MAX_NUMBERING_CHANGE_PATCHES = 65_536');
  expect(nativeKinds).toContain('Numbering');
  expect(nativeKinds).toContain('Self::Numbering => "numbering"');
  expect(nativeTests).toContain(
    'browser_numbering_revisions_survive_restart_and_suggest_mode',
  );
  expect(docxTests).toContain(
    'imports one contiguous native numbering revision as an atomic review card',
  );
  expect(aclSuite).toContain('suite "word-numbering-revision"');
  expect(packageManifest).toContain('test:e2e:numbering-revision');
});

test('publishes Spreadsheet validation alert branches across implementation, docs, and release evidence', async () => {
  const [
    changelog,
    roadmap,
    product,
    readme,
    english,
    chinese,
    releaseData,
    frozenEnglish,
    frozenChinese,
    helper,
    editor,
    fortunePatch,
    unitTest,
    visualTest,
    aclSuite,
    packageManifest,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'website/theme/release-notes-data.ts'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.41.0/en/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.41.0/zh/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-data-validation-interaction.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-editor.tsx',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'patches/@fortune-sheet%2Fcore@1.0.4.patch'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'tests/spreadsheet-data-validation-interaction.test.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'visual-tests/spreadsheet-data-validation.functional.spec.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/spreadsheet-data-validation.acl'),
      'utf8',
    ),
    readFile(path.join(repositoryRoot, 'package.json'), 'utf8'),
  ]);

  expect(changelog).toContain('## 0.41.0 - 2026-09-02');
  expect(changelog).toContain('Aligned Spreadsheet data-validation alerts');
  expect(roadmap).toContain('Warning/Information confirmation branches');
  expect(product).toContain('Version 0.41.0 aligns direct and formula-bar');
  expect(readme).toContain('Version `0.41.0`');

  for (const source of [english, chinese, frozenEnglish, frozenChinese]) {
    expect(source).toContain('errorStyle');
    expect(source).toContain('Warning');
    expect(source).toContain('Information');
    expect(source).toContain('Undo');
  }
  expect(english).toContain('## Office-style error-alert branches');
  expect(chinese).toContain('## 与 Office 一致的错误警告分支');
  expect(releaseData).toContain("version: '0.41.0'");
  expect(releaseData).toContain('office-style-error-alert-branches');
  expect(helper).toContain('spreadsheetDataValidationConfirmLabels');
  expect(helper).toContain('cloneSpreadsheetDataValidationItem');
  expect(editor).toContain('beforeDataValidation: queueDataValidationEdit');
  expect(editor).toContain('skipDataValidation: true');
  expect(fortunePatch).toContain('skipDataValidation?: boolean');
  expect(fortunePatch).toContain('beforeDataValidation');
  expect(unitTest).toContain('maps Office error styles');
  expect(visualTest).toContain(
    'follows Traditional Office error-alert branches',
  );
  expect(aclSuite).toContain('scenario "follow-error-alert-branches"');
  expect(packageManifest).toContain('test:e2e:spreadsheet-data-validation');
});

test('publishes Spreadsheet custom-formula validation across the current and frozen release', async () => {
  const [
    changelog,
    roadmap,
    product,
    readme,
    english,
    chinese,
    releaseData,
    frozenEnglish,
    frozenChinese,
    evaluator,
    dialog,
    xlsx,
    visual,
    acl,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'website/theme/release-notes-data.ts'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.42.0/en/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.42.0/zh/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-data-validation-custom.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-data-validation-dialog.tsx',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/work-xlsx-interop.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'visual-tests/spreadsheet-data-validation.functional.spec.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/spreadsheet-data-validation.acl'),
      'utf8',
    ),
  ]);

  expect(changelog).toContain('## 0.42.0 - 2026-09-02');
  expect(roadmap).toContain('bounded local custom-formula rules');
  expect(product).toContain(
    'Version 0.42.0 adds the follow-up custom-formula rule',
  );
  expect(readme).toContain('Version `0.42.0`');
  for (const source of [english, chinese, frozenEnglish, frozenChinese]) {
    expect(source).toContain('1,024');
    expect(source).toContain('255');
  }
  expect(english).toContain('Custom formula');
  expect(frozenEnglish).toContain('Custom formula');
  expect(chinese).toContain('自定义公式');
  expect(frozenChinese).toContain('自定义公式');
  expect(english).toContain('## Custom formulas');
  expect(chinese).toContain('## 自定义公式');
  expect(releaseData).toContain("version: '0.42.0'");
  expect(releaseData).toContain('custom-formulas');
  expect(evaluator).toContain(
    'MAX_SPREADSHEET_CUSTOM_VALIDATION_REFERENCED_CELLS',
  );
  expect(dialog).toContain('SpreadsheetDataValidationCustomFormulaField');
  expect(xlsx).toContain("custom: 'custom'");
  expect(visual).toContain("'LEN(E2)>0'");
  expect(acl).toContain('custom-formula');
});

test('keeps every documentation index separate from the product home surface', async () => {
  const documentationIndexes = await Promise.all(
    DOCUMENTATION_VERSIONS.flatMap((version) =>
      DOCUMENTATION_LOCALES.map(async ({ lang }) => ({
        lang,
        version,
        contents: await readFile(
          path.join(documentationRoot, version, lang, 'index.mdx'),
          'utf8',
        ),
      })),
    ),
  );

  for (const { contents: index } of documentationIndexes) {
    expect(index).not.toContain('docs-home-hero');
    expect(index).not.toContain('docs-home-collaboration');
    expect(index).not.toContain('docs-home-final');
    expect(index).not.toContain('docs-home-system-window');
    expect(index).not.toContain('pageType: home');
    expect(index).not.toContain('<PlaygroundLink');
    expect(index).not.toContain('在线体验');
    expect(index).not.toContain('Open the Playground');
    expect(index).not.toContain('Product storytelling');
    expect(index).not.toContain('product-home demonstrations');
    expect(index).not.toContain('产品首页');
    expect(index).toContain('./guide/index.');
    expect(index).toContain('./components/index.');
    expect(index).toContain('./automation/index.');
  }
});

test('publishes Presentation entrance and exit animations across implementation, docs, Playground, and release evidence', async () => {
  const [
    changelog,
    roadmap,
    product,
    english,
    chinese,
    releaseEnglish,
    releaseChinese,
    releaseEnglishHome,
    releaseChineseHome,
    templates,
    workspaceHome,
    pptxTest,
    visualSpec,
    aclSuite,
    discoverabilityAcl,
    packageManifest,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/components/presentation.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/presentation.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.39.0/en/components/presentation.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.39.0/zh/components/presentation.mdx'),
      'utf8',
    ),
    readFile(path.join(documentationRoot, '0.39.0/en/index.mdx'), 'utf8'),
    readFile(path.join(documentationRoot, '0.39.0/zh/index.mdx'), 'utf8'),
    readFile(
      path.join(repositoryRoot, 'src/internal/features/work/work-templates.ts'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'playground/src/workspace-home.tsx'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/pptx-animation-round-trip.test.ts'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'visual-tests/presentation-animation.functional.spec.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/presentation-animation.acl'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'tests/e2e/playground-template-discoverability.acl',
      ),
      'utf8',
    ),
    readFile(path.join(repositoryRoot, 'package.json'), 'utf8'),
  ]);

  expect(changelog).toContain('Added editable Presentation exit animations');
  expect(changelog).toContain('## 0.39.0 - 2026-09-01');
  expect(roadmap).toContain('| Object animations and triggers | **Partial**');
  expect(product).toContain('## Current Presentation Milestone');

  for (const document of [english, chinese, releaseEnglish, releaseChinese]) {
    expect(document).toContain('WorkSlide.animations');
    for (const effect of [
      'appear',
      'fade',
      'fly-in',
      'zoom',
      'disappear',
      'fade-out',
      'fly-out',
      'zoom-out',
    ]) {
      expect(document).toContain(`'${effect}'`);
    }
    expect(document).toContain("'on-click'");
    expect(document).toContain("'with-previous'");
    expect(document).toContain("'after-previous'");
    expect(document).toContain('256');
    expect(document).toContain('60,000');
    expect(document).toContain('PresentationML');
  }
  expect(english).toContain('## Entrance and exit animations');
  expect(chinese).toContain('## 进入与退出动画');
  expect(releaseEnglishHome).toContain('Released in 0.39.0');
  expect(releaseEnglishHome).toContain(
    'presentation.html#entrance-and-exit-animations',
  );
  expect(releaseChineseHome).toContain('0.39.0 发布');
  expect(releaseChineseHome).toContain('presentation.html#进入与退出动画');

  expect(templates).toContain("id: 'animated-deck'");
  expect(templates).toContain("effect: 'appear'");
  expect(templates).toContain("effect: 'fade'");
  expect(templates).toContain("effect: 'fly-in'");
  expect(templates).toContain("effect: 'zoom'");
  expect(templates).toContain("effect: 'disappear'");
  expect(templates).toContain("effect: 'fade-out'");
  expect(templates).toContain("effect: 'fly-out'");
  expect(templates).toContain("effect: 'zoom-out'");
  expect(workspaceHome).toContain('data-template-id={template.id}');
  expect(workspaceHome).not.toContain('playground-latest-capabilities');
  expect(pptxTest).toContain(
    'round-trips supported entrance and exit animations through native PPTX timing trees',
  );
  expect(visualSpec).toContain(
    'Presentation entrance and exit animations author and play ordered cues',
  );
  expect(aclSuite).toContain('suite "office-presentation-animation"');
  expect(discoverabilityAcl).toContain(
    "button[data-template-id='animated-deck']",
  );
  expect(discoverabilityAcl).toContain(
    'navigate "open-entrance-exit-animation-documentation"',
  );
  expect(discoverabilityAcl).toContain("main h2[id='进入与退出动画']");
  expect(packageManifest).toContain(
    '"playground:visual:presentation-animation"',
  );
  expect(packageManifest).toContain('"test:e2e:presentation-animation"');
});

test('publishes PDF page organization across docs, Playground, and release evidence', async () => {
  const [
    changelog,
    roadmap,
    product,
    english,
    chinese,
    releaseEnglish,
    releaseChinese,
    workspaceHome,
    visualSpec,
    aclSuite,
    discoverabilityAcl,
    e2eGuide,
    packageManifest,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/components/pdf.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/pdf.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.33.0/en/components/pdf.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.33.0/zh/components/pdf.mdx'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'playground/src/workspace-home.tsx'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'visual-tests/pdf-page-organization.functional.spec.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/pdf-page-organization.acl'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'tests/e2e/playground-template-discoverability.acl',
      ),
      'utf8',
    ),
    readFile(path.join(repositoryRoot, 'tests/e2e/README.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'package.json'), 'utf8'),
  ]);

  expect(changelog).toContain('## 0.33.0 - 2026-08-25');
  expect(roadmap).toContain(
    '| Insert, delete, rotate, reorder, extract, merge, and split pages | **Supported with boundaries**',
  );
  expect(product).toContain('## Current PDF Milestone');

  for (const document of [english, chinese, releaseEnglish, releaseChinese]) {
    expect(document).toContain('onPageExport');
    expect(document).toContain('PdfPageOrganizationExport');
    expect(document).toContain('256 MiB');
    expect(document).toContain('128 MiB');
    expect(document).toContain('4,096');
  }
  expect(english).toContain('## Page organization');
  expect(english).toContain('dedicated Web Worker');
  expect(english).toContain('native PDF history first');
  expect(releaseEnglish).toContain('## Page organization');
  expect(chinese).toContain('## 页面组织');
  expect(chinese).toContain('独立 Web Worker');
  expect(chinese).toContain('原生 PDF 历史优先');
  expect(releaseChinese).toContain('## 页面组织');

  expect(workspaceHome).toContain('PDF 编辑器');
  expect(workspaceHome).toContain('onClick={onOpenPdf}');
  expect(visualSpec).toContain('PDF page organization mutates, exports, saves');
  expect(aclSuite).toContain('suite "office-pdf-page-organization"');
  expect(discoverabilityAcl).toContain(
    'navigate "open-pdf-page-organization-documentation"',
  );
  expect(discoverabilityAcl).toContain("main h2[id='页面组织']");
  expect(e2eGuide).toContain('bun run test:e2e:pdf-page-organization');
  expect(packageManifest).toContain(
    '"playground:visual:pdf-page-organization"',
  );
  expect(packageManifest).toContain('"test:e2e:pdf-page-organization"');
});

test('publishes native Writer proofing languages in docs, roadmap, and Playground guidance', async () => {
  const [
    changelog,
    roadmap,
    english,
    chinese,
    englishRoadmap,
    chineseRoadmap,
    englishArchitecture,
    chineseArchitecture,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/browser-editor-architecture.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/browser-editor-architecture.md'),
      'utf8',
    ),
  ]);

  expect(changelog).toContain('**设置校对语言**');
  expect(changelog).toContain('RustyBuzz');
  expect(roadmap).toContain('independent native `w:lang`');
  expect(english).toContain('## Native proofing languages');
  expect(chinese).toContain('## 原生校对语言');
  for (const document of [
    english,
    chinese,
    englishRoadmap,
    chineseRoadmap,
    englishArchitecture,
    chineseArchitecture,
  ]) {
    expect(document).toContain('`w:lang`');
    expect(document).toContain('`w:noProof`');
    expect(document).toContain('RustyBuzz');
  }
});

test('publishes Writer Table of Contents in docs, roadmap, and Playground guidance', async () => {
  const [changelog, roadmap, english, chinese, templates, playground] =
    await Promise.all([
      readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
      readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
      readFile(
        path.join(documentationRoot, 'latest/en/components/document.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, 'latest/zh/components/document.mdx'),
        'utf8',
      ),
      readFile(
        path.join(
          repositoryRoot,
          'src/internal/features/work/work-templates.ts',
        ),
        'utf8',
      ),
      readFile(
        path.join(repositoryRoot, 'playground/src/workspace-home.tsx'),
        'utf8',
      ),
    ]);

  expect(changelog).toContain('typed Writer Table of Contents block');
  expect(roadmap).toContain(
    '**Supported**: shared semantic-heading/native-outline model',
  );
  expect(english).toContain('## Native table of contents');
  expect(chinese).toContain('## 原生可更新目录');
  expect(templates).toContain("id: 'table-of-contents'");
  expect(playground).toContain('officeTemplates.map((template) =>');
  for (const document of [english, chinese]) {
    expect(document).toContain('512');
    expect(document).toContain('`w:sdt`');
    expect(document).toContain('`TOC`');
  }
});

test('publishes native Writer indexes in docs, roadmap, and Playground guidance', async () => {
  const [
    changelog,
    roadmap,
    english,
    chinese,
    releaseEnglish,
    releaseChinese,
    englishArchitecture,
    chineseArchitecture,
    englishRoadmap,
    chineseRoadmap,
    templates,
    playground,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.31.0/en/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.31.0/zh/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/browser-editor-architecture.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/browser-editor-architecture.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'src/internal/features/work/work-templates.ts'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'playground/src/workspace-home.tsx'),
      'utf8',
    ),
  ]);

  expect(changelog).toContain('native Writer index authoring');
  expect(roadmap).toContain('| Native index authoring | **Supported**');
  expect(english).toContain('## Native document index');
  expect(chinese).toContain('## 原生文档索引');
  expect(releaseEnglish).toContain('## Native document index');
  expect(releaseChinese).toContain('## 原生文档索引');
  expect(templates).toContain("id: 'document-index'");
  expect(playground).toContain('officeTemplates.map((template) =>');
  for (const document of [
    english,
    chinese,
    releaseEnglish,
    releaseChinese,
    englishArchitecture,
    chineseArchitecture,
    englishRoadmap,
    chineseRoadmap,
  ]) {
    expect(document).toContain('`XE`');
    expect(document).toContain('`INDEX`');
    expect(document).toContain('512');
  }
});

test('publishes Writer document compare and combine across product and documentation surfaces', async () => {
  const [
    changelog,
    roadmap,
    product,
    english,
    chinese,
    releaseEnglish,
    releaseChinese,
    englishArchitecture,
    chineseArchitecture,
    englishRoadmap,
    chineseRoadmap,
    e2eGuide,
    templates,
    playground,
    packageManifest,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.32.0/en/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.32.0/zh/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/browser-editor-architecture.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/browser-editor-architecture.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(path.join(repositoryRoot, 'tests/e2e/README.md'), 'utf8'),
    readFile(
      path.join(repositoryRoot, 'src/internal/features/work/work-templates.ts'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'playground/src/workspace-home.tsx'),
      'utf8',
    ),
    readFile(path.join(repositoryRoot, 'package.json'), 'utf8'),
  ]);

  expect(changelog).toContain('deterministic Writer document comparison');
  expect(roadmap).toContain('| Compare/combine documents | **Partial**');
  expect(product).toContain('The forty-fifth milestone');
  expect(english).toContain('## Document compare and combine');
  expect(chinese).toContain('## 文档比较与合并');
  expect(releaseEnglish).toContain('## Document compare and combine');
  expect(releaseChinese).toContain('## 文档比较与合并');
  expect(englishArchitecture).toContain(
    'Document comparison is a pure planning boundary',
  );
  expect(chineseArchitecture).toContain(
    '文档比较在编辑器事务之前形成纯规划边界',
  );
  expect(englishRoadmap).toContain(
    'Writer document compare/combine now forms one completed bounded vertical slice',
  );
  expect(chineseRoadmap).toContain(
    'Writer 文档比较与合并现在形成一个已完成的有界纵向切片',
  );
  expect(e2eGuide).toContain('bun run test:e2e:writer-document-comparison');
  expect(templates).toContain("id: 'document-comparison'");
  expect(playground).toContain('officeTemplates.map((template) =>');
  expect(packageManifest).toContain('"playground:visual:document-comparison"');
  expect(packageManifest).toContain('"test:e2e:writer-document-comparison"');

  for (const document of [english, chinese, releaseEnglish, releaseChinese]) {
    for (const evidence of [
      '1,024',
      '1,000,000',
      '`w:ins`',
      '`w:del`',
      '`w:rPrChange`',
      '`w:pPrChange`',
      'word-document-comparison.acl',
    ]) {
      expect(document).toContain(evidence);
    }
  }
});

test('publishes complete Spreadsheet data-validation settings in product documentation', async () => {
  const [changelog, roadmap, english, chinese, templates] = await Promise.all([
    readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'src/internal/features/work/work-templates.ts'),
      'utf8',
    ),
  ]);

  expect(changelog).toContain('input titles and messages');
  expect(changelog).toContain('home-page **新建 → 数据验证**');
  expect(roadmap).toContain('input and error-alert settings');
  expect(english).toContain('### Input and error settings');
  expect(english).toContain('**新建 → 数据验证**');
  expect(chinese).toContain('## 数据验证');
  expect(chinese).toContain('### 输入信息与错误警告设置');
  expect(chinese).toContain('**新建 → 数据验证**');
  expect(templates).toContain("id: 'data-validation'");
  for (const document of [english, chinese]) {
    expect(document).toContain('`allowBlank`');
    expect(document).toContain('`showDropdownArrow`');
    expect(document).toContain('`errorStyle`');
    expect(document).toContain('`hintTitle`');
  }
});

test('publishes editable formula conditional formatting across code, docs, and Playground', async () => {
  const [
    changelog,
    roadmap,
    product,
    readme,
    english,
    chinese,
    frozenEnglish,
    frozenChinese,
    releaseData,
    templates,
    unitTest,
    visualTest,
    aclSuite,
    packageManifest,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.43.0/en/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.43.0/zh/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'website/theme/release-notes-data.ts'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'src/internal/features/work/work-templates.ts'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'tests/spreadsheet-conditional-format-formula.test.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'visual-tests/spreadsheet-conditional-format.functional.spec.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/spreadsheet-conditional-format.acl'),
      'utf8',
    ),
    readFile(path.join(repositoryRoot, 'package.json'), 'utf8'),
  ]);

  expect(changelog).toContain('## 0.43.0 - 2026-09-02');
  expect(changelog).toContain(
    'editable Spreadsheet formula-conditional-format',
  );
  expect(roadmap).toContain('bounded local formula rules');
  expect(product).toContain('The sixty-second Spreadsheet milestone');
  expect(readme).toContain('Version `0.43.0`');
  for (const document of [english, chinese, frozenEnglish, frozenChinese]) {
    expect(document).toContain('1,024');
    expect(document).toContain('255');
  }
  expect(english).toContain('## Formula conditional formatting');
  expect(chinese).toContain('## 公式条件格式');
  expect(english).toContain('Stop if true');
  expect(chinese).toContain('匹配后停止');
  expect(releaseData).toContain("version: '0.43.0'");
  expect(releaseData).toContain('formula-conditional-formatting');
  expect(templates).toContain("id: 'conditional-format'");
  expect(templates).toContain('Limits!$A$2');
  expect(unitTest).toContain('publishes a formula conditional-format template');
  expect(visualTest).toContain('authors local formula conditional formatting');
  expect(aclSuite).toContain('suite "office-spreadsheet-conditional-format"');
  expect(packageManifest).toContain('test:e2e:spreadsheet-conditional-format');
  expect(packageManifest).toContain(
    'playground:visual:spreadsheet-conditional-format',
  );
});

test('documents the complete native Writer underline contract', async () => {
  const [roadmap, product, english, chinese] = await Promise.all([
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/components/document.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/document.mdx'),
      'utf8',
    ),
  ]);

  expect(roadmap).toContain('all 18 native `w:u` values');
  expect(product).toContain('The thirty-ninth milestone');
  for (const document of [english, chinese]) {
    for (const style of [
      'none',
      'single',
      'words',
      'double',
      'thick',
      'dotted',
      'dottedHeavy',
      'dash',
      'dashedHeavy',
      'dashLong',
      'dashLongHeavy',
      'dotDash',
      'dashDotHeavy',
      'dotDotDash',
      'dashDotDotHeavy',
      'wave',
      'wavyHeavy',
      'wavyDouble',
    ]) {
      expect(document).toContain(`\`${style}\``);
    }
    expect(document).toContain('Cmd/Ctrl+Shift+D');
    expect(document).toContain('Cmd/Ctrl+Shift+W');
  }
});

test('documents the complete native Writer strikethrough contract', async () => {
  const [roadmap, product, english, chinese, englishRoadmap, chineseRoadmap] =
    await Promise.all([
      readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
      readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
      readFile(
        path.join(documentationRoot, 'latest/en/components/document.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, 'latest/zh/components/document.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, 'latest/en/editor-quality-roadmap.md'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, 'latest/zh/editor-quality-roadmap.md'),
        'utf8',
      ),
    ]);

  expect(roadmap).toContain('native `w:strike` / `w:dstrike` state');
  expect(product).toContain('The fortieth milestone');
  for (const document of [english, englishRoadmap]) {
    expect(document).toContain('`none | single | double`');
    expect(document).toContain('`w:strike`');
    expect(document).toContain('`w:dstrike`');
    expect(document).toContain('`Mod+Shift+S`');
  }
  for (const document of [chinese, chineseRoadmap]) {
    expect(document).toContain('`none | single | double`');
    expect(document).toContain('`w:strike`');
    expect(document).toContain('`w:dstrike`');
    expect(document).toContain('`Mod+Shift+S`');
  }
});

test('keeps published documentation indexes frozen and visibly versioned', async () => {
  for (const version of [
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
    '0.34.0',
    '0.33.0',
    '0.32.0',
    '0.31.0',
    '0.28.0',
    '0.27.0',
    '0.26.0',
    '0.25.0',
    '0.24.0',
    '0.23.0',
    '0.22.0',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
    '0.8.1',
    '0.8.0',
    '0.7.3',
    '0.7.2',
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
    expect(english.toLowerCase()).toContain('frozen documentation');
  }
});

test('removes broken online Playground actions from frozen documentation indexes', async () => {
  for (const version of DOCUMENTATION_VERSIONS.filter(
    (candidate) => candidate !== 'latest',
  )) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const documentationIndex = await readFile(
        path.join(documentationRoot, version, lang, 'index.mdx'),
        'utf8',
      );

      expect(documentationIndex).not.toContain('<PlaygroundLink>');
      expect(documentationIndex).not.toContain('在线体验');
      expect(documentationIndex).not.toContain('Open the Playground');
    }
  }
});

test('uses deployable HTML targets in current release documentation indexes', async () => {
  for (const version of [
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
    '0.34.0',
    '0.33.0',
    '0.32.0',
    '0.31.0',
    '0.28.0',
    '0.27.0',
    '0.26.0',
    '0.25.0',
    '0.24.0',
    '0.23.0',
    '0.22.0',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
    '0.8.1',
    '0.8.0',
    '0.7.3',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const documentationIndex = await readFile(
        path.join(documentationRoot, version, lang, 'index.mdx'),
        'utf8',
      );

      expect(documentationIndex).not.toMatch(/href="[^"]+\.mdx(?:[?#][^"]*)?"/);
      expect(documentationIndex).toContain('./guide/index.html');
      expect(documentationIndex).toContain('./components/collaboration.html');
      expect(documentationIndex).toContain(
        './components/collaboration-server.html',
      );
    }
  }
});

test('publishes real-time collaboration as a bilingual first-class capability', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
    '0.8.1',
    '0.8.0',
    '0.7.3',
    '0.7.2',
    '0.7.1',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [
        documentationIndex,
        componentIndex,
        componentNavigation,
        collaboration,
      ] = await Promise.all([
        readFile(path.join(localeRoot, 'index.mdx'), 'utf8'),
        readFile(path.join(localeRoot, 'components/index.mdx'), 'utf8'),
        readFile(path.join(localeRoot, 'components/_meta.json'), 'utf8'),
        readFile(path.join(localeRoot, 'components/collaboration.mdx'), 'utf8'),
      ]);

      expect(documentationIndex).toContain(
        documentationComponentHref(version, 'collaboration'),
      );
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

test('publishes the runnable collaboration backend in latest releases', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
    '0.8.1',
    '0.8.0',
    '0.7.3',
    '0.7.2',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [documentationIndex, componentIndex, componentNavigation, server] =
        await Promise.all([
          readFile(path.join(localeRoot, 'index.mdx'), 'utf8'),
          readFile(path.join(localeRoot, 'components/index.mdx'), 'utf8'),
          readFile(path.join(localeRoot, 'components/_meta.json'), 'utf8'),
          readFile(
            path.join(localeRoot, 'components/collaboration-server.mdx'),
            'utf8',
          ),
        ]);

      expect(documentationIndex).toContain(
        documentationComponentHref(version, 'collaboration-server'),
      );
      expect(componentIndex).toContain('./collaboration-server.mdx');
      expect(componentNavigation).toContain('"collaboration-server"');
      expect(server).toContain('a3s-boot');
      expect(server).toContain('WebSocket');
      expect(server).toContain('Yrs');
    }
  }
});

test('documents ephemeral native agent presence in current releases', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [collaboration, server, cli] = await Promise.all([
        readFile(path.join(localeRoot, 'components/collaboration.mdx'), 'utf8'),
        readFile(
          path.join(localeRoot, 'components/collaboration-server.mdx'),
          'utf8',
        ),
        readFile(path.join(localeRoot, 'cli-reference.md'), 'utf8'),
      ]);

      for (const source of [collaboration, server, cli]) {
        expect(source).toContain('--actor-name');
        expect(source).toContain('outbound-awareness');
        expect(source).toContain('receive-awareness');
        expect(source).toContain('peer-left');
      }
    }
  }
});

test('documents durable Document comments in current releases', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
    '0.8.1',
    '0.8.0',
    '0.7.3',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [collaboration, server, cli] = await Promise.all([
        readFile(path.join(localeRoot, 'components/collaboration.mdx'), 'utf8'),
        readFile(
          path.join(localeRoot, 'components/collaboration-server.mdx'),
          'utf8',
        ),
        readFile(path.join(localeRoot, 'cli-reference.md'), 'utf8'),
      ]);

      for (const document of [collaboration, cli]) {
        expect(document).toContain('document-comment-create');
        expect(document).toContain('document-comment-reply');
        expect(document).toContain('document-comment-set-resolved');
        expect(document).toContain('document-comment-delete');
      }
      expect(collaboration).toContain('startUtf16');
      expect(collaboration).toContain('expectedTextId');
      expect(server).toContain('actorName');
      expect(server).toContain('comment');
      expect(server).toContain('FORBIDDEN');
    }
  }
});

test('documents attributed Document suggestions and native typed mutations', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
    '0.8.1',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [collaboration, server, document, cli] = await Promise.all([
        readFile(path.join(localeRoot, 'components/collaboration.mdx'), 'utf8'),
        readFile(
          path.join(localeRoot, 'components/collaboration-server.mdx'),
          'utf8',
        ),
        readFile(path.join(localeRoot, 'components/document.mdx'), 'utf8'),
        readFile(path.join(localeRoot, 'cli-reference.md'), 'utf8'),
      ]);

      for (const source of [collaboration, document]) {
        expect(source).toContain("mode: 'suggest'");
      }
      expect(collaboration).toContain('document.change-decisions');
      expect(server).toContain('FORBIDDEN');
      expect(server).toContain('Document `suggest`');
      expect(cli).toContain('document-suggestion-create');
      expect(cli).toContain('document-suggestion-decide');
      expect(cli).toContain('v3');
      expect(cli).toContain('changeDecisions');
    }
  }
});

test('documents collaborative character-formatting revisions', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [document, collaboration, server] = await Promise.all([
        readFile(path.join(localeRoot, 'components/document.mdx'), 'utf8'),
        readFile(path.join(localeRoot, 'components/collaboration.mdx'), 'utf8'),
        readFile(
          path.join(localeRoot, 'components/collaboration-server.mdx'),
          'utf8',
        ),
      ]);

      for (const source of [document, collaboration, server]) {
        expect(source).toContain('formatting');
        expect(source).toContain('w:rPrChange');
      }
      expect(collaboration).toContain('changeKind: "formatting"');
      expect(server).toContain('document-suggestion-create');
      expect(server).toContain('document-suggestion-decide');
    }
  }
});

test('documents complete native Writer OpenType typography in both current locales', async () => {
  for (const version of [
    'latest',
    '0.42.0',
    '0.41.0',
    '0.40.0',
    '0.39.0',
    '0.38.1',
    '0.38.0',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const source = await readFile(
        path.join(documentationRoot, version, lang, 'components/document.mdx'),
        'utf8',
      );
      expect(source).toContain('w14:ligatures');
      expect(source).toContain('w14:numForm');
      expect(source).toContain('w14:numSpacing');
      expect(source).toContain('w14:stylisticSets');
      expect(source).toContain('w14:cntxtAlts');
      expect(source).toContain('data-office-opentype-features');
      expect(source).toContain('Cmd/Ctrl+D');
    }
  }
});

test('documents native Writer character spacing in both current locales', async () => {
  for (const version of ['latest', '0.28.0', '0.27.0']) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const source = await readFile(
        path.join(documentationRoot, version, lang, 'components/document.mdx'),
        'utf8',
      );
      expect(source).toContain('w:spacing');
      expect(source).toContain('Cmd/Ctrl+D');
      expect(source).toContain('Worker/WASM');
    }
  }
});

test('documents native Writer kerning thresholds in both current locales', async () => {
  for (const version of ['latest', '0.28.0', '0.27.0']) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const source = await readFile(
        path.join(documentationRoot, version, lang, 'components/document.mdx'),
        'utf8',
      );
      expect(source).toContain('w:kern');
      expect(source).toContain('Cmd/Ctrl+D');
      expect(source).toContain('font-kerning');
      expect(source).toContain('Worker/WASM');
    }
  }
});

test('documents all native Writer emphasis marks in both current locales', async () => {
  const [roadmap, product, englishRoadmap, chineseRoadmap] = await Promise.all([
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/editor-quality-roadmap.md'),
      'utf8',
    ),
  ]);
  expect(roadmap).toContain('all five native `w:em` emphasis values');
  expect(product).toContain('The forty-first milestone');

  for (const version of ['latest', '0.28.0', '0.27.0']) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const component = await readFile(
        path.join(documentationRoot, version, lang, 'components/document.mdx'),
        'utf8',
      );
      for (const mark of ['none', 'dot', 'comma', 'circle', 'underDot']) {
        expect(component).toContain(`\`${mark}\``);
      }
      expect(component).toContain('w:em');
      expect(component).toContain('Cmd/Ctrl+D');
      expect(component).toContain('data-office-emphasis-mark');
      expect(component).toContain('text-emphasis-style');
      expect(component).toContain('Worker/WASM');
      expect(component).toContain('w:rPrChange');
    }
  }
  for (const roadmapSource of [englishRoadmap, chineseRoadmap]) {
    expect(roadmapSource).toContain('w:em');
    expect(roadmapSource).toContain('underDot');
    expect(roadmapSource).toContain('Worker/WASM');
  }
});

test('documents native Writer hidden text in both current locales', async () => {
  const [roadmap, product, englishRoadmap, chineseRoadmap, architecture] =
    await Promise.all([
      readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
      readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
      readFile(
        path.join(documentationRoot, 'latest/en/editor-quality-roadmap.md'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, 'latest/zh/editor-quality-roadmap.md'),
        'utf8',
      ),
      readFile(
        path.join(
          documentationRoot,
          'latest/en/browser-editor-architecture.md',
        ),
        'utf8',
      ),
    ]);

  expect(roadmap).toContain('native `w:vanish` hidden text');
  expect(product).toContain('The forty-second milestone');
  expect(architecture).toContain(
    'Native hidden text is resolved before layout',
  );

  for (const version of ['latest', '0.28.0', '0.27.0']) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const component = await readFile(
        path.join(documentationRoot, version, lang, 'components/document.mdx'),
        'utf8',
      );
      for (const evidence of [
        'w:vanish',
        'Cmd/Ctrl+D',
        'Cmd/Ctrl+Shift+H',
        'data-office-hidden-text',
        'Worker/WASM',
        'w:rPrChange',
      ]) {
        expect(component).toContain(evidence);
      }
    }
  }
  for (const roadmapSource of [englishRoadmap, chineseRoadmap]) {
    expect(roadmapSource).toContain('w:vanish');
    expect(roadmapSource).toContain('Cmd/Ctrl+Shift+H');
    expect(roadmapSource).toContain('Worker/WASM');
  }
});

test('documents native Writer outline, shadow, emboss, and imprint effects', async () => {
  const [roadmap, product, englishRoadmap, chineseRoadmap, architecture] =
    await Promise.all([
      readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
      readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
      readFile(
        path.join(documentationRoot, 'latest/en/editor-quality-roadmap.md'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, 'latest/zh/editor-quality-roadmap.md'),
        'utf8',
      ),
      readFile(
        path.join(
          documentationRoot,
          'latest/en/browser-editor-architecture.md',
        ),
        'utf8',
      ),
    ]);

  expect(roadmap).toContain('independent native `w:outline`');
  expect(product).toContain('The forty-third milestone');
  expect(architecture).toContain('collision-safe nested `w:rStyle` markers');

  for (const version of ['latest', '0.28.0']) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const component = await readFile(
        path.join(documentationRoot, version, lang, 'components/document.mdx'),
        'utf8',
      );
      for (const evidence of [
        'w:outline',
        'w:shadow',
        'w:emboss',
        'w:imprint',
        'Cmd/Ctrl+D',
        'data-office-legacy-text-outline',
        'Worker/WASM',
        'w:rPrChange',
      ]) {
        expect(component).toContain(evidence);
      }
    }
  }
  for (const roadmapSource of [englishRoadmap, chineseRoadmap]) {
    expect(roadmapSource).toContain('w:outline');
    expect(roadmapSource).toContain('w:imprint');
    expect(roadmapSource).toContain('Worker/WASM');
  }
});

test('documents native Writer character borders in both current locales', async () => {
  const [roadmap, product, changelog, architecture] = await Promise.all([
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/browser-editor-architecture.md'),
      'utf8',
    ),
  ]);

  expect(roadmap).toContain('native `w:bdr` character borders');
  expect(product).toContain('The forty-fourth milestone');
  expect(changelog).toContain('## 0.29.0 - 2026-08-24');
  expect(architecture).toContain(
    'Native character borders are resolved as one validated TextStyle value',
  );

  for (const version of ['latest', '0.29.0']) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const component = await readFile(
        path.join(documentationRoot, version, lang, 'components/document.mdx'),
        'utf8',
      );
      for (const evidence of [
        'w:bdr',
        '`nil`',
        '`none`',
        'Cmd/Ctrl+D',
        'data-office-run-border',
        'Worker/WASM',
        'w:rPrChange',
      ]) {
        expect(component).toContain(evidence);
      }
    }
  }
});

test('documents collaborative paragraph-formatting revisions', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
    '0.13.0',
    '0.12.0',
    '0.11.0',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [document, collaboration, server, architecture, roadmap] =
        await Promise.all([
          readFile(path.join(localeRoot, 'components/document.mdx'), 'utf8'),
          readFile(
            path.join(localeRoot, 'components/collaboration.mdx'),
            'utf8',
          ),
          readFile(
            path.join(localeRoot, 'components/collaboration-server.mdx'),
            'utf8',
          ),
          readFile(
            path.join(localeRoot, 'browser-editor-architecture.md'),
            'utf8',
          ),
          readFile(path.join(localeRoot, 'editor-quality-roadmap.md'), 'utf8'),
        ]);

      for (const source of [
        document,
        collaboration,
        server,
        architecture,
        roadmap,
      ]) {
        expect(source).toContain('paragraph-formatting');
        expect(source).toContain('w:pPrChange');
      }
      expect(collaboration).toContain('changeKind: "paragraph-formatting"');
      expect(collaboration).toContain('word-paragraph-formatting-revision.acl');
      expect(server).toContain('A3S Boot');
    }
  }
});

test('documents atomic native Spreadsheet cell batches', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
    '0.13.0',
    '0.12.0',
    '0.11.0',
    '0.10.0',
    '0.9.2',
    '0.9.1',
    '0.9.0',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [collaboration, spreadsheet, cli] = await Promise.all([
        readFile(path.join(localeRoot, 'components/collaboration.mdx'), 'utf8'),
        readFile(path.join(localeRoot, 'components/spreadsheet.mdx'), 'utf8'),
        readFile(path.join(localeRoot, 'cli-reference.md'), 'utf8'),
      ]);

      for (const source of [collaboration, spreadsheet, cli]) {
        expect(source).toContain('spreadsheet-batch-cells');
        expect(source).toContain('nextCell: null');
      }
    }
  }
});

test('documents maximum sparse spreadsheets and cancellable imports in 0.12.0', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
    '0.13.0',
    '0.12.0',
  ]) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const localeRoot = path.join(documentationRoot, version, lang);
      const [guide, spreadsheet] = await Promise.all([
        readFile(path.join(localeRoot, 'guide/index.mdx'), 'utf8'),
        readFile(path.join(localeRoot, 'components/spreadsheet.mdx'), 'utf8'),
      ]);

      for (const value of [
        '1,048,576',
        '16,384',
        'dataValidationRanges',
        'cellProtectionRanges',
      ]) {
        expect(spreadsheet).toContain(value);
      }
      for (const value of ['AbortSignal', 'reading', 'finalizing']) {
        expect(guide).toContain(value);
      }
    }
  }
});

test('documents Traditional Office four-direction cell fill in 0.13.0', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
    '0.13.0',
  ]) {
    const spreadsheet = await readFile(
      path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
      'utf8',
    );

    for (const evidence of [
      'Fill Down',
      'Fill Right',
      'Fill Up',
      'Fill Left',
      'Cmd/Ctrl+D',
      'Cmd/Ctrl+R',
      '50,000 target cells',
    ]) {
      expect(spreadsheet).toContain(evidence);
    }
  }
});

test('documents Traditional Office number formats and cell styles in 0.13.1', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
  ]) {
    const spreadsheet = await readFile(
      path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
      'utf8',
    );

    for (const evidence of [
      'CNY Currency',
      'Scientific',
      'Fraction',
      '17 Office-familiar built-in choices',
      '10,000 cells',
      'theme, indexed, automatic, and valid tint colors',
    ]) {
      expect(spreadsheet).toContain(evidence);
    }
  }
});

test('documents Traditional Office Paste Special and its bounded rich clipboard in 0.14.0', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
  ]) {
    const spreadsheet = await readFile(
      path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
      'utf8',
    );

    for (const evidence of [
      'Paste Special and the rich clipboard',
      'Cmd/Ctrl+Alt+V',
      'All except borders',
      'Values and number formats',
      'Column widths',
      '50,000 destination cells',
      'one Undo record',
      'spreadsheet-paste-special.acl',
    ]) {
      expect(spreadsheet).toContain(evidence);
    }
  }
});

test('documents Traditional Office hyperlinks and immutable workbook updates in 0.15.0', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
  ]) {
    const spreadsheet = await readFile(
      path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
      'utf8',
    );

    for (const evidence of [
      '## Hyperlinks',
      'Cmd/Ctrl+K',
      'Web page',
      'Cell range',
      'existing dense `data` or sparse',
      'one Undo record',
      'spreadsheet-hyperlink.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(spreadsheet).toContain(evidence);
    }
  }
});

test('documents Traditional Office data validation and native XLSX semantics in 0.15.0', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
  ]) {
    const spreadsheet = await readFile(
      path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
      'utf8',
    );

    for (const evidence of [
      '## Data validation',
      'Whole number',
      'Text length',
      '1900 or 1904 date system',
      '10,000 cells',
      'dataValidationRanges',
      'spreadsheet-data-validation.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(spreadsheet).toContain(evidence);
    }
  }
});

test('documents native Spreadsheet Tables, collaboration, and performance boundaries', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
  ]) {
    const [spreadsheet, collaboration, architecture, chineseSpreadsheet] =
      await Promise.all([
        readFile(
          path.join(
            documentationRoot,
            version,
            'en/components/spreadsheet.mdx',
          ),
          'utf8',
        ),
        readFile(
          path.join(
            documentationRoot,
            version,
            'en/components/collaboration.mdx',
          ),
          'utf8',
        ),
        readFile(
          path.join(
            documentationRoot,
            version,
            'en/browser-editor-architecture.md',
          ),
          'utf8',
        ),
        readFile(
          path.join(
            documentationRoot,
            version,
            'zh/components/spreadsheet.mdx',
          ),
          'utf8',
        ),
      ]);

    for (const evidence of [
      '## Native Tables and ListObjects',
      'Cmd/Ctrl+T',
      'Light 1–21',
      'Medium 1–28',
      'Dark 1–11',
      '100,000 cells',
      'Convert to Range',
      'spreadsheet-table.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(spreadsheet).toContain(evidence);
    }
    for (const evidence of ['tableOrder', 'creation claim', 'two-client']) {
      expect(collaboration).toContain(evidence);
    }
    for (const evidence of [
      'content-visibility: auto',
      'semantic `WorkSpreadsheetTable`/OOXML ListObject benchmark',
    ]) {
      expect(architecture).toContain(evidence);
    }
    for (const evidence of [
      '## 原生 Table 与 ListObject',
      '0.480 秒',
      '120.1 FPS',
      '12.9 ms',
    ]) {
      expect(chineseSpreadsheet).toContain(evidence);
    }
  }
});

test('documents Traditional Office font-size and border shortcuts in 0.17.0', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
  ]) {
    const [english, chinese] = await Promise.all([
      readFile(
        path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, version, 'zh/components/spreadsheet.mdx'),
        'utf8',
      ),
    ]);

    for (const evidence of [
      '## Font-size and border shortcuts',
      'Cmd/Ctrl+Shift+.',
      'Cmd/Ctrl+]',
      'Cmd/Ctrl+Shift+&',
      'Cmd/Ctrl+Shift+_',
      '`batchCallApis`',
      '10,000 cells',
      '`aria-keyshortcuts`',
      'spreadsheet-font-size-border-shortcuts.acl',
    ]) {
      expect(english).toContain(evidence);
    }
    for (const evidence of [
      '## 字号步进与框线快捷键',
      '增大字号',
      '减小字号',
      '外侧框线',
      '清除框线',
      '10,000 个',
    ]) {
      expect(chinese).toContain(evidence);
    }
  }
});

test('documents exact advanced Spreadsheet underline styles in 0.18.0', async () => {
  for (const version of ['latest', '0.21.0', '0.20.0', '0.19.0', '0.18.0']) {
    const [english, chinese] = await Promise.all([
      readFile(
        path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, version, 'zh/components/spreadsheet.mdx'),
        'utf8',
      ),
    ]);

    for (const evidence of [
      '## Advanced underline styles',
      'Single accounting',
      'Double accounting',
      '`Cmd/Ctrl+U`',
      '`menuitemradio`',
      'val="singleAccounting"',
      'val="doubleAccounting"',
      'spreadsheet-underline-styles.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(english).toContain(evidence);
    }
    for (const evidence of [
      '## 高级下划线样式',
      '单会计用下划线',
      '双会计用下划线',
      '`Cmd/Ctrl+U`',
      'val="singleAccounting"',
      'val="doubleAccounting"',
      'spreadsheet-underline-styles.acl',
    ]) {
      expect(chinese).toContain(evidence);
    }
  }
});

test('documents text orientation and bounded row/column visibility in 0.19.0', async () => {
  for (const version of ['latest', '0.21.0', '0.20.0', '0.19.0']) {
    const [english, chinese] = await Promise.all([
      readFile(
        path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, version, 'zh/components/spreadsheet.mdx'),
        'utf8',
      ),
    ]);

    for (const evidence of [
      '## Text orientation and row/column visibility',
      'Angle Counterclockwise',
      'Angle Clockwise',
      "`tr='3'`",
      'XLSX `textRotation`',
      '`255`',
      '`Cmd/Ctrl+Shift+9`',
      '`Cmd/Ctrl+Shift+0`',
      '10,000 rows',
      '1,000 columns',
      'spreadsheet-ribbon-orientation-visibility.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(english).toContain(evidence);
    }
    for (const evidence of [
      '## 文字方向与行列显隐',
      '逆时针倾斜',
      '顺时针倾斜',
      "`tr='3'`",
      '`255`',
      '`Cmd/Ctrl+Shift+9`',
      '`Cmd/Ctrl+Shift+0`',
      '10,000 行',
      '1,000 列',
      'spreadsheet-ribbon-orientation-visibility.acl',
    ]) {
      expect(chinese).toContain(evidence);
    }
  }
});

test('documents direct color resets, Traditional Office font aliases, and XLSX color identity in 0.20.0', async () => {
  for (const version of ['latest', '0.21.0', '0.20.0']) {
    const [english, chinese] = await Promise.all([
      readFile(
        path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, version, 'zh/components/spreadsheet.mdx'),
        'utf8',
      ),
    ]);

    for (const evidence of [
      '## Direct color resets, Traditional Office font aliases, and XLSX color identity',
      'Automatic Color',
      'No Fill',
      '`Ctrl+2`',
      '`Ctrl+3`',
      '`Ctrl+4`',
      'theme="4"',
      'indexed="0"',
      'palette slot',
      'spreadsheet-font-colors-shortcuts.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(english).toContain(evidence);
    }
    for (const evidence of [
      '## 直接颜色重置、传统 Office 字体快捷键与 XLSX 颜色身份',
      '自动颜色',
      '无填充',
      '`Ctrl+2`',
      '`Ctrl+3`',
      '`Ctrl+4`',
      'theme="4"',
      'indexed="0"',
      '调色板槽位',
      'spreadsheet-font-colors-shortcuts.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(chinese).toContain(evidence);
    }
  }
});

test('documents independent Spreadsheet diagonal borders in 0.21.0', async () => {
  for (const version of ['latest', '0.21.0']) {
    const [english, chinese] = await Promise.all([
      readFile(
        path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, version, 'zh/components/spreadsheet.mdx'),
        'utf8',
      ),
    ]);

    for (const evidence of [
      '## Cell borders',
      'diagonal-down',
      'diagonal-up',
      'crossed border',
      '`diagonalDown`',
      '`diagonalUp`',
      '4,096',
      'spreadsheet-diagonal-borders.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(english).toContain(evidence);
    }
    for (const evidence of [
      '## 单元格框线',
      '斜下',
      '斜上',
      '交叉框线',
      '`diagonalDown`',
      '`diagonalUp`',
      '4,096',
      'spreadsheet-diagonal-borders.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(chinese).toContain(evidence);
    }
  }
});

test('documents static Spreadsheet date and time entry in 0.22.0', async () => {
  for (const version of ['latest', '0.22.0']) {
    const [english, chinese] = await Promise.all([
      readFile(
        path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, version, 'zh/components/spreadsheet.mdx'),
        'utf8',
      ),
    ]);

    for (const evidence of [
      '### Static current date and time',
      '`Ctrl+;`',
      '`Ctrl+Shift+;`',
      '`yyyy-MM-dd`',
      '`hh:mm`',
      'Only the active cell',
      'one Undo',
      'spreadsheet-date-time.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(english).toContain(evidence);
    }
    for (const evidence of [
      '### 静态当前日期与时间',
      '`Ctrl+;`',
      '`Ctrl+Shift+;`',
      '`yyyy-MM-dd`',
      '`hh:mm`',
      '活动单元格',
      '一次撤销',
      'spreadsheet-date-time.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(chinese).toContain(evidence);
    }
  }
});

test('documents exact Spreadsheet copy from above in 0.23.0', async () => {
  for (const version of ['latest', '0.23.0']) {
    const [english, chinese] = await Promise.all([
      readFile(
        path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, version, 'zh/components/spreadsheet.mdx'),
        'utf8',
      ),
    ]);

    for (const evidence of [
      '### Copy a formula or value from above',
      "`Ctrl+'`",
      "`Ctrl+Shift+'`",
      'without translating relative references',
      'target keeps its own',
      'one Undo',
      'spreadsheet-copy-from-above.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(english).toContain(evidence);
    }
    for (const evidence of [
      '### 从上方复制公式或值',
      "`Ctrl+'`",
      "`Ctrl+Shift+'`",
      '不平移相对引用',
      '目标自己的',
      '一次撤销',
      'spreadsheet-copy-from-above.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(chinese).toContain(evidence);
    }
  }
});

test('documents focused Spreadsheet font-dialog shortcuts in 0.24.0', async () => {
  for (const version of ['latest', '0.24.0']) {
    const [english, chinese] = await Promise.all([
      readFile(
        path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, version, 'zh/components/spreadsheet.mdx'),
        'utf8',
      ),
    ]);

    for (const evidence of [
      '## Format Cells',
      '`Cmd/Ctrl+Shift+F`',
      '`Cmd/Ctrl+Shift+P`',
      'font-family',
      'control focused',
      'relative luminance',
      'spreadsheet-font-dialog-shortcuts.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(english).toContain(evidence);
    }
    for (const evidence of [
      '## 设置单元格格式与字体快捷键',
      '`Cmd/Ctrl+Shift+F`',
      '`Cmd/Ctrl+Shift+P`',
      '字体下拉框',
      '相对亮度',
      'spreadsheet-font-dialog-shortcuts.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(chinese).toContain(evidence);
    }
  }
});

test('documents native Spreadsheet rich-text authoring through formatted paste', async () => {
  for (const version of ['latest', '0.28.0', '0.27.0', '0.26.0', '0.25.0']) {
    const [english, chinese] = await Promise.all([
      readFile(
        path.join(documentationRoot, version, 'en/components/spreadsheet.mdx'),
        'utf8',
      ),
      readFile(
        path.join(documentationRoot, version, 'zh/components/spreadsheet.mdx'),
        'utf8',
      ),
    ]);

    for (const evidence of [
      '## Native XLSX rich-text cells',
      "`ct.t='inlineStr'`",
      '`xml:space="preserve"`',
      '32,767 characters',
      '512 runs',
      '100,000 runs',
      'spreadsheet-rich-text.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(english).toContain(evidence);
    }
    if (version !== '0.25.0') {
      for (const evidence of [
        'non-collapsed text selection',
        'Home ribbon applies font',
        'bold, italic, underline, or strikethrough',
        'invalid UTF-16 surrogate boundaries',
        'formula bar or the F2',
        'one controlled revision',
        'one Undo record',
        'Formatted rich-text paste',
      ]) {
        expect(english).toContain(evidence);
      }
    } else {
      expect(english).toContain('partial-run authoring');
    }
    for (const evidence of [
      '## 原生 XLSX 单元格富文本',
      "`ct.t='inlineStr'`",
      '`xml:space="preserve"`',
      '32,767 个字符',
      '512 个文字片段',
      '100,000 个片段',
      'spreadsheet-rich-text.acl',
      'A3S Test 1.0.0',
    ]) {
      expect(chinese).toContain(evidence);
    }
    if (version !== '0.25.0') {
      for (const evidence of [
        '选中非空文字',
        '字体、字号、颜色、粗体、斜体、下划线与删除线',
        'UTF-16 代理对',
        '公式栏或 F2',
        '一次受控修订',
        '一条撤销记录',
        '带格式富文本粘贴',
      ]) {
        expect(chinese).toContain(evidence);
      }
    } else {
      expect(chinese).toContain('局部片段创作');
    }
  }
});

test('publishes slice 33 rich-text reconciliation architecture', async () => {
  const [englishRoadmap, chineseRoadmap, architecture] = await Promise.all([
    readFile(
      path.join(documentationRoot, 'latest/en/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/browser-editor-architecture.md'),
      'utf8',
    ),
  ]);

  expect(englishRoadmap).toContain('The thirty-third slice');
  expect(englishRoadmap).toMatch(/text-stable focus\s+callbacks/);
  expect(chineseRoadmap).toContain('第三十三个 Spreadsheet 纵向切片');
  expect(chineseRoadmap).toMatch(/文字未变化的聚焦\s*回调/);
  expect(architecture).toContain(
    'Native XLSX rich-text editing is reconciled at the controlled Fortune boundary',
  );
  expect(architecture).toContain(
    'Authenticated `data[row][column]` operations',
  );
});

test('publishes reproducible 100k Document performance evidence', async () => {
  for (const version of [
    'latest',
    '0.21.0',
    '0.20.0',
    '0.19.0',
    '0.18.0',
    '0.17.0',
    '0.16.0',
    '0.15.0',
    '0.14.0',
    '0.13.1',
    '0.13.0',
  ]) {
    const [document, architecture] = await Promise.all([
      readFile(
        path.join(documentationRoot, version, 'en/components/document.mdx'),
        'utf8',
      ),
      readFile(
        path.join(
          documentationRoot,
          version,
          'en/browser-editor-architecture.md',
        ),
        'utf8',
      ),
    ]);

    for (const evidence of [
      '100,000 text paragraphs',
      '100,000 table rows',
      '0.397 s',
      '120.0 FPS',
      '70.4 MiB',
      '53.9 ms',
    ]) {
      expect(document).toContain(evidence);
      expect(architecture).toContain(evidence);
    }
    for (const command of [
      'bun run performance:large-documents',
      'bun run performance:large-document-edits',
      'bun run test:e2e:large-documents',
    ]) {
      expect(document).toContain(command);
    }
    expect(document).toContain(
      '../browser-editor-architecture.md#current-100000-unit-evidence',
    );
  }
});

test('publishes bounded Spreadsheet structured references across code, docs, and Playground', async () => {
  const [
    roadmap,
    changelog,
    product,
    english,
    chinese,
    architecture,
    chineseArchitecture,
    quality,
    chineseQuality,
    releaseEnglish,
    releaseChinese,
    templates,
    workspaceHome,
    discoverability,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/browser-editor-architecture.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/browser-editor-architecture.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.37.0/en/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.37.0/zh/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'src/internal/features/work/work-templates.ts'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'playground/src/workspace-home.tsx'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'tests/e2e/playground-template-discoverability.acl',
      ),
      'utf8',
    ),
  ]);

  expect(roadmap).toContain('bounded structured-reference calculation');
  expect(roadmap).toContain('Automatic calculated-column fill');
  expect(roadmap).toContain('now\n  supported');
  expect(changelog).toContain(
    'bounded Spreadsheet structured-reference calculation',
  );
  expect(changelog).toContain('automatic Spreadsheet calculated-column fill');
  expect(product).toContain('follow-up structured-reference slice');
  expect(product).toContain('fills only newly');
  expect(product).toContain('inserted');
  expect(product).toContain('body rows');
  expect(english).toContain('## Structured-reference calculation');
  expect(chinese).toContain('## 结构化引用计算');
  expect(english).toContain('=SUM(Sales[Revenue])');
  expect(chinese).toContain('=SUM(Sales[Revenue])');
  expect(english).toContain('<calculatedColumnFormula>');
  expect(chinese).toContain('calculatedColumnFormula');
  expect(english).toContain('Rows inserted outside the table');
  expect(english).toContain('do not trigger a fill');
  expect(chinese).toContain('表格外插入的行不会触发填充');
  expect(architecture).toContain('worksheet-qualified tables');
  expect(chineseArchitecture).toContain('工作表限定');
  expect(architecture).toContain('newly inserted body rows');
  expect(chineseArchitecture).toContain('新插入的正文空单元格');
  expect(quality).toContain('The follow-up structured-reference slice');
  expect(chineseQuality).toContain('结构化引用计算切片');
  expect(quality).toContain('conflicts remove the rule');
  expect(chineseQuality).toContain('检测到冲突时会删除规则');
  expect(releaseEnglish).toContain('calculated-column rule');
  expect(releaseEnglish).toContain('<calculatedColumnFormula>');
  expect(releaseChinese).toContain('计算列规则');
  expect(releaseChinese).toContain('calculatedColumnFormula');
  expect(templates).toContain("id: 'structured-references'");
  expect(templates).toContain('=[@Units]*[@[Unit price]]');
  expect(templates).toContain('=SUM(Sales[Revenue])');
  expect(templates).toContain('插入表格正文行会自动补齐 Revenue');
  expect(workspaceHome).toContain('data-template-id={template.id}');
  expect(workspaceHome).not.toContain('data-release');
  expect(discoverability).toContain(
    "button[data-template-id='structured-references']",
  );
});

test('publishes formula-safe Spreadsheet sorting, owned-range reconciliation, and bounded custom-list preferences without rewriting 0.37.0', async () => {
  const [
    roadmap,
    changelog,
    product,
    english,
    chinese,
    architecture,
    chineseArchitecture,
    quality,
    chineseQuality,
    releaseEnglish,
    releaseChinese,
    sortCommand,
    structureCommand,
    e2e,
    rangeE2e,
    ownedRangeE2e,
    ownedRangeVisual,
    filterReconciliation,
    customListModel,
    customListStore,
    core,
    vueAdapter,
    webComponent,
    customListDialog,
    customListManager,
    sortOrderControls,
    customListE2e,
    customListVisual,
    appearanceModel,
    appearanceE2e,
    appearanceVisual,
    sortMatrix,
    sortOptionsDialog,
    collationModel,
    rowSortE2e,
    rowSortVisual,
    textSortE2e,
    textSortVisual,
    focusTest,
    currentRegion,
    rangeDialog,
    packageJson,
    gate,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/browser-editor-architecture.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/browser-editor-architecture.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.37.0/en/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, '0.37.0/zh/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-sort-command.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-structure-command.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/spreadsheet-custom-sort.acl'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/spreadsheet-sort-range.acl'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/spreadsheet-owned-range-sort.acl'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'visual-tests/spreadsheet-owned-range-sort.functional.spec.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-filter-reconciliation.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-sort-custom-list.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-sort-custom-list-store.ts',
      ),
      'utf8',
    ),
    readFile(path.join(repositoryRoot, 'src/core.ts'), 'utf8'),
    readFile(path.join(repositoryRoot, 'src/vue.ts'), 'utf8'),
    readFile(path.join(repositoryRoot, 'src/web-component.tsx'), 'utf8'),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-sort-dialog.tsx',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-sort-custom-list-manager.tsx',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-sort-order-controls.tsx',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/spreadsheet-custom-list-sort.acl'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'visual-tests/spreadsheet-custom-list-sort.functional.spec.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-sort-appearance.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/spreadsheet-appearance-sort.acl'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'visual-tests/spreadsheet-appearance-sort.functional.spec.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-sort-matrix.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-sort-options-dialog.tsx',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-sort-collation.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/spreadsheet-row-sort.acl'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'visual-tests/spreadsheet-row-sort.functional.spec.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/spreadsheet-text-sort.acl'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'visual-tests/spreadsheet-text-sort.functional.spec.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/spreadsheet-editor-focus.test.ts'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-current-region.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-sort-range-dialog.tsx',
      ),
      'utf8',
    ),
    readFile(path.join(repositoryRoot, 'package.json'), 'utf8'),
    readFile(
      path.join(repositoryRoot, 'scripts/run-a3s-test-web-gate.sh'),
      'utf8',
    ),
  ]);

  expect(roadmap).toContain(
    'stable value/custom-list/effective-color/conditional-icon sorting',
  );
  expect(roadmap).toContain('exact-or-expand warnings');
  expect(roadmap).toContain('bounded host-stored/session sequences');
  expect(roadmap).toContain(
    'accessible preference manager for atomic user-sequence creation',
  );
  expect(roadmap).not.toContain('Custom-list deletion/reordering UI');
  expect(roadmap).toContain('effective-color/conditional-icon sorting');
  expect(roadmap).toContain('left-to-right complete columns');
  expect(roadmap).toContain('pinyin/stroke comparison');
  expect(roadmap).toContain('exact table/AutoFilter owner expansion');
  expect(roadmap).not.toContain(
    'table/AutoFilter-integrated custom sorting remain incomplete',
  );
  expect(changelog).toContain('Data → Custom Sort');
  expect(changelog).toContain('coordinate-owned hyperlink');
  expect(changelog).toContain('Sort Warning');
  expect(changelog).toContain('effective cell color');
  expect(changelog).toContain('Sort Options');
  expect(changelog).toContain('lowercase before uppercase');
  expect(changelog).toContain('LocalStorageSpreadsheetSortCustomListStore');
  expect(changelog).toContain(
    'Spreadsheet **Custom Lists** preference manager',
  );
  expect(changelog).toContain(
    'native table and worksheet\n  AutoFilter ownership',
  );
  expect(product).toContain('The forty-eighth milestone');
  expect(product).toContain('The forty-ninth milestone');
  expect(product).toContain('The fiftieth milestone');
  expect(product).toContain('The fifty-first milestone');
  expect(product).toContain('The fifty-second milestone');
  expect(product).toContain('The fifty-third milestone');
  expect(product).toContain('The fifty-fourth milestone');
  expect(product).toContain('The fifty-ninth milestone');
  expect(product).toContain('responsive, keyboard-accessible manager');
  expect(english).toContain('## Multi-key custom sort');
  expect(chinese).toContain('## 多关键字自定义排序');
  expect(english).toContain('coordinate-owned');
  expect(chinese).toContain('坐标归属');
  expect(english).toContain('Traditional Office-style Sort Warning');
  expect(chinese).toContain('传统 Office 风格“排序提醒”');
  expect(english).toContain('seven built-in Chinese or English month');
  expect(chinese).toContain('七个中英文月份/星期内置序列');
  expect(english).toContain('effective cell color');
  expect(chinese).toContain('有效单元格颜色');
  expect(english).toContain('Sort left to right');
  expect(chinese).toContain('按行排序');
  expect(english).toContain('**Pinyin** or **Stroke**');
  expect(chinese).toContain('“拼音排序”或“笔画排序”');
  expect(english).toContain('`sortCustomListStore`');
  expect(chinese).toContain('`sortCustomListStore`');
  expect(english).toContain('LocalStorageSpreadsheetSortCustomListStore');
  expect(english).toContain('**Custom Lists** opens a keyboard-accessible');
  expect(chinese).toContain('LocalStorageSpreadsheetSortCustomListStore');
  expect(chinese).toContain('可用键盘操作的偏好管理器');
  expect(english).toContain('exactly one semantic ListObject or');
  expect(english).toContain('opaque `caljs` state');
  expect(chinese).toContain('唯一一张语义 ListObject');
  expect(chinese).toContain('不透明 `caljs` 状态');
  expect(english).toContain(
    "element's `sortCustomListStore` JavaScript property",
  );
  expect(chinese).toContain('元素的 `sortCustomListStore` JavaScript 属性');
  expect(architecture).toContain('`spreadsheetSort` editor extension');
  expect(chineseArchitecture).toContain('`spreadsheetSort` 编辑器扩展');
  expect(architecture).toContain('`spreadsheet-sort-appearance` model');
  expect(chineseArchitecture).toContain('`spreadsheet-sort-appearance` 模型');
  expect(architecture).toContain('direction-neutral matrix engine');
  expect(chineseArchitecture).toContain('方向中立的矩阵引擎');
  expect(architecture).toContain('`spreadsheet-sort-collation` boundary');
  expect(chineseArchitecture).toContain('`spreadsheet-sort-collation` 边界');
  expect(architecture).toContain(
    '`LocalStorageSpreadsheetSortCustomListStore`',
  );
  expect(architecture).toContain('A dedicated manager stages create, edit');
  expect(chineseArchitecture).toContain(
    '`LocalStorageSpreadsheetSortCustomListStore`',
  );
  expect(architecture).toContain('`spreadsheet-current-region` model');
  expect(chineseArchitecture).toContain('`spreadsheet-current-region` 模型');
  expect(architecture).toContain('`spreadsheet-filter-reconciliation`');
  expect(chineseArchitecture).toContain('`spreadsheet-filter-reconciliation`');
  expect(quality).toContain('The thirty-eighth Spreadsheet slice');
  expect(chineseQuality).toContain('第三十八个 Spreadsheet 纵向切片');
  expect(quality).toContain('The thirty-ninth Spreadsheet slice');
  expect(chineseQuality).toContain('第三十九个 Spreadsheet 纵向切片');
  expect(quality).toContain('The fortieth Spreadsheet slice');
  expect(chineseQuality).toContain('第四十个 Spreadsheet 纵向切片');
  expect(quality).toContain('The forty-first Spreadsheet slice');
  expect(chineseQuality).toContain('第四十一个 Spreadsheet 纵向切片');
  expect(quality).toContain('The forty-second Spreadsheet slice');
  expect(chineseQuality).toContain('第四十二个 Spreadsheet 纵向切片');
  expect(quality).toContain('The forty-third Spreadsheet slice');
  expect(chineseQuality).toContain('第四十三个 Spreadsheet 纵向切片');
  expect(quality).toContain('The forty-fourth Spreadsheet slice');
  expect(chineseQuality).toContain('第四十四个 Spreadsheet 纵向切片');
  expect(quality).toContain('The fiftieth Spreadsheet slice');
  expect(chineseQuality).toContain('第五十个 Spreadsheet 纵向切片');
  expect(quality).toContain('The fifty-first Spreadsheet slice');
  expect(chineseQuality).toContain('第五十一个 Spreadsheet 纵向切片');
  expect(releaseEnglish).not.toContain('## Multi-key custom sort');
  expect(releaseChinese).not.toContain('## 多关键字自定义排序');
  expect(releaseEnglish).not.toContain('Traditional Office-style Sort Warning');
  expect(releaseChinese).not.toContain('传统 Office 风格“排序提醒”');
  expect(sortCommand).toContain("name: 'spreadsheetSort'");
  expect(sortCommand).toContain('spreadsheetSortRangeHasStructuralConflict');
  expect(structureCommand).not.toContain('sortSelectedCells');
  expect(currentRegion).toContain('export function spreadsheetCurrentRegion');
  expect(rangeDialog).toContain('扩展选定区域');
  expect(rangeDialog).toContain('以当前选定区域排序');
  expect(e2e).toContain('moved-formula-is-rebased');
  expect(e2e).toContain('original-formula-is-restored-in-one-step');
  expect(e2e).toContain('spreadsheet-after-custom-sort-accessibility');
  expect(rangeE2e).toContain('expand-is-default');
  expect(rangeE2e).toContain('exact-range-has-no-header');
  expect(rangeE2e).toContain('expanded-formula-restored-in-one-step');
  expect(ownedRangeE2e).toContain('single-cell-sort-is-disabled');
  expect(ownedRangeE2e).toContain('row-orientation-disabled');
  expect(ownedRangeVisual).toContain(
    'reapplies AutoFilter criteria after an owned-range sort and undo',
  );
  expect(ownedRangeVisual).toContain('spreadsheet-auto-filter-owned-sort');
  expect(filterReconciliation).toContain(
    'reconcileSpreadsheetFiltersAfterSort',
  );
  expect(filterReconciliation).toContain(
    'spreadsheetFilterWithRemappedOpaqueEntries',
  );
  expect(customListModel).toContain(
    'MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRIES = 256',
  );
  expect(customListModel).toContain('SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS');
  expect(customListStore).toContain(
    'DEFAULT_SPREADSHEET_SORT_CUSTOM_LIST_STORAGE_KEY',
  );
  expect(customListStore).toContain(
    'SPREADSHEET_SORT_CUSTOM_LIST_STORAGE_VERSION = 1',
  );
  expect(customListStore).toContain(
    'normalizeStoredSpreadsheetSortCustomLists',
  );
  expect(core).toContain('LocalStorageSpreadsheetSortCustomListStore');
  expect(vueAdapter).toContain(
    'sortCustomListStore: props.sortCustomListStore',
  );
  expect(webComponent).toContain(
    'get sortCustomListStore(): SpreadsheetSortCustomListStore | undefined',
  );
  expect(sortOrderControls).toContain('新建自定义序列…');
  expect(customListDialog).toContain('onRememberCustomList');
  expect(customListDialog).toContain('管理自定义序列');
  expect(customListManager).toContain('MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS');
  expect(customListManager).toContain('managedCustomListResult');
  expect(customListE2e).toContain('duplicate-custom-list-is-explained');
  expect(customListE2e).toContain(
    'saved-custom-list-is-reusable-before-reload',
  );
  expect(customListE2e).toContain('saved-custom-list-survives-reload');
  expect(customListE2e).toContain('custom-list-manager-accessibility');
  expect(customListE2e).toContain('delete-temporary-list');
  expect(customListVisual).toContain('spreadsheet-custom-list-sort-dialog.png');
  expect(customListVisual).toContain(
    'spreadsheet-custom-list-manager-dialog.png',
  );
  expect(customListVisual).toContain(
    'a3s-office.spreadsheet-sort-custom-lists.v1',
  );
  expect(appearanceModel).toContain('createSpreadsheetSortAppearanceRows');
  expect(appearanceModel).toContain('spreadsheetConditionalFormatStyles');
  expect(appearanceE2e).toContain('put-high-icon-first');
  expect(appearanceE2e).toContain('moved-formula-is-rebased');
  expect(appearanceVisual).toContain('spreadsheet-appearance-sort-');
  expect(sortMatrix).toContain('export function sortSpreadsheetMatrix');
  expect(sortMatrix).toContain('translateSpreadsheetFormula');
  expect(sortOptionsDialog).toContain('按行排序');
  expect(sortOptionsDialog).toContain('区分大小写');
  expect(sortOptionsDialog).toContain('笔画排序');
  expect(collationModel).toContain('createSpreadsheetSortTextComparator');
  expect(collationModel).toContain('numeric: false');
  expect(rowSortE2e).toContain('choose-row-orientation');
  expect(rowSortE2e).toContain('moved-formula-is-rebased');
  expect(rowSortVisual).toContain('spreadsheet-row-sort-');
  expect(textSortE2e).toContain('choose-stroke-order');
  expect(textSortE2e).toContain('lowercase-precedes-uppercase');
  expect(textSortE2e).toContain('inspect-restored-row-without-delay');
  expect(textSortVisual).toContain('spreadsheet-text-sort-');
  expect(focusTest).toContain(
    'preserves directly focused grid overlays across a controlled remount',
  );
  expect(packageJson).toContain('playground:visual:spreadsheet-custom-sort');
  expect(packageJson).toContain('test:e2e:spreadsheet-custom-sort');
  expect(packageJson).toContain(
    'playground:visual:spreadsheet-custom-list-sort',
  );
  expect(packageJson).toContain('test:e2e:spreadsheet-custom-list-sort');
  expect(packageJson).toContain(
    'playground:visual:spreadsheet-appearance-sort',
  );
  expect(packageJson).toContain('test:e2e:spreadsheet-appearance-sort');
  expect(packageJson).toContain('playground:visual:spreadsheet-row-sort');
  expect(packageJson).toContain('test:e2e:spreadsheet-row-sort');
  expect(packageJson).toContain('playground:visual:spreadsheet-text-sort');
  expect(packageJson).toContain('test:e2e:spreadsheet-text-sort');
  expect(packageJson).toContain('playground:visual:spreadsheet-sort-range');
  expect(packageJson).toContain('test:e2e:spreadsheet-sort-range');
  expect(packageJson).toContain(
    'playground:visual:spreadsheet-owned-range-sort',
  );
  expect(packageJson).toContain('test:e2e:spreadsheet-owned-range-sort');
  expect(gate).toContain('tests/e2e/spreadsheet-custom-sort.acl');
  expect(gate).toContain('tests/e2e/spreadsheet-custom-list-sort.acl');
  expect(gate).toContain('tests/e2e/spreadsheet-appearance-sort.acl');
  expect(gate).toContain('tests/e2e/spreadsheet-row-sort.acl');
  expect(gate).toContain('tests/e2e/spreadsheet-text-sort.acl');
  expect(gate).toContain('tests/e2e/spreadsheet-sort-range.acl');
  expect(gate).toContain('tests/e2e/spreadsheet-owned-range-sort.acl');
});

test('publishes the workbook-owned 1900 and 1904 date-system contract', async () => {
  const [
    readme,
    changelog,
    roadmap,
    product,
    englishSpreadsheet,
    chineseSpreadsheet,
    englishArchitecture,
    chineseArchitecture,
    englishQuality,
    chineseQuality,
    englishRelease,
    chineseRelease,
    previousEnglishSpreadsheet,
    workTypes,
    dynamicFilter,
    fileImport,
    fileExport,
    dateTimeCommand,
    collaborationValidation,
    xlsxTest,
    dateTimeTest,
    e2e,
    playground,
    fixtures,
    gate,
    packageJson,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
    readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
    readFile(
      path.join(documentationRoot, 'latest/en/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/browser-editor-architecture.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/browser-editor-architecture.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/en/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(
      path.join(documentationRoot, 'latest/zh/editor-quality-roadmap.md'),
      'utf8',
    ),
    readFile(path.join(documentationRoot, '0.37.5/en/index.mdx'), 'utf8'),
    readFile(path.join(documentationRoot, '0.37.5/zh/index.mdx'), 'utf8'),
    readFile(
      path.join(documentationRoot, '0.37.4/en/components/spreadsheet.mdx'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'src/internal/features/work/work-types.ts'),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/work-spreadsheet-dynamic-filter.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/work-spreadsheet-file-import.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/work-spreadsheet-file-export.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/features/work/editors/spreadsheet-date-time-command.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(
        repositoryRoot,
        'src/internal/collaboration/office-spreadsheet-collaboration-validation.ts',
      ),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/spreadsheet-date-system-xlsx.test.ts'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/spreadsheet-date-time-command.test.ts'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'tests/e2e/spreadsheet-1904-date-system.acl'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'playground/src/editor-workspace.tsx'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'scripts/create-e2e-fixtures.ts'),
      'utf8',
    ),
    readFile(
      path.join(repositoryRoot, 'scripts/run-a3s-test-web-gate.sh'),
      'utf8',
    ),
    readFile(path.join(repositoryRoot, 'package.json'), 'utf8'),
  ]);

  expect(readme).toContain('Version `0.41.0`');
  expect(readme).toContain("workbook's native 1900 or");
  expect(changelog).toContain('## 0.37.5 - 2026-09-01');
  expect(roadmap).toContain(
    'exact source-level 1900/1904 date-system and typed numeric-serial retention',
  );
  expect(product).toContain('The sixtieth Spreadsheet milestone');

  for (const document of [
    englishSpreadsheet,
    englishArchitecture,
    englishQuality,
  ]) {
    expect(document).toContain('workbook');
    expect(document).toContain('1904');
    expect(document).toContain('serial zero');
  }
  for (const document of [
    chineseSpreadsheet,
    chineseArchitecture,
    chineseQuality,
  ]) {
    expect(document).toContain('工作簿');
    expect(document).toContain('1904');
    expect(document).toContain('序列 0');
  }
  expect(englishSpreadsheet).toContain('## XLSX 1904 date-system retention');
  expect(chineseSpreadsheet).toContain('## XLSX 1904 日期系统保留');
  expect(englishRelease).toContain('# A3S Office 0.37.5 documentation');
  expect(englishRelease).toContain('frozen documentation index');
  expect(chineseRelease).toContain('# A3S Office 0.37.5 文档');
  expect(chineseRelease).toContain('冻结文档索引');
  expect(previousEnglishSpreadsheet).toContain('Source-level 1904');
  expect(previousEnglishSpreadsheet).toContain(
    'date-system retention and large aggregate/rank Worker/WASM offload remain',
  );

  expect(workTypes).toContain(
    "export type WorkSpreadsheetDateSystem = '1900' | '1904';",
  );
  expect(workTypes).toContain('dateSystem?: WorkSpreadsheetDateSystem;');
  expect(dynamicFilter).toContain('EXCEL_1904_EPOCH_UTC');
  expect(fileImport).toContain('cellDates: false');
  expect(fileImport).toContain("? '1904' : '1900'");
  expect(fileExport).toContain("date1904: content.dateSystem === '1904'");
  expect(dateTimeCommand).toContain('context.content.dateSystem');
  expect(collaborationValidation).toContain('content.dateSystem');
  expect(xlsxTest).toContain('retains a 1904 workbook');
  expect(xlsxTest).toContain('expect(raw.Sheets.Dates?.A2?.v).toBe(0)');
  expect(dateTimeTest).toContain('workbook-owned 1904 epoch');
  expect(e2e).toContain('imported-1904-filter-state');
  expect(e2e).toContain('reopened-1904-filter-state');
  expect(playground).toContain('data-february-hidden');
  expect(fixtures).toContain('createSpreadsheet1904DateSystemFixture');
  expect(gate).toContain('tests/e2e/spreadsheet-1904-date-system.acl');
  expect(packageJson).toContain('test:e2e:spreadsheet-1904-date-system');
});

test('keeps the 0.8.0 native suggestion limitation frozen', async () => {
  for (const { lang } of DOCUMENTATION_LOCALES) {
    const cli = await readFile(
      path.join(documentationRoot, '0.8.0', lang, 'cli-reference.md'),
      'utf8',
    );

    expect(cli).toContain('document-suggestion-*');
    expect(cli).toContain('document-change-decision-*');
  }
});
