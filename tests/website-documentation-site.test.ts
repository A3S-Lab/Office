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

function homepageComponentHref(version: string, component: string): string {
  const extension =
    version === 'latest' ||
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

test('keeps the five-editor README capability comparison complete', async () => {
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

  expect(
    readme.match(
      /\| Capability \| A3S Office today \| Traditional Office baseline \|/g,
    ),
  ).toHaveLength(5);
});

test('makes the latest main capabilities discoverable from README and both documentation homes', async () => {
  const [readme, englishHome, chineseHome] = await Promise.all([
    readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
    readFile(path.join(documentationRoot, 'latest/en/index.mdx'), 'utf8'),
    readFile(path.join(documentationRoot, 'latest/zh/index.mdx'), 'utf8'),
  ]);

  expect(readme).toContain('## Latest on `main`');
  expect(readme).toContain('Latest capabilities → 字符底纹');
  expect(readme).toContain('Latest capabilities → 校对语言');
  expect(readme).toContain('Latest capabilities → 数据验证');

  expect(englishHome).toContain('aria-label="Latest capabilities on main"');
  expect(englishHome).toContain('Character shading');
  expect(englishHome).toContain('Proofing languages');
  expect(englishHome).toContain('Data validation');

  expect(chineseHome).toContain('aria-label="main 分支最新能力"');
  expect(chineseHome).toContain('原生字符底纹');
  expect(chineseHome).toContain('原生校对语言');
  expect(chineseHome).toContain('spreadsheet.html#数据验证');
});

test('publishes native Writer proofing languages in README, docs, roadmap, and Playground guidance', async () => {
  const [
    readme,
    changelog,
    roadmap,
    english,
    chinese,
    englishRoadmap,
    chineseRoadmap,
    englishArchitecture,
    chineseArchitecture,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
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

  expect(readme).toContain('Native Writer proofing languages');
  expect(readme).toContain('data-office-proofing-languages');
  expect(readme).toContain("Playground's **校对语言** template");
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

test('publishes complete Spreadsheet data-validation settings in every public surface', async () => {
  const [readme, changelog, roadmap, english, chinese, templates] =
    await Promise.all([
      readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
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
        path.join(
          repositoryRoot,
          'src/internal/features/work/work-templates.ts',
        ),
        'utf8',
      ),
    ]);

  expect(readme).toContain('Complete Spreadsheet data-validation settings');
  expect(readme).toContain('**新建 → 数据验证**');
  expect(readme).toContain('not an `?e2e=` fixture');
  expect(changelog).toContain('input titles and messages');
  expect(changelog).toContain('home-page **新建 → 数据验证**');
  expect(roadmap).toContain('input and error-alert settings');
  expect(english).toContain('### Input and error settings');
  expect(english).toContain('**新建 → 数据验证**');
  expect(chinese).toContain('## 数据验证');
  expect(chinese).toContain('### 输入信息与错误警告设置');
  expect(chinese).toContain('**新建 → 数据验证**');
  expect(templates).toContain("id: 'data-validation'");
  for (const document of [readme, english, chinese]) {
    expect(document).toContain('`allowBlank`');
    expect(document).toContain('`showDropdownArrow`');
    expect(document).toContain('`errorStyle`');
    expect(document).toContain('`hintTitle`');
  }
});

test('documents the complete native Writer underline contract', async () => {
  const [readme, roadmap, product, english, chinese] = await Promise.all([
    readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
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

  expect(readme).toContain('all 18 native DOCX underline values');
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
  const [
    readme,
    roadmap,
    product,
    english,
    chinese,
    englishRoadmap,
    chineseRoadmap,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
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

  expect(readme).toContain('independent native single/double strikethrough');
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

test('keeps published release homepages frozen and visibly versioned', async () => {
  for (const version of [
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

test('removes broken online Playground actions from frozen homepages', async () => {
  for (const version of DOCUMENTATION_VERSIONS.filter(
    (candidate) => candidate !== 'latest',
  )) {
    for (const { lang } of DOCUMENTATION_LOCALES) {
      const homepage = await readFile(
        path.join(documentationRoot, version, lang, 'index.mdx'),
        'utf8',
      );

      expect(homepage).not.toContain('<PlaygroundLink>');
      expect(homepage).not.toContain('在线体验');
      expect(homepage).not.toContain('Open the Playground');
    }
  }
});

test('uses deployable HTML targets in current release homepage actions', async () => {
  for (const version of [
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
      const homepage = await readFile(
        path.join(documentationRoot, version, lang, 'index.mdx'),
        'utf8',
      );

      expect(homepage).not.toMatch(/href="[^"]+\.mdx(?:[?#][^"]*)?"/);
      expect(homepage).toContain('href="./guide/index.html"');
      expect(homepage).toContain('href="./components/collaboration.html"');
      expect(homepage).toContain(
        'href="./components/collaboration-server.html"',
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

      expect(homepage).toContain(
        homepageComponentHref(version, 'collaboration'),
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
      const [homepage, componentIndex, componentNavigation, server] =
        await Promise.all([
          readFile(path.join(localeRoot, 'index.mdx'), 'utf8'),
          readFile(path.join(localeRoot, 'components/index.mdx'), 'utf8'),
          readFile(path.join(localeRoot, 'components/_meta.json'), 'utf8'),
          readFile(
            path.join(localeRoot, 'components/collaboration-server.mdx'),
            'utf8',
          ),
        ]);

      expect(homepage).toContain(
        homepageComponentHref(version, 'collaboration-server'),
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
  const [readme, roadmap, product, englishRoadmap, chineseRoadmap] =
    await Promise.all([
      readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
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
  expect(readme).toContain('all five native East Asian emphasis marks');
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
  const [
    readme,
    roadmap,
    product,
    englishRoadmap,
    chineseRoadmap,
    architecture,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
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
      path.join(documentationRoot, 'latest/en/browser-editor-architecture.md'),
      'utf8',
    ),
  ]);

  expect(readme).toContain('native hidden text with explicit visible resets');
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
  const [
    readme,
    roadmap,
    product,
    englishRoadmap,
    chineseRoadmap,
    architecture,
  ] = await Promise.all([
    readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
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
      path.join(documentationRoot, 'latest/en/browser-editor-architecture.md'),
      'utf8',
    ),
  ]);

  expect(readme).toContain('native outline/shadow/emboss/imprint effects');
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
  const [readme, roadmap, product, changelog, architecture] = await Promise.all(
    [
      readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
      readFile(path.join(repositoryRoot, 'ROADMAP.md'), 'utf8'),
      readFile(path.join(repositoryRoot, 'PRODUCT.md'), 'utf8'),
      readFile(path.join(repositoryRoot, 'CHANGELOG.md'), 'utf8'),
      readFile(
        path.join(
          documentationRoot,
          'latest/en/browser-editor-architecture.md',
        ),
        'utf8',
      ),
    ],
  );

  expect(readme).toContain('25 visible line styles plus `nil` and `none`');
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
  const [englishRoadmap, chineseRoadmap, architecture, readme] =
    await Promise.all([
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
      readFile(path.join(repositoryRoot, 'README.md'), 'utf8'),
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
  expect(readme).toContain('direct formula-bar/F2 insertion or deletion');
  expect(readme).toContain('Traditional Office baseline');
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
