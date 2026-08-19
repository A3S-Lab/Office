import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { expect, test } from '@rstest/core';
import {
  playgroundAssetHrefFromDocsRoute,
  playgroundHrefFromDocsRoute,
} from '../website/theme/site-navigation';

test('derives a deployment-relative Playground link from every docs route depth', () => {
  expect(playgroundHrefFromDocsRoute('/')).toBe('../');
  expect(playgroundHrefFromDocsRoute('/index.html')).toBe('../');
  expect(playgroundHrefFromDocsRoute('/native-office-engine.html')).toBe('../');
  expect(playgroundHrefFromDocsRoute('/guide/')).toBe('../../');
  expect(playgroundHrefFromDocsRoute('/guide/index.html')).toBe('../../');
  expect(playgroundHrefFromDocsRoute('/components/react.html?tab=usage')).toBe(
    '../../',
  );
  expect(playgroundHrefFromDocsRoute('/en/components/react.html')).toBe(
    '../../../',
  );
  expect(playgroundHrefFromDocsRoute('/0.1.0/en/components/react.html')).toBe(
    '../../../../',
  );
});

test('derives deployment-relative Playground assets from localized version routes', () => {
  expect(
    playgroundAssetHrefFromDocsRoute(
      '/0.1.0/en/automation/',
      '/downloads/a3s-office-skill.tar.gz',
    ),
  ).toBe('../../../../downloads/a3s-office-skill.tar.gz');
  expect(playgroundAssetHrefFromDocsRoute('/', '')).toBe('../');
});

test('uses the shared A3S navigation without an online Playground button', async () => {
  const themeRoot = path.resolve(import.meta.dirname, '../website/theme');
  const websiteRoot = path.resolve(import.meta.dirname, '../website');
  const [config, themeEntry, contentStyles, homeStyles, navSource, logo] =
    await Promise.all([
      readFile(path.join(websiteRoot, 'rspress.config.ts'), 'utf8'),
      readFile(path.join(themeRoot, 'index.tsx'), 'utf8'),
      readFile(path.join(themeRoot, 'docs-content.css'), 'utf8'),
      readFile(path.join(themeRoot, 'index.css'), 'utf8'),
      readFile(path.join(themeRoot, 'components/Nav.tsx'), 'utf8'),
      readFile(
        path.resolve(import.meta.dirname, '../docs/public/a3s-logo.png'),
      ),
    ]);
  const themeStyles = `${contentStyles}\n${homeStyles}`;

  expect(config).toContain("logo: '/a3s-logo.png'");
  expect(themeEntry).toContain("export { Nav } from './components/Nav'");
  expect(navSource).toContain('aria-controls="office-mobile-navigation"');
  expect(navSource).toContain(
    "language === 'zh' ? '打开导航' : 'Open navigation'",
  );
  expect(themeStyles).toMatch(/--rp-nav-height:\s*72px/);
  expect(themeStyles).toMatch(/width:\s*31px/);
  expect(createHash('sha256').update(logo).digest('hex')).toBe(
    'ecfcf5c9f783c2c49bf7623cab825a81f500ca7313cd33540d948f276e59e46d',
  );
  expect(themeStyles).not.toContain('.office-docs-playground-link');
});

test('matches the A3S UI documentation rendering contract', async () => {
  const themeRoot = path.resolve(import.meta.dirname, '../website/theme');
  const [themeEntry, themeStyles] = await Promise.all([
    readFile(path.join(themeRoot, 'index.tsx'), 'utf8'),
    Promise.all([
      readFile(path.join(themeRoot, 'docs-content.css'), 'utf8'),
      readFile(path.join(themeRoot, 'index.css'), 'utf8'),
    ]).then((styles) => styles.join('\n')),
  ]);

  expect(themeEntry).toContain("import '@fontsource-variable/geist'");
  expect(themeEntry).toContain("import '@fontsource-variable/geist-mono'");
  expect(themeStyles).toContain(
    'A3S-Lab/UI@e3440e29fa3058fa31a2aacab715833c595fdbc7',
  );
  expect(themeStyles).toMatch(/--rp-content-max-width:\s*920px/);
  expect(themeStyles).toContain('Geist, "Avenir Next"');
  expect(themeStyles).toContain('"Geist Mono"');
  expect(themeStyles).toContain('--rp-code-title-bg');
  expect(themeStyles).toContain('--rp-code-block-bg');
  expect(themeStyles).toContain('.rp-codeblock__content__scroll-container');
  expect(themeStyles).toContain('.rp-doc .rp-code-button-group__button');
  expect(themeStyles).toMatch(
    /\.rp-doc \.rp-code-button-group__button\s*\{[^}]*opacity:\s*1/s,
  );
  expect(themeStyles).toMatch(/:not\(pre\)\s*> code/);
  expect(themeStyles).toContain('.rp-doc .rp-table-scroll-container');
});
