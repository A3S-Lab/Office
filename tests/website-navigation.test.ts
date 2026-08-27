import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { expect, test } from '@rstest/core';
import {
  isProductHomeRoute,
  normalizeNavigationPath,
  productSiteBase,
  playgroundAssetHrefFromDocsRoute,
  playgroundHrefFromDocsRoute,
  siteNavigationHref,
  sitePathFromRoute,
} from '../website/theme/site-navigation';

test('normalizes relative and index navigation links before merging menus', () => {
  expect(normalizeNavigationPath('/', '/index.html')).toBe('/');
  expect(normalizeNavigationPath('/', '../')).toBe('/');
  expect(
    normalizeNavigationPath(
      '/docs/components/document.html',
      '../../../playground/index.html',
    ),
  ).toBe('/playground');
  expect(
    normalizeNavigationPath(
      '/docs/components/document.html',
      'https://github.com/A3S-Lab/Office',
    ),
  ).toBeNull();
});

test('derives a deployment-relative Playground link from every docs route depth', () => {
  expect(playgroundHrefFromDocsRoute('/')).toBe('../playground/');
  expect(playgroundHrefFromDocsRoute('/index.html')).toBe('../playground/');
  expect(playgroundHrefFromDocsRoute('/native-office-engine.html')).toBe(
    '../playground/',
  );
  expect(playgroundHrefFromDocsRoute('/guide/')).toBe('../../playground/');
  expect(playgroundHrefFromDocsRoute('/guide/index.html')).toBe(
    '../../playground/',
  );
  expect(playgroundHrefFromDocsRoute('/components/react.html?tab=usage')).toBe(
    '../../playground/',
  );
  expect(playgroundHrefFromDocsRoute('/en/components/react.html')).toBe(
    '../../../playground/',
  );
  expect(playgroundHrefFromDocsRoute('/0.1.0/en/components/react.html')).toBe(
    '../../../../playground/',
  );
});

test('derives deployment-relative Playground assets from localized version routes', () => {
  expect(
    playgroundAssetHrefFromDocsRoute(
      '/0.1.0/en/automation/',
      '/downloads/a3s-office-skill.tar.gz',
    ),
  ).toBe('../../../../playground/downloads/a3s-office-skill.tar.gz');
  expect(playgroundAssetHrefFromDocsRoute('/', '')).toBe('../playground/');
});

test('keeps cross-mount navigation valid for root and GitHub Pages deployments', () => {
  expect(productSiteBase('/docs/')).toBe('/');
  expect(productSiteBase('/Office/docs/')).toBe('/Office/');
  expect(sitePathFromRoute('/components/document.html', '/Office/docs/')).toBe(
    '/Office/docs/components/document.html',
  );
  expect(
    siteNavigationHref(
      '/Office/docs/components/document.html',
      '/Office/docs/',
      '/playground/index.html',
    ),
  ).toBe('../../playground/index.html');
  expect(siteNavigationHref('/', '/Office/', '/docs/index.html')).toBe(
    '/Office/docs/index.html',
  );
  expect(isProductHomeRoute('/Office/en/', '/Office/')).toBe(true);
  expect(isProductHomeRoute('/Office/docs/', '/Office/docs/')).toBe(false);
});

test('uses the shared A3S navigation with Playground as a primary route', async () => {
  const themeRoot = path.resolve(import.meta.dirname, '../website/theme');
  const websiteRoot = path.resolve(import.meta.dirname, '../website');
  const [
    config,
    docsConfig,
    themeEntry,
    contentStyles,
    homeStyles,
    navSource,
    logo,
  ] = await Promise.all([
    readFile(path.join(websiteRoot, 'rspress.config.ts'), 'utf8'),
    readFile(path.join(websiteRoot, 'rspress.docs.config.ts'), 'utf8'),
    readFile(path.join(themeRoot, 'index.tsx'), 'utf8'),
    readFile(path.join(themeRoot, 'docs-content.css'), 'utf8'),
    readFile(path.join(themeRoot, 'index.css'), 'utf8'),
    readFile(path.join(themeRoot, 'components/Nav.tsx'), 'utf8'),
    readFile(path.resolve(import.meta.dirname, '../docs/public/a3s-logo.png')),
  ]);
  const themeStyles = `${contentStyles}\n${homeStyles}`;

  expect(config).toContain("logo: '/a3s-logo.png'");
  expect(config).toContain(
    "root: path.resolve(import.meta.dirname, 'product')",
  );
  expect(config).toContain("{ text: '文档', link: '/docs/' }");
  expect(config).toContain("{ text: 'Playground', link: '/playground/' }");
  expect(docsConfig).toContain(
    "root: path.resolve(import.meta.dirname, '../docs')",
  );
  expect(docsConfig).toContain('base: docsBase');
  expect(themeEntry).toContain("export { Nav } from './components/Nav'");
  expect(navSource).toContain('aria-controls="office-mobile-navigation"');
  expect(navSource).toContain(
    "language === 'zh' ? '打开导航' : 'Open navigation'",
  );
  expect(navSource).toContain('siteNavigationHref(pathname, site.base');
  expect(navSource).toContain("'/playground/index.html'");
  expect(navSource).toContain("'/index.html'");
  expect(navSource).toContain('<NavScreenMenu menuItems={menuItems} />');
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
