import path from 'node:path';
import { defineConfig } from '@rspress/core';
import { siteBaseFromEnvironment } from './site-paths';

const siteBase = siteBaseFromEnvironment();
const siteOrigin = process.env.DOCS_ORIGIN ?? 'https://a3s-lab.github.io';

const zhNav = [
  { text: '文档', link: '/docs/' },
  { text: 'Playground', link: '/playground/' },
  { text: '协作', link: '/docs/components/collaboration.html' },
];

const enNav = [
  { text: 'Docs', link: '/docs/' },
  { text: 'Playground', link: '/playground/' },
  { text: 'Collaboration', link: '/docs/components/collaboration.html' },
];

const resourceItems = {
  text: '资源',
  position: 'right' as const,
  items: [
    { text: 'GitHub', link: 'https://github.com/A3S-Lab/Office' },
    {
      text: 'npm',
      link: 'https://www.npmjs.com/package/@a3s-lab/office',
    },
  ],
};

const resourceItemsEn = {
  ...resourceItems,
  text: 'Resources',
};

export default defineConfig({
  root: path.resolve(import.meta.dirname, 'product'),
  themeDir: path.resolve(import.meta.dirname, 'product-theme'),
  base: siteBase,
  siteOrigin,
  title: 'A3S Office',
  description:
    'Bring document, Markdown, spreadsheet, presentation, and PDF editing into your product with one host contract.',
  lang: 'zh',
  locales: [
    {
      lang: 'zh',
      label: '简体中文',
      title: 'A3S Office',
      description:
        '将文档、Markdown、表格、演示文稿和 PDF 编辑能力带进你的产品。',
    },
    {
      lang: 'en',
      label: 'English',
      title: 'A3S Office',
      description:
        'Bring document, Markdown, spreadsheet, presentation, and PDF editing into your product.',
    },
  ],
  icon: '/favicon.svg',
  logo: '/a3s-logo.png',
  logoText: 'A3S Office',
  logoHref: '/',
  outDir: path.resolve(import.meta.dirname, '../playground-dist'),
  head: [
    ['meta', { name: 'theme-color', content: '#f7faff' }],
    ['meta', { property: 'og:type', content: 'website' }],
    ['meta', { property: 'og:site_name', content: 'A3S Office' }],
    (route) => [
      'link',
      {
        rel: 'canonical',
        href: `${siteOrigin}${siteBase.replace(/\/$/, '')}${route.routePath}`,
      },
    ],
  ],
  builderConfig: {
    server: {
      publicDir: {
        name: path.resolve(import.meta.dirname, '../docs/public'),
      },
    },
  },
  themeConfig: {
    darkMode: 'force-light',
    enableContentAnimation: true,
    nav: [...zhNav, resourceItems],
    locales: [
      { lang: 'zh', label: '简体中文', nav: [...zhNav, resourceItems] },
      { lang: 'en', label: 'English', nav: [...enNav, resourceItemsEn] },
    ],
    socialLinks: [
      {
        icon: 'github',
        mode: 'link',
        content: 'https://github.com/A3S-Lab/Office',
      },
    ],
  },
});
