import path from 'node:path';
import { defineConfig } from '@rspress/core';

const playgroundBase = normalizeBase(
  process.env.A3S_OFFICE_PLAYGROUND_BASE ?? '/',
);
const docsBase = `${playgroundBase}docs/`;
const siteOrigin = process.env.DOCS_ORIGIN ?? 'https://a3s-lab.github.io';

export default defineConfig({
  root: path.resolve(import.meta.dirname, '../docs'),
  themeDir: path.resolve(import.meta.dirname, 'theme'),
  base: docsBase,
  siteOrigin,
  title: 'A3S Office',
  description:
    'Documentation for embedding A3S Office editors, extending their behavior, and automating Office files with the CLI and coding-agent Skill.',
  lang: 'en',
  icon: '/favicon.svg',
  logo: '/a3s-office-mark.svg',
  logoText: 'A3S Office',
  outDir: path.resolve(import.meta.dirname, '../playground-dist/docs'),
  llms: true,
  head: [
    ['meta', { name: 'theme-color', content: '#f7f7f8' }],
    ['meta', { property: 'og:type', content: 'website' }],
    ['meta', { property: 'og:site_name', content: 'A3S Office Docs' }],
    (route) => [
      'link',
      {
        rel: 'canonical',
        href: `${siteOrigin}${docsBase.replace(/\/$/, '')}${route.routePath}`,
      },
    ],
  ],
  themeConfig: {
    darkMode: false,
    search: true,
    enableContentAnimation: true,
    editLink: {
      docRepoBaseUrl: 'https://github.com/A3S-Lab/Office/tree/main/docs',
    },
    lastUpdated: true,
    llmsUI: {
      placement: 'outline',
      viewOptions: ['markdownLink', 'chatgpt', 'claude'],
    },
    socialLinks: [
      {
        icon: 'github',
        mode: 'link',
        content: 'https://github.com/A3S-Lab/Office',
      },
    ],
  },
});

function normalizeBase(value: string): string {
  const withLeadingSlash = value.startsWith('/') ? value : `/${value}`;
  return withLeadingSlash.endsWith('/')
    ? withLeadingSlash
    : `${withLeadingSlash}/`;
}
