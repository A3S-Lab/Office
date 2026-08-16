import path from 'node:path';
import { defineConfig } from '@rspress/core';
import {
  DOCUMENTATION_DEFAULT_LANGUAGE,
  DOCUMENTATION_DEFAULT_VERSION,
  DOCUMENTATION_LOCALES,
  DOCUMENTATION_VERSIONS,
} from './documentation-site';

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
    '在应用中接入、扩展 A3S Office 编辑器，并通过 CLI 与编码智能体处理 Office 文件。',
  lang: DOCUMENTATION_DEFAULT_LANGUAGE,
  locales: [...DOCUMENTATION_LOCALES],
  multiVersion: {
    default: DOCUMENTATION_DEFAULT_VERSION,
    versions: [...DOCUMENTATION_VERSIONS],
  },
  route: {
    localeRedirect: 'never',
  },
  search: {
    mode: 'local',
    versioned: true,
  },
  icon: '/favicon.svg',
  logo: '/a3s-logo.png',
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
    search: true,
    fallbackHeadingTitle: false,
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
