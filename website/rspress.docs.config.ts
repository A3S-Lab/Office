import path from 'node:path';
import { defineConfig } from '@rspress/core';
import {
  DOCUMENTATION_DEFAULT_LANGUAGE,
  DOCUMENTATION_DEFAULT_VERSION,
  DOCUMENTATION_LOCALES,
  DOCUMENTATION_VERSIONS,
} from './documentation-site';
import { siteBaseFromEnvironment } from './site-paths';

const siteBase = siteBaseFromEnvironment();
const docsBase = `${siteBase}docs/`;
const siteOrigin = process.env.DOCS_ORIGIN ?? 'https://a3s-lab.github.io';

export default defineConfig({
  root: path.resolve(import.meta.dirname, '../docs'),
  themeDir: path.resolve(import.meta.dirname, 'theme'),
  base: docsBase,
  siteOrigin,
  title: 'A3S Office Docs',
  description:
    'Versioned documentation for integrating A3S Office editors, collaboration, and automation.',
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
  outDir: path.resolve(import.meta.dirname, '../.docs-build'),
  llms: true,
  head: [
    ['meta', { name: 'theme-color', content: '#f5f7fb' }],
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
  builderConfig: {
    server: {
      publicDir: {
        name: path.resolve(import.meta.dirname, '../docs/public'),
      },
    },
  },
  themeConfig: {
    search: true,
    fallbackHeadingTitle: false,
    enableContentAnimation: false,
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
