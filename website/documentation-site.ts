export const DOCUMENTATION_DEFAULT_LANGUAGE = 'zh';

export const DOCUMENTATION_LOCALES = [
  {
    lang: 'zh',
    label: '简体中文',
    title: 'A3S Office 文档',
    description:
      '在应用中接入、扩展 A3S Office 编辑器，并通过 CLI 与编码智能体处理 Office 文件。',
  },
  {
    lang: 'en',
    label: 'English',
    title: 'A3S Office documentation',
    description:
      'Embed and extend A3S Office editors, then automate Office files with the CLI and coding-agent Skill.',
  },
] as const;

export const DOCUMENTATION_DEFAULT_VERSION = 'latest';

export const DOCUMENTATION_VERSIONS = [
  'latest',
  '0.153.0',
  '0.152.0',
  '0.151.0',
  '0.150.0',
  '0.149.0',
  '0.148.0',
  '0.147.0',
  '0.146.0',
  '0.145.0',
  '0.144.0',
  '0.143.0',
  '0.142.0',
  '0.141.0',
  '0.140.0',
  '0.139.0',
  '0.138.0',
  '0.137.0',
  '0.136.0',
  '0.135.0',
  '0.134.0',
  '0.38.0',
  '0.1.0',
] as const;

export const DOCUMENTATION_REQUIRED_ROUTES = [
  'index.mdx',
  'changelog.mdx',
  'guide/index.mdx',
  'components/index.mdx',
  'components/react.mdx',
  'components/vue.mdx',
  'components/web-component.mdx',
  'components/document.mdx',
  'components/markdown.mdx',
  'components/spreadsheet.mdx',
  'components/presentation.mdx',
  'components/pdf.mdx',
  'components/extensions.mdx',
  'automation/index.mdx',
  'browser-editor-architecture.md',
  'native-office-engine.md',
  'editor-quality-roadmap.md',
  'cli-reference.md',
] as const;
