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
  '0.224.0',
  '0.223.0',
  '0.222.0',
  '0.221.0',
  '0.220.0',
  '0.219.0',
  '0.218.0',
  '0.217.0',
  '0.216.0',
  '0.215.0',
  '0.214.0',
  '0.213.0',
  '0.212.0',
  '0.211.0',
  '0.210.0',
  '0.209.0',
  '0.208.0',
  '0.207.0',
  '0.206.0',
  '0.205.0',
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
