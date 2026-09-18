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
  '0.281.0',
  '0.280.0',
  '0.279.0',
  '0.278.0',
  '0.277.0',
  '0.276.0',
  '0.275.0',
  '0.274.0',
  '0.273.0',
  '0.272.0',
  '0.271.0',
  '0.270.0',
  '0.269.0',
  '0.268.0',
  '0.267.0',
  '0.266.0',
  '0.265.0',
  '0.264.0',
  '0.263.0',
  '0.262.0',
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
