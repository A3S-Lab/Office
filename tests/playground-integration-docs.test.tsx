import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { IntegrationDocsPage } from '../playground/src/integration-docs-page';

const pageProps = {
  rawSkillUrl: '/downloads/a3s-office-skill/SKILL.md',
  sidebarOpen: true,
  skillDownloadUrl: '/downloads/a3s-office-skill.tar.gz',
  onOpenSidebar: () => undefined,
};

test('shows component, CLI, and Skill setup in one continuous guide', () => {
  window.history.replaceState(null, '', '#guide');

  render(<IntegrationDocsPage {...pageProps} />);

  expect(
    screen.getByRole('heading', { name: '接入文档', level: 1 }),
  ).toBeVisible();
  expect(
    screen.queryByRole('tablist', { name: '接入内容' }),
  ).not.toBeInTheDocument();

  const guideNavigation = screen.getByRole('navigation', {
    name: '接入方式',
  });
  expect(
    within(guideNavigation)
      .getAllByRole('link')
      .map((link) => link.textContent),
  ).toEqual(['前端组件', '组件 API', '命令行与 AI']);
  expect(
    within(guideNavigation).getByRole('link', { name: '前端组件' }),
  ).toHaveAttribute('href', '#guide/components');
  expect(
    within(guideNavigation).getByRole('link', { name: '组件 API' }),
  ).toHaveAttribute('href', '#guide/api');
  expect(
    within(guideNavigation).getByRole('link', { name: '命令行与 AI' }),
  ).toHaveAttribute('href', '#guide/automation');

  expect(
    screen.getByRole('heading', { name: '前端组件', level: 2 }),
  ).toBeVisible();
  expect(
    screen.getByRole('heading', { name: '组件 API', level: 2 }),
  ).toBeVisible();
  expect(screen.getByRole('heading', { name: '安装', level: 3 })).toBeVisible();
  expect(
    screen.getByRole('heading', { name: '命令行与 AI', level: 2 }),
  ).toBeVisible();
  expect(
    screen.getByText('a3s-office validate report.docx --json'),
  ).toBeVisible();
  expect(screen.getByRole('link', { name: '下载 CLI Skill' })).toHaveAttribute(
    'href',
    '/downloads/a3s-office-skill.tar.gz',
  );
});

test('documents every editor contract and its supported extension boundary', () => {
  render(<IntegrationDocsPage {...pageProps} />);

  const tablist = screen.getByRole('tablist', { name: '编辑器 API' });
  const editors = [
    ['DocumentEditor', '文字'],
    ['MarkdownEditor', 'Markdown'],
    ['SpreadsheetEditor', '表格'],
    ['PresentationEditor', '演示'],
    ['PdfViewer', 'PDF'],
  ] as const;

  expect(within(tablist).getAllByRole('tab')).toHaveLength(editors.length);

  for (const [component, label] of editors) {
    const tab = within(tablist).getByRole('tab', {
      name: new RegExp(`${label}.*${component}`),
    });
    fireEvent.click(tab);
    const panel = screen.getByRole('tabpanel', {
      name: new RegExp(component),
    });
    expect(
      within(panel).getByRole('heading', { name: component, level: 3 }),
    ).toBeVisible();
    expect(
      within(panel).getByRole('heading', { name: '属性参数', level: 3 }),
    ).toBeVisible();
    expect(
      within(panel).getByRole('heading', {
        name: component === 'PdfViewer' ? '文件生命周期' : '内容模型',
        level: 3,
      }),
    ).toBeVisible();
  }

  fireEvent.click(
    within(tablist).getByRole('tab', {
      name: /文字.*DocumentEditor/,
    }),
  );
  const documentPanel = screen.getByRole('tabpanel', {
    name: /DocumentEditor/,
  });
  expect(
    within(documentPanel).getByRole('rowheader', { name: 'extensions' }),
  ).toBeVisible();
  expect(
    within(documentPanel).getByRole('rowheader', {
      name: 'getSelectionMenuItems',
    }),
  ).toBeVisible();
  expect(
    within(documentPanel).getByRole('heading', {
      name: '选区右键菜单',
      level: 3,
    }),
  ).toBeVisible();
  expect(
    within(documentPanel).getByText('支持 TipTap Extensions'),
  ).toBeVisible();
  const examples = documentPanel.querySelectorAll(
    'pre[data-code-language="tsx"]',
  );
  expect(examples).toHaveLength(2);
  expect(examples[0]).toHaveTextContent('GetDocumentSelectionMenuItems');
  expect(examples[0]).toHaveTextContent('context.document.content');
  expect(examples[0]).toHaveTextContent('context.commands.replaceText');
  expect(examples[1]).toHaveTextContent('Extension.create({');

  fireEvent.click(
    within(tablist).getByRole('tab', {
      name: /PDF.*PdfViewer/,
    }),
  );
  const pdfPanel = screen.getByRole('tabpanel', { name: /PdfViewer/ });
  expect(
    within(pdfPanel).getByRole('rowheader', { name: 'loadSource' }),
  ).toBeVisible();
  expect(within(pdfPanel).getByText('当前使用加载与保存端口')).toBeVisible();
});

test('redirects old CLI and Skill links to the combined setup section', async () => {
  window.history.replaceState(null, '', '#skill');

  render(<IntegrationDocsPage {...pageProps} />);

  await waitFor(() => expect(window.location.hash).toBe('#guide/automation'));
  expect(
    screen.getByRole('heading', { name: '命令行与 AI', level: 2 }),
  ).toBeVisible();
});
