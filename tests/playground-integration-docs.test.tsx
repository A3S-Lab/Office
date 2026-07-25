import { expect, test } from '@rstest/core';
import { render, screen, waitFor, within } from '@testing-library/react';
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
  ).toEqual(['前端组件', '命令行与 AI']);
  expect(
    within(guideNavigation).getByRole('link', { name: '前端组件' }),
  ).toHaveAttribute('href', '#guide/components');
  expect(
    within(guideNavigation).getByRole('link', { name: '命令行与 AI' }),
  ).toHaveAttribute('href', '#guide/automation');

  expect(
    screen.getByRole('heading', { name: '前端组件', level: 2 }),
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

test('redirects old CLI and Skill links to the combined setup section', async () => {
  window.history.replaceState(null, '', '#skill');

  render(<IntegrationDocsPage {...pageProps} />);

  await waitFor(() => expect(window.location.hash).toBe('#guide/automation'));
  expect(
    screen.getByRole('heading', { name: '命令行与 AI', level: 2 }),
  ).toBeVisible();
});
