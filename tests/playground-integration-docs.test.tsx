import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, within } from '@testing-library/react';
import { IntegrationDocsPage } from '../playground/src/integration-docs-page';

const pageProps = {
  rawSkillUrl: '/downloads/a3s-office-skill/SKILL.md',
  sidebarOpen: true,
  skillDownloadUrl: '/downloads/a3s-office-skill.tar.gz',
  onOpenSidebar: () => undefined,
};

test('combines component, CLI, and Skill guidance in one document page', () => {
  window.history.replaceState(null, '', '#guide');

  render(<IntegrationDocsPage {...pageProps} />);

  expect(
    screen.getByRole('heading', { name: '接入文档', level: 1 }),
  ).toBeVisible();
  const guideNavigation = screen.getByRole('tablist', {
    name: '接入内容',
  });
  expect(
    within(guideNavigation)
      .getAllByRole('tab')
      .map((tab) => tab.textContent),
  ).toEqual(['前端组件', 'Office CLI', 'CLI Skill']);
  expect(
    screen.getByRole('heading', { name: '安装组件', level: 2 }),
  ).toBeVisible();

  fireEvent.click(
    within(guideNavigation).getByRole('tab', { name: 'Office CLI' }),
  );
  expect(window.location.hash).toBe('#guide/cli');
  expect(
    screen.getByRole('heading', { name: 'Office CLI', level: 2 }),
  ).toBeVisible();
  expect(
    screen.getByText('a3s-office validate report.docx --json'),
  ).toBeVisible();

  fireEvent.click(
    within(guideNavigation).getByRole('tab', { name: 'CLI Skill' }),
  );
  expect(window.location.hash).toBe('#guide/skill');
  expect(
    screen.getByRole('heading', { name: 'CLI Skill', level: 2 }),
  ).toBeVisible();
  expect(screen.getByRole('link', { name: '下载 CLI Skill' })).toHaveAttribute(
    'href',
    '/downloads/a3s-office-skill.tar.gz',
  );

  window.history.pushState(null, '', '#cli');
  fireEvent(window, new HashChangeEvent('hashchange'));
  expect(window.location.hash).toBe('#guide/cli');
  expect(
    within(guideNavigation).getByRole('tab', { name: 'Office CLI' }),
  ).toHaveAttribute('aria-selected', 'true');
});

test('keeps legacy Skill links working inside the unified guide', () => {
  window.history.replaceState(null, '', '#skill');

  render(<IntegrationDocsPage {...pageProps} />);

  expect(screen.getByRole('tab', { name: 'CLI Skill' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  expect(window.location.hash).toBe('#guide/skill');
});
