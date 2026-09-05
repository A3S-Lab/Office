import { expect, test } from '@playwright/test';

test('documentation changelog stays scannable, localized, and version-aware', async ({
  page,
}, testInfo) => {
  await page.goto('/docs/changelog.html');

  await expect(
    page.getByRole('heading', { level: 1, name: '更新日志' }),
  ).toBeVisible();
  const cards = page.locator('.office-release-card');
  await expect(cards).toHaveCount(30);
  await expect(cards.first()).toContainText('v0.52.0');
  await expect(cards.first()).toContainText(
    'Writer 整段修订保持原生与原子语义',
  );
  await expect(cards.nth(1)).toContainText('v0.51.0');
  await expect(cards.first()).toContainText('一个段落，一项决定');
  await expect(cards.nth(1)).toContainText('Writer 比较配对同一分节内文字移动');
  await expect(cards.nth(1)).toContainText('段落范围保持完整');
  await expect(cards.nth(2)).toContainText('v0.50.0');
  await expect(cards.nth(2)).toContainText('Writer 比较识别有界文字移动');
  await expect(cards.nth(3)).toContainText('v0.49.0');
  await expect(cards.nth(3)).toContainText('Writer 移动修订保持成对并原生往返');
  await expect(cards.nth(4)).toContainText('v0.48.1');
  await expect(cards.nth(4)).toContainText('Writer 选区控件恢复紧凑视觉契约');
  await expect(cards.nth(4)).toContainText('浏览器原生默认样式');
  await expect(
    page.locator('.office-release-card[data-version="0.41.0"]'),
  ).toContainText('表格数据验证警告现在与 Office 决策一致');
  await expect(
    page.locator('.office-release-card[data-version="0.34.0"]'),
  ).toContainText('演示文稿入场动画');
  await expect(cards.first().locator('time')).toHaveAttribute(
    'datetime',
    '2026-09-05',
  );
  await expect(
    cards.first().locator('.office-release-card__highlights > li'),
  ).toHaveCount(3);

  const geometry = await page.evaluate(() => {
    const highlights = document.querySelector<HTMLElement>(
      '.office-release-card__highlights',
    );
    if (!highlights) throw new Error('Release highlights are missing.');
    return {
      documentClientWidth: document.documentElement.clientWidth,
      documentScrollWidth: document.documentElement.scrollWidth,
      highlightColumns:
        getComputedStyle(highlights).gridTemplateColumns.split(' ').length,
    };
  });
  expect(geometry.documentScrollWidth).toBeLessThanOrEqual(
    geometry.documentClientWidth + 1,
  );
  expect(geometry.highlightColumns).toBe(
    testInfo.project.name === 'compact-768' ? 1 : 3,
  );

  await page.screenshot({
    fullPage: true,
    path: testInfo.outputPath('docs-changelog.png'),
  });

  await page.getByRole('link', { name: '阅读移动修订指南' }).click();
  await expect(page.locator('h2#移动修订')).toBeInViewport();

  await page.goto('/docs/changelog.html');
  await page.getByRole('link', { name: '阅读协作合同' }).click();
  await expect(page.locator('h3#同步移动修订')).toBeInViewport();

  await page.goto('/docs/changelog.html');
  await page.getByRole('link', { name: '阅读选区工具栏说明' }).click();
  await expect(page.locator('h2#选择工具栏控件')).toBeInViewport();

  await page.goto('/docs/changelog.html');
  await page.getByRole('link', { name: '阅读内容控件指南' }).click();
  await expect(page.locator('h2#原生内容控件')).toBeInViewport();

  await page.goto('/docs/changelog.html');
  await page.getByRole('link', { name: '阅读公式条件格式指南' }).click();
  await expect(page.locator('h2#公式条件格式')).toBeInViewport();

  await page.goto('/docs/0.38.0/changelog.html');
  await expect(page.locator('.office-release-card')).toHaveCount(14);
  await expect(page.locator('.office-release-card').first()).toContainText(
    'v0.38.0',
  );
  await expect(page.getByText('v0.38.1', { exact: true })).toHaveCount(0);

  await page.goto('/docs/en/changelog.html');
  await expect(
    page.getByRole('heading', { level: 1, name: "What's new" }),
  ).toBeVisible();
  await expect(
    page.locator('.office-release-card[data-version="0.43.0"]'),
  ).toContainText('Spreadsheet conditional formatting is now formula-editable');
  await expect(
    page.locator('.office-release-card[data-version="0.42.0"]'),
  ).toContainText('Spreadsheet rules can now be local custom formulas');
});
