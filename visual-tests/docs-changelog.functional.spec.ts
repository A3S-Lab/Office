import { expect, test } from '@playwright/test';

test('documentation changelog stays scannable, localized, and version-aware', async ({
  page,
}, testInfo) => {
  await page.goto('/docs/changelog.html');

  await expect(
    page.getByRole('heading', { level: 1, name: '更新日志' }),
  ).toBeVisible();
  const cards = page.locator('.office-release-card');
  await expect(cards).toHaveCount(15);
  await expect(cards.first()).toContainText('v0.38.1');
  await expect(cards.nth(1)).toContainText('Writer 原生 OpenType 排版');
  await expect(
    page.locator('.office-release-card[data-version="0.34.0"]'),
  ).toContainText('演示文稿入场动画');
  await expect(cards.first().locator('time')).toHaveAttribute(
    'datetime',
    '2026-09-01',
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

  await page.getByRole('link', { name: '阅读 Writer 指南' }).click();
  await expect(page.locator('h2#原生-opentype-排版')).toBeInViewport();

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
  await expect(page.locator('.office-release-card').first()).toContainText(
    'Release notes that explain the product change',
  );
});
