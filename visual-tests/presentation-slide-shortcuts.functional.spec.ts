import { expect, type Page, test } from '@playwright/test';

test('Presentation manages slides with WPS Ctrl+M / Ctrl+D / Delete', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });

  await openPresentationFixture(page);

  await expect(page.getByRole('button', { name: '新建幻灯片' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+M Meta+Shift+N',
  );
  await expect(page.getByRole('button', { name: '复制幻灯片' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+D Meta+D',
  );
  await expect(page.getByRole('button', { name: '删除幻灯片' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Delete Backspace',
  );

  const strip = page.locator('.work-slide-strip');
  await expect(strip.locator('[data-slide-thumbnail][data-slide-index="2"]')).toBeVisible();

  const active = strip.locator('[data-slide-thumbnail].active');
  await active.click();

  await page.keyboard.press('Control+m');
  await expect(
    strip.locator('[data-slide-thumbnail][data-slide-index="3"].active'),
  ).toBeVisible();

  await page.keyboard.press('Control+d');
  await expect(
    strip.locator('[data-slide-thumbnail][data-slide-index="4"].active'),
  ).toBeVisible();

  await page.keyboard.press('Delete');
  await expect(
    strip.locator('[data-slide-thumbnail][data-slide-index="3"].active'),
  ).toBeVisible();
  await expect(
    strip.locator('[data-slide-thumbnail][data-slide-index="4"]'),
  ).toHaveCount(0);

  await page.screenshot({
    path: testInfo.outputPath(
      `presentation-slide-shortcuts-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});

async function openPresentationFixture(page: Page): Promise<void> {
  await page.goto('/playground/');
  await page
    .getByRole('button', {
      name: '业务策略汇报 PPTX · 本次会话',
    })
    .click();
  await page.locator('.work-slide-canvas.interactive').waitFor();
  await expect(page.locator('.work-presentation-editor')).toBeVisible();
}
