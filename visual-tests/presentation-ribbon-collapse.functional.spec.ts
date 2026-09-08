import { expect, type Page, test } from '@playwright/test';

test('Presentation collapses and expands ribbon with WPS Ctrl+F1', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });

  await openPresentationFixture(page);

  const ribbon = page.locator('.work-presentation-ribbon');
  const collapse = ribbon.locator('.work-office-ribbon-collapse');
  await expect(collapse).toHaveAttribute('aria-label', '折叠功能区');
  await expect(collapse).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+F1 Meta+F1',
  );

  await page.keyboard.press('Control+F1');
  await expect(ribbon).toHaveAttribute('data-collapsed', 'true');
  await expect(collapse).toHaveAttribute('aria-label', '展开功能区');
  await expect(collapse).toHaveAttribute('aria-expanded', 'false');

  await page.keyboard.press('Control+F1');
  await expect(ribbon).not.toHaveAttribute('data-collapsed', 'true');
  await expect(collapse).toHaveAttribute('aria-label', '折叠功能区');
  await expect(collapse).toHaveAttribute('aria-expanded', 'true');

  await page.screenshot({
    path: testInfo.outputPath(
      `presentation-ribbon-collapse-${testInfo.project.name}.png`,
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
