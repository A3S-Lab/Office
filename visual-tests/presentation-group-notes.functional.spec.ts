import { expect, type Page, test } from '@playwright/test';

test('Presentation advertises group shortcuts and toggles View notes', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });

  await openPresentationFixture(page);

  await expect(page.locator('.work-presentation-editor')).toHaveAttribute(
    'data-notes',
    'visible',
  );
  await expect(page.getByLabel('演讲者备注')).toBeVisible();

  await page.getByRole('tab', { name: '视图' }).click();
  const notesToggle = page.getByRole('button', { name: '备注' });
  await notesToggle.click();
  await expect(page.locator('.work-presentation-editor')).toHaveAttribute(
    'data-notes',
    'hidden',
  );
  await notesToggle.click();
  await expect(page.locator('.work-presentation-editor')).toHaveAttribute(
    'data-notes',
    'visible',
  );
  await expect(page.getByLabel('演讲者备注')).toBeVisible();

  if (testInfo.project.name === 'desktop-1280') {
    await page.getByRole('tab', { name: '开始' }).click();
    await page
      .locator('.work-slide-canvas.interactive > .work-slide-element')
      .first()
      .click();
    const arrange = page.locator(
      '.work-presentation-ribbon .work-office-ribbon-group[aria-label="排列"]',
    );
    await expect(arrange).toBeVisible();
    await expect(page.locator('button[aria-label="组合"]')).toHaveAttribute(
      'aria-keyshortcuts',
      'Control+G Meta+G',
    );
    await expect(page.locator('button[aria-label="取消组合"]')).toHaveAttribute(
      'aria-keyshortcuts',
      'Control+Shift+G Meta+Shift+G',
    );
    await expect(page.locator('button[aria-label="组合"]')).toBeDisabled();
    await expect(page.locator('button[aria-label="取消组合"]')).toBeDisabled();
  }

  await page.screenshot({
    path: testInfo.outputPath(
      `presentation-group-notes-${testInfo.project.name}.png`,
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
