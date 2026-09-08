import { expect, test } from '@playwright/test';

test('Markdown applies bold and opens link dialog with WPS shortcuts', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });

  await page.goto('/playground/');
  await page.getByRole('button', { name: '# 产品说明 MD · 本次会话' }).click();
  await expect(
    page.locator('.work-markdown-editor .work-markdown-ribbon'),
  ).toBeVisible();

  const ribbon = page.locator('.work-markdown-ribbon');
  await expect(ribbon.getByRole('button', { name: '加粗' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+B Meta+B',
  );
  await expect(ribbon.getByRole('button', { name: '斜体' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+I Meta+I',
  );

  const source = page.getByRole('textbox', { name: 'Markdown 源码' });
  await source.click();
  await page.keyboard.press('Control+b');
  await expect(ribbon.getByRole('button', { name: '加粗' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );

  await page.getByRole('tab', { name: '插入' }).click();
  await expect(
    ribbon.getByRole('button', { name: '添加链接' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+K Meta+K');
  await source.click();
  await page.keyboard.press('Control+k');
  await expect(page.getByRole('dialog', { name: '添加链接' })).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(source).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `markdown-format-shortcuts-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});
