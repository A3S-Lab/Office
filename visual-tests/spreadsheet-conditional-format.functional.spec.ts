import { expect, test } from '@playwright/test';

for (const viewport of [
  { height: 800, name: 'desktop', width: 1280 },
  { height: 800, name: 'compact', width: 768 },
] as const) {
  test(`Spreadsheet authors local formula conditional formatting at ${viewport.name} width`, async ({
    page,
  }, testInfo) => {
    const browserErrors: string[] = [];
    page.on('console', (message) => {
      if (message.type() === 'error') browserErrors.push(message.text());
    });
    page.on('pageerror', (error) => browserErrors.push(error.message));
    await page.setViewportSize(viewport);
    await page.goto('/playground/');
    await page.locator("button[data-template-id='conditional-format']").click();

    const grid = page.locator('.fortune-sheet-overlay');
    await expect(grid).toBeVisible();
    await expect(page.getByRole('textbox', { name: '文件名' })).toHaveValue(
      '公式条件格式示例',
    );

    const ribbon = page.locator('.work-spreadsheet-ribbon');
    await ribbon.getByRole('button', { name: '条件格式' }).click();
    const panel = page.locator('.work-spreadsheet-conditional-manager');
    await expect(panel).toBeVisible();
    await expect(
      panel.getByRole('textbox', { name: '条件格式公式' }),
    ).toHaveValue('=$D4="阻塞"');
    await expect(panel).toContainText('只读取本地缓存值');
    await expect(
      panel.getByRole('checkbox', { name: '公式规则使用文字颜色' }),
    ).toBeChecked();
    await expect(
      panel.getByRole('checkbox', { name: '公式规则使用填充颜色' }),
    ).toBeChecked();

    await panel
      .getByRole('textbox', { name: '条件格式公式' })
      .fill('=AND($D4="阻塞",$C4<1)');
    await panel.getByRole('button', { name: '保存规则' }).click();
    await expect(
      panel.getByRole('textbox', { name: '条件格式公式' }),
    ).toHaveValue('=AND($D4="阻塞",$C4<1)');

    await page.screenshot({
      path: testInfo.outputPath(
        `spreadsheet-conditional-format-${viewport.name}.png`,
      ),
      animations: 'disabled',
    });
    expect(browserErrors).toEqual([]);
  });
}
