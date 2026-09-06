import { expect, type Page, test } from '@playwright/test';

test('Spreadsheet phone Find keeps touch controls and restores grid focus', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));

  await page.setViewportSize({ width: 390, height: 700 });
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  await grid.focus();
  await page.keyboard.press('Control+f');

  const query = page.getByRole('textbox', { name: '查找当前工作表' });
  const bar = page.locator('.work-spreadsheet-find-bar');
  await expect(bar).toBeVisible();
  await expect(query).toBeFocused();

  const geometry = await bar.evaluate((element) => {
    const input = element.querySelector<HTMLElement>(
      '[aria-label="查找当前工作表"]',
    );
    const buttons = Array.from(
      element.querySelectorAll<HTMLElement>(
        '.work-spreadsheet-find-actions button',
      ),
    );
    if (!input || buttons.length !== 3) {
      throw new Error('Phone Spreadsheet Find controls are incomplete.');
    }
    const barBounds = element.getBoundingClientRect();
    const inputBounds = input.getBoundingClientRect();
    return {
      barLeft: barBounds.left,
      barRight: barBounds.right,
      buttonBounds: buttons.map((button) => {
        const bounds = button.getBoundingClientRect();
        return { height: bounds.height, width: bounds.width };
      }),
      inputHeight: inputBounds.height,
      inputWidth: inputBounds.width,
      viewportWidth: document.documentElement.clientWidth,
    };
  });

  expect(geometry.barLeft).toBeGreaterThanOrEqual(0);
  expect(geometry.barRight).toBeLessThanOrEqual(geometry.viewportWidth);
  expect(geometry.inputHeight).toBeGreaterThanOrEqual(40);
  expect(geometry.inputWidth).toBeGreaterThanOrEqual(96);
  for (const bounds of geometry.buttonBounds) {
    expect(bounds.height).toBeGreaterThanOrEqual(40);
    expect(bounds.width).toBeGreaterThanOrEqual(40);
  }

  await query.fill('客户洞察报告');
  await expect(page.getByText('1 个匹配', { exact: true })).toBeVisible();
  await query.press('Enter');
  await expect(page.locator('.fortune-name-box')).toHaveText('A4');
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-phone-find.png'),
    animations: 'disabled',
  });

  await query.press('Escape');
  await expect(bar).toBeHidden();
  await expect(grid).toBeFocused();
  expect(browserErrors).toEqual([]);
});

async function openSpreadsheetFixture(page: Page): Promise<void> {
  await page.goto('/playground/');
  await page
    .getByRole('button', {
      name: '季度执行计划 XLSX · 本次会话',
    })
    .click();
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();
}
