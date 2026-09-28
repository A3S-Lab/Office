import { expect, type TestInfo, test } from '@playwright/test';

function desktopOnly(testInfo: TestInfo) {
  test.skip(testInfo.project.name !== 'desktop-1280');
}

test('Spreadsheet Escape after AutoFilter restores grid focus for ArrowDown', async ({
  page,
}, testInfo) => {
  desktopOnly(testInfo);
  await page.goto('/playground/');
  await page.locator("button[data-template-id='blank-spreadsheet']").click();
  const grid = page.locator('.work-spreadsheet-editor .fortune-sheet-overlay');
  await expect(grid).toBeVisible();
  const nameBox = page.locator('.fortune-name-box');

  await grid.focus();
  await page.keyboard.press('Control+Home');
  await expect(nameBox).toHaveText('A1');
  await page.keyboard.type('Header');
  await page.keyboard.press('Enter');
  await page.keyboard.type('10');
  await page.keyboard.press('Enter');
  await page.keyboard.type('20');
  await page.keyboard.press('Enter');
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  await page
    .locator(".work-spreadsheet-ribbon [role='tab'][data-tab-id='data']")
    .click();
  await page.getByRole('button', { name: '自动筛选' }).click();
  await expect(
    page.locator(".work-spreadsheet-editor[data-auto-filter='active']"),
  ).toBeVisible();

  await grid.focus();
  await page.keyboard.press('Control+Home');
  await expect(nameBox).toHaveText('A1');
  await page.keyboard.press('Escape');
  await expect
    .poll(() =>
      page.evaluate(() => ({
        className: document.activeElement?.className ?? '',
        id: document.activeElement?.id ?? '',
      })),
    )
    .toMatchObject({
      className: expect.stringContaining('fortune-sheet-overlay'),
    });
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('A2');
});
