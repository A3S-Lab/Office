import { expect, test } from '@playwright/test';

test('Spreadsheet copies the exact formula or value from above without replacing target styles', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.goto('/?e2e=spreadsheet-copy-from-above');

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const status = page.getByTestId('spreadsheet-copy-from-above-status');
  const preservedStyle = {
    'data-target-bold': '1',
    'data-target-fill': '#ddebf7',
    'data-target-format': '0.00',
    'data-target-italic': '1',
  };

  await expect(grid).toBeVisible();
  await expect(grid).toBeFocused();
  await expect(nameBox).toHaveText('B2');
  await expect(formulaBar).toHaveText('Replace me');
  await expect(status).toHaveAttribute('data-revision', '1');
  await expect(status).toHaveAttribute('data-target-formula', '');
  await expect(status).toHaveAttribute('data-target-value', 'Replace me');

  await page.keyboard.press("Control+'");
  await expect(status).toHaveAttribute('data-revision', '2');
  await expect(status).toHaveAttribute('data-target-formula', '=$A$1+A1');
  await expect(formulaBar).toHaveText('=$A$1+A1');
  for (const [name, value] of Object.entries(preservedStyle)) {
    await expect(status).toHaveAttribute(name, value);
  }

  await page.keyboard.press('Control+z');
  await expect(status).toHaveAttribute('data-revision', '3');
  await expect(status).toHaveAttribute('data-target-formula', '');
  await expect(status).toHaveAttribute('data-target-value', 'Replace me');

  await page.keyboard.press("Control+Shift+'");
  await expect(status).toHaveAttribute('data-revision', '4');
  await expect(status).toHaveAttribute('data-target-formula', '');
  await expect(status).toHaveAttribute('data-target-value', '6');
  await expect(formulaBar).toHaveText('6');
  for (const [name, value] of Object.entries(preservedStyle)) {
    await expect(status).toHaveAttribute(name, value);
  }

  await page.keyboard.press('Control+z');
  await expect(status).toHaveAttribute('data-revision', '5');
  await expect(status).toHaveAttribute('data-target-formula', '');
  await expect(status).toHaveAttribute('data-target-value', 'Replace me');
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-copy-from-above.png'),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});
