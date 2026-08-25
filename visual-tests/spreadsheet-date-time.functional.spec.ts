import { expect, test } from '@playwright/test';

const FIXED_LOCAL_TIME = new Date('2026-08-22T01:07:58.900Z');

test('Spreadsheet inserts static WPS date and time values with one-step history', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.clock.setFixedTime(FIXED_LOCAL_TIME);
  await page.goto('/playground/?e2e=spreadsheet-date-time');

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const revision = page.getByTestId('spreadsheet-date-time-status');
  const numberGroup = page
    .locator('.work-spreadsheet-ribbon')
    .getByRole('region', { name: '数字' });
  const formatSelect = numberGroup.getByRole('combobox', {
    name: '数字格式',
  });
  const dateTime = numberGroup.getByRole('button', { name: '日期和时间' });

  await expect(grid).toBeVisible();
  await expect(revision).toHaveAttribute('data-revision', '1');
  await expect(nameBox).toHaveText('B2');
  await expect(formulaBar).toHaveText('等待输入');
  await expect(dateTime).toHaveAttribute(
    'title',
    '插入当前日期或时间（Ctrl+; / Ctrl+Shift+;）',
  );

  await dateTime.click();
  const menu = page.getByRole('menu', { name: '插入日期和时间' });
  const insertDate = menu.getByRole('menuitem', { name: '插入当前日期' });
  const insertTime = menu.getByRole('menuitem', { name: '插入当前时间' });
  await expect(insertDate).toHaveAttribute('aria-keyshortcuts', 'Control+;');
  await expect(insertTime).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+;',
  );
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-date-time-menu.png'),
    animations: 'disabled',
  });
  await insertDate.click();

  await expect(grid).toBeFocused();
  await expect(formulaBar).toHaveText('2026-08-22');
  await expect(formatSelect).toContainText('短日期');
  await expect(revision).toHaveAttribute('data-revision', '2');

  await page.keyboard.press('Control+z');
  await expect(formulaBar).toHaveText('等待输入');
  await expect(formatSelect).toContainText('常规');
  await expect(revision).toHaveAttribute('data-revision', '3');
  await page.keyboard.press('Control+Shift+z');
  await expect(formulaBar).toHaveText('2026-08-22');
  await expect(formatSelect).toContainText('短日期');
  await expect(revision).toHaveAttribute('data-revision', '4');

  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('B3');
  await expect(formulaBar).toHaveText('等待输入');
  await page.keyboard.press('Control+Shift+;');
  await expect(formulaBar).toHaveText('09:07');
  await expect(formatSelect).toContainText('时间');
  await expect(revision).toHaveAttribute('data-revision', '5');

  await page.keyboard.press('Control+z');
  await expect(formulaBar).toHaveText('等待输入');
  await expect(formatSelect).toContainText('常规');
  await expect(revision).toHaveAttribute('data-revision', '6');
  await page.keyboard.press('ArrowUp');
  await expect(formulaBar).toHaveText('2026-08-22');
  await expect(formatSelect).toContainText('短日期');
  expect(browserErrors).toEqual([]);
});
