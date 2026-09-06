import { expect, type Page, test } from '@playwright/test';

test('Spreadsheet phone worksheet rename explains errors and restores tab actions', async ({
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
  await page.keyboard.press('Shift+F11');

  const createdSheet = page.getByRole('tab', { name: '工作表 2' });
  await expect(createdSheet).toHaveAttribute('aria-selected', 'true');
  await createdSheet.dblclick();

  const renameInput = page.getByRole('textbox', { name: '重命名工作表 2' });
  await renameInput.fill('Bad/Name');
  await renameInput.press('Enter');
  const renameError = page.getByRole('alert').filter({
    hasText: '名称不能包含 \\ / ? * [ ] :',
  });
  await expect(renameError).toBeVisible();
  await expect(renameInput).toHaveAttribute('aria-invalid', 'true');
  await expect(renameInput).toBeFocused();

  const geometry = await page
    .locator('.work-spreadsheet-footer')
    .evaluate((footer) => {
      const input = footer.querySelector<HTMLInputElement>(
        '.work-spreadsheet-sheet-tab.invalid > input',
      );
      const error = footer.querySelector<HTMLElement>(
        '.work-spreadsheet-sheet-rename-error',
      );
      const status = footer.querySelector<HTMLElement>(
        '.work-spreadsheet-status',
      );
      const tools = footer.querySelector<HTMLElement>(
        '.work-spreadsheet-sheet-tools',
      );
      if (!(input && error && status && tools)) {
        throw new Error('Phone worksheet rename feedback is incomplete.');
      }
      const footerBounds = footer.getBoundingClientRect();
      const inputBounds = input.getBoundingClientRect();
      const errorBounds = error.getBoundingClientRect();
      return {
        errorBottom: errorBounds.bottom,
        errorLeft: errorBounds.left,
        errorRight: errorBounds.right,
        footerBottom: footerBounds.bottom,
        footerHeight: footerBounds.height,
        inputLeft: inputBounds.left,
        inputRight: inputBounds.right,
        statusDisplay: getComputedStyle(status).display,
        toolsDisplay: getComputedStyle(tools).display,
        viewportHeight: document.documentElement.clientHeight,
        viewportWidth: document.documentElement.clientWidth,
      };
    });

  expect(geometry.footerHeight).toBeGreaterThanOrEqual(48);
  expect(geometry.statusDisplay).toBe('none');
  expect(geometry.toolsDisplay).toBe('none');
  expect(geometry.inputLeft).toBeGreaterThanOrEqual(0);
  expect(geometry.inputRight).toBeLessThanOrEqual(geometry.viewportWidth);
  expect(geometry.errorLeft).toBeGreaterThanOrEqual(0);
  expect(geometry.errorRight).toBeLessThanOrEqual(geometry.viewportWidth);
  expect(geometry.errorBottom).toBeLessThanOrEqual(geometry.footerBottom);
  expect(geometry.footerBottom).toBeLessThanOrEqual(geometry.viewportHeight);

  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-phone-invalid-name.png'),
    animations: 'disabled',
  });

  await renameInput.fill('移动端看板');
  await expect(renameError).toHaveCount(0);
  await renameInput.press('Enter');
  const renamedSheet = page.getByRole('tab', { name: '移动端看板' });
  await expect(renamedSheet).toBeFocused();

  await renamedSheet.dblclick();
  const cancelledInput = page.getByRole('textbox', {
    name: '重命名移动端看板',
  });
  await cancelledInput.fill('不应保存');
  await cancelledInput.press('Escape');
  await expect(page.getByRole('tab', { name: '移动端看板' })).toBeFocused();

  await page.keyboard.press('Shift+F10');
  const menu = page.getByRole('menu', { name: '移动端看板工作表操作' });
  await expect(menu).toBeVisible();
  await expect(menu.getByRole('menuitem').first()).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await expect(
    page.getByRole('button', { name: '移动端看板选项' }),
  ).toBeFocused();
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
