import { expect, type Page, test } from '@playwright/test';

test('Spreadsheet routes formatting and history shortcuts from the live grid', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  await grid.click();

  const shortcuts = [
    { button: page.getByRole('button', { name: '加粗' }), key: 'b' },
    { button: page.getByRole('button', { name: '斜体' }), key: 'i' },
    { button: page.getByRole('button', { name: '下划线' }), key: 'u' },
  ];
  for (const shortcut of shortcuts) {
    await expect(shortcut.button).toHaveAttribute('aria-pressed', 'false');
    await page.keyboard.press(`Control+${shortcut.key}`);
    await expect(shortcut.button).toHaveAttribute('aria-pressed', 'true');
    await expectGridFocus(page);
    await page.keyboard.press(`Control+${shortcut.key}`);
    await expect(shortcut.button).toHaveAttribute('aria-pressed', 'false');
  }

  const bold = shortcuts[0].button;
  await page.keyboard.press('Control+b');
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await page.keyboard.press('Control+z');
  await expect(bold).toHaveAttribute('aria-pressed', 'false');
  await page.keyboard.press('Control+y');
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await expectGridFocus(page);
});

test('Spreadsheet restores grid focus and navigates worksheets from the keyboard', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await page.getByRole('button', { name: '新建工作表' }).click();
  const createdSheet = page.getByRole('tab', { name: '工作表 2' });
  await expect(createdSheet).toHaveAttribute('aria-selected', 'true');
  await expectGridFocus(page);

  await page.keyboard.press('Control+PageUp');
  await expect(page.getByRole('tab', { name: '执行看板' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expectGridFocus(page);
});

async function openSpreadsheetFixture(page: Page) {
  await page.goto('/');
  await page
    .getByRole('button', {
      name: '季度执行计划 XLSX · 本次会话',
    })
    .click();
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();
}

async function expectGridFocus(page: Page) {
  await expect
    .poll(() =>
      page.evaluate(() =>
        document.activeElement?.classList.contains('fortune-sheet-overlay'),
      ),
    )
    .toBe(true);
}
