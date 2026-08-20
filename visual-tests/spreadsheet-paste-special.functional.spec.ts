import { expect, test } from '@playwright/test';

test('Spreadsheet Paste Special keeps rich content atomic at every layout', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.goto('/?e2e=spreadsheet-paste-special');
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const clipboard = ribbon.getByRole('region', { name: '剪贴板' });
  const primaryPaste = clipboard.getByRole('button', {
    name: '粘贴',
    exact: true,
  });
  const disclosure = clipboard.getByRole('button', {
    name: '更多粘贴方式',
  });

  await expect(primaryPaste).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+V Meta+V',
  );
  await expect(disclosure).toHaveAttribute('aria-haspopup', 'menu');
  await grid.focus();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Shift+ArrowRight');
  await expect(nameBox).toHaveText('A1:B1');
  await page.keyboard.press('Control+c');

  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ArrowRight');
  await expect(nameBox).toHaveText('C3');
  await page.keyboard.press('Control+Alt+v');

  const dialog = page.getByRole('dialog', { name: '选择性粘贴' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('radio')).toHaveCount(15);
  await expect(
    dialog.getByRole('radio', { name: '格式', exact: true }),
  ).toBeEnabled();
  await expect(dialog.getByLabel('剪贴板摘要')).toContainText('A3S 富剪贴板');
  await expectDialogInsideViewport(page, dialog);

  await dialog.getByRole('radio', { name: '值', exact: true }).click();
  await dialog.getByRole('radio', { name: '加', exact: true }).click();
  await dialog.getByRole('checkbox', { name: '跳过空白单元格' }).click();
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-paste-special-dialog.png'),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '粘贴' }).click();

  await expect(dialog).toHaveCount(0);
  await expect(grid).toBeFocused();
  await expect(nameBox).toHaveText('C3:D3');
  await expect(formulaBar).toHaveText('110');
  await page.keyboard.press('ArrowRight');
  await expect(formulaBar).toHaveText('220');

  await page.keyboard.press('Control+z');
  await expect(formulaBar).toHaveText('200');
  await page.keyboard.press('ArrowLeft');
  await expect(formulaBar).toHaveText('100');

  await disclosure.click();
  const menu = page.getByRole('menu', { name: '粘贴选项' });
  await expect(menu.getByRole('menuitem')).toHaveCount(5);
  await expect(menu.getByRole('menuitem', { name: '全部' })).toBeFocused();
  await expect(
    menu.getByRole('menuitem', { name: /选择性粘贴/ }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+Alt+V Meta+Alt+V');
  await menu.getByRole('menuitem', { name: '公式', exact: true }).click();

  await expect(menu).toHaveCount(0);
  await expect(grid).toBeFocused();
  await expect(nameBox).toHaveText('C3:D3');
  await expect(formulaBar).toHaveText('10');
  await page.keyboard.press('ArrowRight');
  await expect(formulaBar).toHaveText('=$A3+C$1');
  await page.keyboard.press('Control+z');
  await expect(formulaBar).toHaveText('200');
  expect(browserErrors).toEqual([]);
});

async function expectDialogInsideViewport(
  page: import('@playwright/test').Page,
  dialog: import('@playwright/test').Locator,
) {
  const bounds = await dialog.boundingBox();
  const viewport = page.viewportSize();
  expect(bounds).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect(bounds?.x ?? -1).toBeGreaterThanOrEqual(0);
  expect(bounds?.y ?? -1).toBeGreaterThanOrEqual(0);
  expect((bounds?.x ?? 0) + (bounds?.width ?? 0)).toBeLessThanOrEqual(
    viewport?.width ?? 0,
  );
  expect((bounds?.y ?? 0) + (bounds?.height ?? 0)).toBeLessThanOrEqual(
    viewport?.height ?? 0,
  );
}
