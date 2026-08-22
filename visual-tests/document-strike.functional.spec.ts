import { expect, test } from '@playwright/test';

test('Writer preserves native double strike, explicit reset, shortcut ownership, and Undo', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.goto('/');
  await page
    .getByRole('button', {
      name: '空白文字 从一张干净的 A4 页面开始',
    })
    .click();

  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  await editor.fill('Native strike fidelity');
  await editor.click();
  await page.keyboard.press(`${modifier}+a`);

  const fontGroup = page.getByRole('region', { name: '字体' });
  const strike = fontGroup.getByRole('button', {
    name: '删除线',
    exact: true,
  });
  await expect(strike).not.toHaveAttribute('aria-keyshortcuts', /.+/);
  await fontGroup.getByRole('button', { name: '更多删除线' }).click();

  const menu = page.getByRole('menu', { name: '删除线样式' });
  await expect(menu).toBeVisible();
  await expect(menu.getByRole('menuitemradio')).toHaveCount(3);
  await expect(
    menu.getByRole('menuitemradio', { name: '无删除线' }),
  ).toBeFocused();
  await page.screenshot({
    path: testInfo.outputPath('writer-strike-menu.png'),
    animations: 'disabled',
  });
  await menu.getByRole('menuitemradio', { name: '双删除线' }).click();

  const doubleStrike = editor.locator('s[data-office-strike-style="double"]');
  await expect(doubleStrike).toHaveText('Native strike fidelity');
  await expect(doubleStrike).toHaveCSS('text-decoration-style', 'double');
  await expect(strike).toHaveAttribute('title', '删除线（双删除线）');
  await page.keyboard.press(`${modifier}+Shift+s`);
  await expect(doubleStrike).toHaveCount(1);

  await strike.click();
  const explicitNone = editor.locator('s[data-office-strike-style="none"]');
  await expect(explicitNone).toHaveText('Native strike fidelity');
  await expect(explicitNone).toHaveCSS('text-decoration-line', 'none');
  await page.keyboard.press(`${modifier}+z`);
  await expect(doubleStrike).toHaveCount(1);
  await expect(editor).toHaveAttribute('data-pagination-text-engine', 'wasm');
  expect(browserErrors).toEqual([]);
});
