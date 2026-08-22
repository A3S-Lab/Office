import { expect, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

test('Spreadsheet keeps WPS font aliases and direct color resets aligned across ribbon densities', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const fontGroup = ribbon.getByRole('region', { name: '字体' });
  const bold = fontGroup.getByRole('button', { name: '加粗' });
  const italic = fontGroup.getByRole('button', { name: '斜体' });
  const underline = fontGroup.getByRole('button', {
    name: '下划线',
    exact: true,
  });
  const textColor = fontGroup.getByRole('button', { name: '文字颜色' });
  const fillColor = fontGroup.getByRole('button', { name: '填充颜色' });

  await expect(bold).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+B Meta+B Control+2',
  );
  await expect(italic).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+I Meta+I Control+3',
  );
  await expect(underline).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+U Meta+U Control+4',
  );

  await grid.focus();
  await page.keyboard.press('Shift+F11');
  await page.keyboard.type('A3S formatting');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');

  for (const [key, control] of [
    ['Control+2', bold],
    ['Control+3', italic],
    ['Control+4', underline],
  ] as const) {
    await expect(control).toHaveAttribute('aria-pressed', 'false');
    await page.keyboard.press(key);
    await expect(control).toHaveAttribute('aria-pressed', 'true');
    await expect(grid).toBeFocused();
  }

  await textColor.click();
  let colorDialog = page.getByRole('dialog', { name: '文字颜色' });
  await expect(
    colorDialog.getByRole('button', { name: '自动颜色' }),
  ).toBeVisible();
  await colorDialog.getByRole('option', { name: '颜色 #ff0000' }).click();
  await expect(textColor.locator('.work-office-color-swatch')).toHaveCSS(
    'background-color',
    'rgb(255, 0, 0)',
  );

  await textColor.click();
  colorDialog = page.getByRole('dialog', { name: '文字颜色' });
  await expect(
    colorDialog.getByRole('option', { name: '颜色 #ff0000' }),
  ).toHaveAttribute('aria-selected', 'true');
  await colorDialog.getByRole('button', { name: '自动颜色' }).click();
  await expect(grid).toBeFocused();
  await expect(textColor.locator('.work-office-color-swatch')).toHaveCSS(
    'background-color',
    'rgb(23, 32, 51)',
  );

  await ribbon.getByRole('button', { name: '撤销' }).click();
  await expect(textColor.locator('.work-office-color-swatch')).toHaveCSS(
    'background-color',
    'rgb(255, 0, 0)',
  );
  await textColor.click();
  await page
    .getByRole('dialog', { name: '文字颜色' })
    .getByRole('button', { name: '自动颜色' })
    .click();

  await fillColor.click();
  colorDialog = page.getByRole('dialog', { name: '填充颜色' });
  await expect(
    colorDialog.getByRole('button', { name: '无填充' }),
  ).toBeVisible();
  await colorDialog.getByRole('option', { name: '颜色 #ffff00' }).click();
  await expect(fillColor.locator('.work-office-color-swatch')).toHaveCSS(
    'background-color',
    'rgb(255, 255, 0)',
  );

  await fillColor.click();
  colorDialog = page.getByRole('dialog', { name: '填充颜色' });
  await expect(
    colorDialog.getByRole('option', { name: '颜色 #ffff00' }),
  ).toHaveAttribute('aria-selected', 'true');
  await colorDialog.getByRole('button', { name: '无填充' }).click();
  await expect(grid).toBeFocused();
  await expect(fillColor.locator('.work-office-color-swatch')).toHaveCSS(
    'background-color',
    'rgb(255, 255, 255)',
  );

  await ribbon.getByRole('button', { name: '撤销' }).click();
  await expect(fillColor.locator('.work-office-color-swatch')).toHaveCSS(
    'background-color',
    'rgb(255, 255, 0)',
  );
  await fillColor.click();
  colorDialog = page.getByRole('dialog', { name: '填充颜色' });
  await colorDialog.getByRole('button', { name: '无填充' }).click();
  await fillColor.click();
  colorDialog = page.getByRole('dialog', { name: '填充颜色' });
  await expect(
    colorDialog.getByRole('option', { name: '颜色 #ffff00' }),
  ).toHaveAttribute('aria-selected', 'false');
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-color-reset-menu.png'),
    animations: 'disabled',
  });
  await page.keyboard.press('Escape');

  expect(browserErrors).toEqual([]);
});
