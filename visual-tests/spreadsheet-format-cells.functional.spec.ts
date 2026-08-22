import { expect, type Locator, type Page, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

test('Spreadsheet Format Cells applies six tabs and undoes once at every layout', async ({
  page,
}) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  await grid.focus();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ArrowRight');
  await expect(page.locator('.fortune-name-box')).toHaveText('C4');

  const dialog = await openFormatCells(page, grid);
  await expectDialogInsideViewport(page, dialog);
  await expect(dialog.getByRole('tab')).toHaveCount(6);

  const numberCode = dialog.getByRole('textbox', {
    name: '数字格式代码',
  });
  const initialNumberCode = await numberCode.inputValue();
  const targetNumberCode = initialNumberCode === '0.000' ? '0.0000' : '0.000';
  await numberCode.fill(targetNumberCode);

  await dialog.getByRole('tab', { name: '对齐' }).click();
  const horizontal = dialog.getByRole('combobox', { name: '水平对齐' });
  const initialHorizontal = (await horizontal.textContent())?.trim() ?? '常规';
  const targetHorizontal = initialHorizontal === '居中' ? '右对齐' : '居中';
  await horizontal.click();
  await page
    .getByRole('listbox', { name: '水平对齐' })
    .getByRole('option', { name: targetHorizontal })
    .click();
  const wrap = dialog.getByRole('checkbox', { name: '自动换行' });
  const initialWrap = await wrap.isChecked();
  await wrap.click();

  await dialog.getByRole('tab', { name: '字体' }).click();
  const bold = dialog.getByRole('checkbox', { name: '加粗' });
  const initialBold = await bold.isChecked();
  await bold.click();

  await dialog.getByRole('tab', { name: '边框' }).click();
  const bottomBorder = dialog.getByRole('button', {
    name: '下框线',
    exact: true,
  });
  const initialBottomBorder =
    (await bottomBorder.getAttribute('aria-pressed')) === 'true';
  await bottomBorder.click();

  await dialog.getByRole('tab', { name: '填充' }).click();
  const fill = dialog.getByRole('button', { name: '单元格填充颜色' });
  const initialFill = (await fill.textContent())?.trim() ?? '#FFFFFF';
  const targetFill = initialFill.toLowerCase().includes('#fff2cc')
    ? '#d9ead3'
    : '#fff2cc';
  await fill.click();
  await page
    .getByRole('listbox', { name: '颜色' })
    .getByRole('option', { name: `颜色 ${targetFill}` })
    .click();

  await dialog.getByRole('tab', { name: '保护' }).click();
  const hidden = dialog.getByRole('checkbox', { name: '隐藏公式' });
  const initialHidden = await hidden.isChecked();
  await hidden.click();
  await dialog.getByRole('button', { name: '应用' }).click();

  await expect(dialog).toHaveCount(0);
  await expect(grid).toBeFocused();

  const applied = await openFormatCells(page, grid);
  await expect(
    applied.getByRole('textbox', { name: '数字格式代码' }),
  ).toHaveValue(targetNumberCode);
  await applied.getByRole('tab', { name: '对齐' }).click();
  await expect(
    applied.getByRole('combobox', { name: '水平对齐' }),
  ).toContainText(targetHorizontal);
  await expectChecked(
    applied.getByRole('checkbox', { name: '自动换行' }),
    !initialWrap,
  );
  await applied.getByRole('tab', { name: '字体' }).click();
  await expectChecked(
    applied.getByRole('checkbox', { name: '加粗' }),
    !initialBold,
  );
  await applied.getByRole('tab', { name: '边框' }).click();
  await expect(
    applied.getByRole('button', { name: '下框线', exact: true }),
  ).toHaveAttribute('aria-pressed', String(!initialBottomBorder));
  await applied.getByRole('tab', { name: '填充' }).click();
  await expect(
    applied.getByRole('button', { name: '单元格填充颜色' }),
  ).toContainText(targetFill.toUpperCase());
  await applied.getByRole('tab', { name: '保护' }).click();
  await expectChecked(
    applied.getByRole('checkbox', { name: '隐藏公式' }),
    !initialHidden,
  );
  await page.keyboard.press('Escape');

  await expect(applied).toHaveCount(0);
  await expect(grid).toBeFocused();
  await page.keyboard.press('Control+z');

  const restored = await openFormatCells(page, grid);
  await expect(
    restored.getByRole('textbox', { name: '数字格式代码' }),
  ).toHaveValue(initialNumberCode);
  await restored.getByRole('tab', { name: '对齐' }).click();
  await expect(
    restored.getByRole('combobox', { name: '水平对齐' }),
  ).toContainText(initialHorizontal);
  await expectChecked(
    restored.getByRole('checkbox', { name: '自动换行' }),
    initialWrap,
  );
  await restored.getByRole('tab', { name: '字体' }).click();
  await expectChecked(
    restored.getByRole('checkbox', { name: '加粗' }),
    initialBold,
  );
  await restored.getByRole('tab', { name: '边框' }).click();
  await expect(
    restored.getByRole('button', { name: '下框线', exact: true }),
  ).toHaveAttribute('aria-pressed', String(initialBottomBorder));
  await restored.getByRole('tab', { name: '填充' }).click();
  await expect(
    restored.getByRole('button', { name: '单元格填充颜色' }),
  ).toContainText(initialFill);
  await restored.getByRole('tab', { name: '保护' }).click();
  await expectChecked(
    restored.getByRole('checkbox', { name: '隐藏公式' }),
    initialHidden,
  );
  await page.keyboard.press('Escape');

  await expect(restored).toHaveCount(0);
  await expect(grid).toBeFocused();
  expect(browserErrors).toEqual([]);
});

async function openFormatCells(page: Page, grid: Locator): Promise<Locator> {
  await grid.focus();
  await page.keyboard.press('Control+1');
  const dialog = page.getByRole('dialog', { name: '设置单元格格式' });
  await expect(dialog).toBeVisible();
  return dialog;
}

async function expectChecked(
  locator: Locator,
  checked: boolean,
): Promise<void> {
  if (checked) await expect(locator).toBeChecked();
  else await expect(locator).not.toBeChecked();
}

async function expectDialogInsideViewport(
  page: Page,
  dialog: Locator,
): Promise<void> {
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
