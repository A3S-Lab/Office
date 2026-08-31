import { expect, type Locator, type Page, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

test('Spreadsheet sorts the complete table-owned range with structural controls locked', async ({
  page,
}, testInfo) => {
  const browserErrors = collectBrowserErrors(page);
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  await selectQuarterlyDataRange(page, grid, nameBox);
  await page.keyboard.press(
    process.platform === 'darwin' ? 'Meta+t' : 'Control+t',
  );
  const createTable = page.getByRole('dialog', { name: '创建表格' });
  await expect(
    createTable.getByRole('textbox', { name: '表格区域' }),
  ).toHaveValue('A3:G7');
  await createTable.getByRole('button', { name: '确定' }).click();

  await grid.focus();
  await page.keyboard.press('Control+Home');
  await move(page, 'ArrowDown', 3);
  await move(page, 'ArrowRight', 5);
  await expect(nameBox).toHaveText('F4');

  await ribbon.getByRole('tab', { name: '数据' }).click();
  const customSort = ribbon.getByRole('button', { name: '自定义排序' });
  await customSort.click();
  const warning = page.getByRole('dialog', { name: '排序提醒' });
  await expect(warning).toContainText('选定单元格位于表格中');
  await expect(warning).toContainText('A3:G7');
  await expect(warning).toContainText('F4');
  await expect(
    warning.getByRole('radio', { name: /对整个表格数据区域排序/ }),
  ).toBeChecked();
  await expect(
    warning.getByRole('radio', { name: /以当前选定区域排序/ }),
  ).toBeDisabled();
  await expectInsideViewport(page, warning);
  await warning.screenshot({
    path: testInfo.outputPath(
      `spreadsheet-table-owned-sort-warning-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await warning.getByRole('button', { name: '排序' }).click();

  const dialog = page.getByRole('dialog', { name: '自定义排序' });
  await expect(dialog).toContainText('执行看板!A3:G7');
  await expect(dialog).toContainText('当前范围由表格拥有');
  const header = dialog.getByRole('checkbox', { name: '数据包含标题' });
  await expect(header).toBeChecked();
  await expect(header).toBeDisabled();
  await assertStructuralSortOptions(page, dialog);
  await expect(
    dialog.getByRole('combobox', { name: '排序条件 1 列' }),
  ).toHaveValue('5');
  await dialog
    .getByRole('combobox', { name: '排序条件 1 次序' })
    .selectOption('descending');
  await expectInsideViewport(page, dialog);
  await dialog.screenshot({
    path: testInfo.outputPath(
      `spreadsheet-table-owned-sort-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '确定' }).click();
  await expect(customSort).toBeFocused();

  await grid.focus();
  await page.keyboard.press('Control+Home');
  await move(page, 'ArrowDown', 3);
  await move(page, 'ArrowRight', 5);
  await expect(nameBox).toHaveText('F4');
  await expect(formulaBar).toHaveText('=AVERAGE(C4:E4)');
  await page.keyboard.press('Control+z');
  await expect(formulaBar).toHaveText('=SUM(C4:E4)/3');
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet reapplies AutoFilter criteria after an owned-range sort and undo', async ({
  page,
}, testInfo) => {
  const browserErrors = collectBrowserErrors(page);
  await openSpreadsheetFixture(page);

  const editor = page.locator('.work-spreadsheet-editor');
  const canvas = page.locator('.work-spreadsheet-canvas');
  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  await grid.focus();
  await page.keyboard.press('Control+Home');
  await move(page, 'ArrowDown', 2);
  await expect(nameBox).toHaveText('A3');
  await ribbon.getByRole('tab', { name: '数据' }).click();
  const filter = ribbon.getByRole('button', { name: '自动筛选' });
  await filter.click();
  await expect(editor).toHaveAttribute('data-auto-filter', 'active');

  await grid.focus();
  await move(page, 'ArrowRight', 6);
  await expect(nameBox).toHaveText('G3');
  await page.keyboard.press('Alt+ArrowDown');
  const filterDialog = canvas.getByRole('dialog', { name: '状态 筛选' });
  await filterDialog.getByRole('button', { name: '清除', exact: true }).click();
  await filterDialog.getByRole('checkbox', { name: '显示 有风险' }).check();
  await filterDialog.getByRole('button', { name: /确\s*认/ }).click();
  await expect(canvas.getByRole('button', { name: '状态 筛选' })).toHaveClass(
    /luckysheet-filter-options-active/,
  );

  await expectVisibleRiskRow(page, canvas, nameBox, formulaBar, 'G5');
  await page.keyboard.press('ArrowLeft');
  await expect(nameBox).toHaveText('F5');
  await ribbon.getByRole('button', { name: '自定义排序' }).click();
  const warning = page.getByRole('dialog', { name: '排序提醒' });
  await expect(warning).toContainText('选定单元格位于筛选区域中');
  await expect(warning).toContainText('A3:G7');
  await expect(warning).toContainText('F5');
  await expect(
    warning.getByRole('radio', { name: /对整个筛选数据区域排序/ }),
  ).toBeChecked();
  await expect(
    warning.getByRole('radio', { name: /以当前选定区域排序/ }),
  ).toBeDisabled();
  await expectInsideViewport(page, warning);
  await warning.screenshot({
    path: testInfo.outputPath(
      `spreadsheet-auto-filter-owned-sort-warning-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await warning.getByRole('button', { name: '排序' }).click();

  const dialog = page.getByRole('dialog', { name: '自定义排序' });
  await expect(dialog).toContainText('执行看板!A3:G7');
  await expect(dialog).toContainText('当前范围由 AutoFilter 拥有');
  const header = dialog.getByRole('checkbox', { name: '数据包含标题' });
  await expect(header).toBeChecked();
  await expect(header).toBeDisabled();
  await assertStructuralSortOptions(page, dialog);
  await expect(
    dialog.getByRole('combobox', { name: '排序条件 1 列' }),
  ).toHaveValue('5');
  await dialog
    .getByRole('combobox', { name: '排序条件 1 次序' })
    .selectOption('descending');
  await dialog.screenshot({
    path: testInfo.outputPath(
      `spreadsheet-auto-filter-owned-sort-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '确定' }).click();

  await expectVisibleRiskRow(page, canvas, nameBox, formulaBar, 'G7');
  await move(page, 'ArrowLeft', 6);
  await expect(nameBox).toHaveText('A7');
  await expect(formulaBar).toHaveText('新版发布');

  await page.keyboard.press('Control+z');
  await expectVisibleRiskRow(page, canvas, nameBox, formulaBar, 'G5');
  await move(page, 'ArrowLeft', 6);
  await expect(nameBox).toHaveText('A5');
  await expect(formulaBar).toHaveText('新版发布');
  expect(browserErrors).toEqual([]);
});

async function selectQuarterlyDataRange(
  page: Page,
  grid: Locator,
  nameBox: Locator,
): Promise<void> {
  await grid.focus();
  await page.keyboard.press('Control+Home');
  await move(page, 'ArrowDown', 2);
  await move(page, 'Shift+ArrowRight', 6);
  await move(page, 'Shift+ArrowDown', 4);
  await expect(nameBox).toHaveText('A3:G7');
}

async function assertStructuralSortOptions(
  page: Page,
  dialog: Locator,
): Promise<void> {
  await dialog.getByRole('button', { name: '选项…' }).click();
  const options = page.getByRole('dialog', { name: '排序选项' });
  await expect(options).toContainText('结构化数据区域仅支持按列排序');
  await expect(options.getByRole('radio', { name: /按列排序/ })).toBeChecked();
  await expect(options.getByRole('radio', { name: /按行排序/ })).toBeDisabled();
  await expectInsideViewport(page, options);
  await options.getByRole('button', { name: '取消' }).click();
  await expect(options).toHaveCount(0);
}

async function expectVisibleRiskRow(
  page: Page,
  canvas: Locator,
  nameBox: Locator,
  formulaBar: Locator,
  expectedAddress: string,
): Promise<void> {
  const triggerBounds = await canvas
    .getByRole('button', { name: '状态 筛选' })
    .boundingBox();
  expect(triggerBounds).not.toBeNull();
  await page.mouse.click(
    (triggerBounds?.x ?? 0) + (triggerBounds?.width ?? 0) / 2,
    (triggerBounds?.y ?? 0) + (triggerBounds?.height ?? 0) + 12,
  );
  await expect(nameBox).toHaveText(expectedAddress);
  await expect(formulaBar).toHaveText('有风险');
}

function collectBrowserErrors(page: Page): string[] {
  const errors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });
  page.on('pageerror', (error) => errors.push(error.message));
  return errors;
}

async function move(page: Page, key: string, count: number): Promise<void> {
  for (let index = 0; index < count; index += 1) {
    await page.keyboard.press(key);
  }
}

async function expectInsideViewport(
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
