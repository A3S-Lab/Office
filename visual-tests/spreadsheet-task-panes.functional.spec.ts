import { expect, test, type Page } from '@playwright/test';

test('Spreadsheet closes a ribbon-opened workbook pane with Escape', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await page.getByRole('tab', { name: '插入', exact: true }).click();
  await page.getByRole('button', { name: '插入图表' }).click();
  const chartPane = page.getByRole('region', { name: '图表管理器' });
  await expect(chartPane).toBeVisible();

  await page.keyboard.press('Escape');
  await expect(chartPane).toBeHidden();
});

test('Spreadsheet cancels a dirty chart draft before closing its pane', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await page.getByRole('tab', { name: '插入', exact: true }).click();
  await page.getByRole('button', { name: '插入图表' }).click();
  const chartPane = page.getByRole('region', { name: '图表管理器' });
  await chartPane.getByRole('button', { name: '根据当前选区新建' }).click();
  const chartName = chartPane.getByRole('textbox', { name: '图表对象名称' });
  const savedName = await chartName.inputValue();
  await chartName.fill('尚未保存的图表');

  await page.keyboard.press('Escape');
  await expect(chartPane).toBeVisible();
  await expect(chartName).toHaveValue(savedName);
  await expect(
    chartPane.getByRole('button', { name: '取消更改' }),
  ).toBeDisabled();

  await page.keyboard.press('Escape');
  await expect(chartPane).toBeHidden();
});

test('Spreadsheet cancels dirty conditional formatting before closing its pane', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await page.getByRole('tab', { name: '插入', exact: true }).click();
  await page.getByRole('button', { name: '条件格式' }).click();
  const conditionalPane = page.getByRole('region', {
    name: '条件格式管理器',
  });
  const reference = conditionalPane.getByRole('textbox', {
    name: '条件格式范围',
  });
  await expect(reference).toHaveValue('A1:A10');
  await reference.fill('A1:A20');

  await page.keyboard.press('Escape');
  await expect(conditionalPane).toBeVisible();
  await expect(reference).toHaveValue('A1:A10');

  await page.keyboard.press('Escape');
  await expect(conditionalPane).toBeHidden();
});

test('Spreadsheet cancels dirty calculation settings before closing its pane', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await page.getByRole('tab', { name: '公式', exact: true }).click();
  await page.getByRole('button', { name: '公式与计算' }).click();
  const formulaPane = page.getByRole('region', {
    name: '公式与计算',
    exact: true,
  });
  const fullPrecision = formulaPane.getByRole('checkbox', {
    name: '使用完整精度',
  });
  await expect(fullPrecision).toBeChecked();
  await fullPrecision.click();
  await expect(fullPrecision).not.toBeChecked();

  await page.keyboard.press('Escape');
  await expect(formulaPane).toBeVisible();
  await expect(fullPrecision).toBeChecked();
  await expect(
    formulaPane.getByRole('button', { name: '保存计算设置' }),
  ).toBeDisabled();

  await page.keyboard.press('Escape');
  await expect(formulaPane).toBeHidden();
});

test('Spreadsheet cancels a dirty defined name before closing its pane', async ({
  page,
}) => {
  await expectTextDraftBeforePaneClose(page, {
    tab: '公式',
    button: '名称管理器',
    region: '名称管理器',
    field: '名称',
    initial: '',
    edited: 'RevenueDraft',
  });
});

test('Spreadsheet cancels dirty print settings before closing its pane', async ({
  page,
}) => {
  await expectTextDraftBeforePaneClose(page, {
    tab: '页面布局',
    button: '打印设置',
    region: '打印设置',
    field: '打印范围',
    initial: '',
    edited: 'A1:G20',
  });
});

test('Spreadsheet cancels a dirty editable range before closing its pane', async ({
  page,
}) => {
  await expectTextDraftBeforePaneClose(page, {
    tab: '审阅',
    button: '工作表保护',
    region: '工作表保护',
    field: '可编辑区域名称',
    initial: '',
    edited: 'InputCells',
  });
});

async function expectTextDraftBeforePaneClose(
  page: Page,
  options: {
    tab: string;
    button: string;
    region: string;
    field: string;
    initial: string;
    edited: string;
  },
) {
  await openSpreadsheetFixture(page);
  await page.getByRole('tab', { name: options.tab, exact: true }).click();
  await page.getByRole('button', { name: options.button }).click();
  const pane = page.getByRole('region', {
    name: options.region,
    exact: true,
  });
  const field = pane.getByRole('textbox', {
    name: options.field,
    exact: true,
  });
  await expect(field).toHaveValue(options.initial);
  await field.fill(options.edited);

  await page.keyboard.press('Escape');
  await expect(pane).toBeVisible();
  await expect(field).toHaveValue(options.initial);

  await page.keyboard.press('Escape');
  await expect(pane).toBeHidden();
}

async function openSpreadsheetFixture(page: Page) {
  await page.goto('/');
  await page
    .getByRole('button', {
      name: '季度执行计划 XLSX · 本次会话',
    })
    .click();
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();
}
