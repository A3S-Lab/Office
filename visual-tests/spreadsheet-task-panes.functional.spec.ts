import { expect, test, type Page } from '@playwright/test';

test('Spreadsheet closes a ribbon-opened workbook pane with Escape', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await page.getByRole('tab', { name: '插入', exact: true }).click();
  const trigger = page.getByRole('button', { name: '插入图表' });
  await trigger.click();
  const chartPane = spreadsheetWorkbookPane(page, '图表管理器');
  await expect(chartPane).toBeVisible();
  await expect(trigger).toHaveAttribute('aria-expanded', 'true');

  await page.keyboard.press('Escape');
  await expect(chartPane).toBeHidden();
  await expect(trigger).toBeFocused();
});

test('Spreadsheet keeps workbook tools in a bounded right task pane', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await page.getByRole('tab', { name: '数据', exact: true }).click();
  const trigger = page.getByRole('button', { name: '数据透视表' });
  await trigger.click();
  const pane = spreadsheetWorkbookPane(page, '数据透视表管理器');
  await expect(pane).toBeVisible();
  await expect(
    pane.getByRole('button', { name: '关闭数据透视表' }),
  ).toBeVisible();

  const geometry = await page
    .locator('.work-spreadsheet-workspace')
    .evaluate((workspace) => {
      const canvas = workspace.querySelector<HTMLElement>(
        '.work-spreadsheet-canvas',
      );
      const pane = workspace.querySelector<HTMLElement>(
        '.work-spreadsheet-workbook-panel',
      );
      const paneBody = workspace.querySelector<HTMLElement>(
        '.work-spreadsheet-workbook-panel-body',
      );
      if (!(canvas && pane && paneBody)) {
        throw new Error('Spreadsheet task-pane geometry is incomplete.');
      }
      const workspaceRect = workspace.getBoundingClientRect();
      const canvasRect = canvas.getBoundingClientRect();
      const paneRect = pane.getBoundingClientRect();
      return {
        canvasBottom: canvasRect.bottom,
        canvasHeight: canvasRect.height,
        canvasRight: canvasRect.right,
        canvasTop: canvasRect.top,
        paneBodyClientWidth: paneBody.clientWidth,
        paneBodyScrollWidth: paneBody.scrollWidth,
        paneBottom: paneRect.bottom,
        paneLeft: paneRect.left,
        paneRight: paneRect.right,
        paneTop: paneRect.top,
        paneWidth: paneRect.width,
        position: getComputedStyle(pane).position,
        workspaceRight: workspaceRect.right,
        workspaceWidth: workspaceRect.width,
      };
    });

  expect(geometry.paneTop).toBeCloseTo(geometry.canvasTop, 0);
  expect(geometry.paneBottom).toBeCloseTo(geometry.canvasBottom, 0);
  expect(geometry.paneRight).toBeCloseTo(geometry.workspaceRight, 0);
  expect(geometry.paneBodyScrollWidth).toBeLessThanOrEqual(
    geometry.paneBodyClientWidth + 1,
  );
  expect(geometry.canvasHeight).toBeGreaterThan(420);

  if (geometry.position === 'absolute') {
    expect(geometry.canvasRight).toBeCloseTo(geometry.workspaceRight, 0);
    expect(geometry.paneWidth).toBeLessThan(geometry.workspaceWidth);
  } else {
    expect(geometry.canvasRight).toBeCloseTo(geometry.paneLeft, 0);
    expect(geometry.paneWidth).toBeGreaterThanOrEqual(320);
    expect(geometry.paneWidth).toBeLessThanOrEqual(460);
  }
});

test('Spreadsheet task pane fills the phone workspace without page overflow', async ({
  page,
}, testInfo) => {
  test.skip(
    testInfo.project.name !== 'desktop-1280',
    'The phone contract only needs one browser project.',
  );
  await page.setViewportSize({ width: 390, height: 700 });
  await openSpreadsheetFixture(page);

  await page.getByRole('tab', { name: '数据', exact: true }).click();
  const trigger = page.getByRole('button', { name: '数据透视表' });
  await trigger.click();

  const pane = page.getByRole('dialog', { name: '数据透视表管理器' });
  const close = pane.getByRole('button', { name: '关闭数据透视表' });
  const create = pane.getByRole('button', { name: '根据当前选区新建' });
  await expect(pane).toHaveAttribute('aria-modal', 'true');
  await expect(page.locator('.work-spreadsheet-ribbon')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(page.locator('.work-spreadsheet-canvas')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(page.locator('.work-spreadsheet-footer')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(close).toBeFocused();

  await close.press('Tab');
  await expect(create).toBeFocused();
  await create.press('Shift+Tab');
  await expect(close).toBeFocused();

  const geometry = await page
    .locator('.work-spreadsheet-workspace')
    .evaluate((workspace) => {
      const pane = workspace.querySelector<HTMLElement>(
        '.work-spreadsheet-workbook-panel',
      );
      const paneBody = workspace.querySelector<HTMLElement>(
        '.work-spreadsheet-workbook-panel-body',
      );
      if (!(pane && paneBody)) {
        throw new Error('Phone spreadsheet task pane is unavailable.');
      }
      const workspaceRect = workspace.getBoundingClientRect();
      const paneRect = pane.getBoundingClientRect();
      return {
        documentScrollWidth: document.documentElement.scrollWidth,
        paneBodyClientWidth: paneBody.clientWidth,
        paneBodyScrollWidth: paneBody.scrollWidth,
        paneLeft: paneRect.left,
        paneRight: paneRect.right,
        workspaceLeft: workspaceRect.left,
        workspaceRight: workspaceRect.right,
        viewportWidth: document.documentElement.clientWidth,
      };
    });

  expect(geometry.documentScrollWidth).toBeLessThanOrEqual(
    geometry.viewportWidth + 1,
  );
  expect(geometry.paneLeft).toBeCloseTo(geometry.workspaceLeft, 0);
  expect(geometry.paneRight).toBeCloseTo(geometry.workspaceRight, 0);
  expect(geometry.paneBodyScrollWidth).toBeLessThanOrEqual(
    geometry.paneBodyClientWidth + 1,
  );

  await close.press('Escape');
  await expect(pane).toBeHidden();
  await expect(trigger).toBeFocused();
  await expect(page.locator('.work-spreadsheet-ribbon')).not.toHaveAttribute(
    'inert',
    '',
  );
});

test('Spreadsheet cancels a dirty chart draft before closing its pane', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await page.getByRole('tab', { name: '插入', exact: true }).click();
  await page.getByRole('button', { name: '插入图表' }).click();
  const chartPane = spreadsheetWorkbookPane(page, '图表管理器');
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

  await page.getByRole('tab', { name: '开始', exact: true }).click();
  await page.getByRole('button', { name: '条件格式' }).click();
  const conditionalPane = spreadsheetWorkbookPane(page, '条件格式管理器');
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
  const formulaPane = spreadsheetWorkbookPane(page, '公式与计算');
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
  const pane = spreadsheetWorkbookPane(page, options.region);
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
  await page.goto('/playground/');
  await page
    .getByRole('button', {
      name: '季度执行计划 XLSX · 本次会话',
    })
    .click();
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();
}

function spreadsheetWorkbookPane(page: Page, name: string) {
  return page.locator(`.work-spreadsheet-workbook-panel[aria-label="${name}"]`);
}
