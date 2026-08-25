import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  stabilizeVisualSurface,
  waitForDocumentFixture,
} from './visual-test-support';

test('document table Design and Layout stay visual and survive preview', async ({
  page,
}) => {
  await page.goto('/playground/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  await page.locator('.work-document-editable .ProseMirror p').first().click();
  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '插入表格' }).click();
  const picker = page.getByRole('dialog', { name: '选择表格大小' });
  await expect(picker.getByRole('button', { name: '1 行 1 列' })).toBeFocused();
  await picker.getByRole('button', { name: '3 行 3 列' }).click();

  const designTab = page.getByRole('tab', { name: '表格设计' });
  const layoutTab = page.getByRole('tab', { name: '表格布局' });
  await expect(designTab).toHaveAttribute('aria-selected', 'true');
  await expect(layoutTab).toBeVisible();
  await page.getByRole('radio', { name: '应用表格样式：蓝色条纹' }).click();

  const table = page
    .locator('.work-document-editable .ProseMirror table')
    .first();
  const header = table.locator('th').first();
  await expect(table.locator(':is(th, td)')).toHaveCount(9);
  await expect(header).toHaveCSS('background-color', 'rgb(217, 234, 247)');
  await expect(header).toHaveCSS('border-top-color', 'rgb(159, 186, 208)');

  await layoutTab.click();
  await page.getByRole('button', { name: '单元格垂直居中' }).click();
  await page.getByRole('button', { name: '单元格水平居中' }).click();
  await expect(header).toHaveCSS('vertical-align', 'middle');
  await expect(header.locator('p')).toHaveCSS('text-align', 'center');

  await designTab.click();
  await table.scrollIntoViewIfNeeded();
  await stabilizeVisualSurface(page);
  await expect(page).toHaveScreenshot('document-table-design.png');

  await page.getByRole('button', { name: '预览' }).click();
  const previewHeader = page
    .locator('.work-document-preview-page table th')
    .first();
  await expect(previewHeader).toBeVisible();
  await expect(previewHeader).toHaveCSS(
    'background-color',
    'rgb(217, 234, 247)',
  );
  await expect(previewHeader).toHaveCSS('vertical-align', 'middle');
  await expect(previewHeader).toHaveCSS(
    'border-top-color',
    'rgb(159, 186, 208)',
  );
});

test('table sizing uses contained touch controls on phone', async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/playground/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  await page.locator('.work-document-editable .ProseMirror p').first().click();
  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '插入表格' }).click();

  const picker = page.getByRole('dialog', { name: '选择表格大小' });
  const rows = picker.getByRole('spinbutton', { name: '行数' });
  const columns = picker.getByRole('spinbutton', { name: '列数' });
  await expect(rows).toBeFocused();
  await expect(picker.getByRole('button', { name: '1 行 1 列' })).toHaveCount(
    0,
  );

  const decrement = picker.getByRole('button', { name: '减少行数' });
  const decrementBounds = await decrement.boundingBox();
  expect(decrementBounds?.width).toBeGreaterThanOrEqual(40);
  expect(decrementBounds?.height).toBeGreaterThanOrEqual(40);
  const pickerBounds = await picker.boundingBox();
  expect(pickerBounds?.x).toBeGreaterThanOrEqual(0);
  expect(
    (pickerBounds?.x ?? 0) + (pickerBounds?.width ?? 0),
  ).toBeLessThanOrEqual(390);

  const submit = picker.getByRole('button', {
    name: '插入 1 × 1 表格',
  });
  await expect(submit).toHaveCSS('background-color', 'rgb(40, 100, 232)');
  await expect(submit).toHaveCSS('color', 'rgb(255, 255, 255)');

  await rows.fill('3');
  await columns.fill('3');
  await picker.getByRole('button', { name: '插入 3 × 3 表格' }).click();

  await expect(
    page.locator('.work-document-editable .ProseMirror table :is(th, td)'),
  ).toHaveCount(9);
  await expect(page.getByRole('tab', { name: '表格设计' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
});

test('table sizing stays usable on desktop and compact ribbons', async ({
  page,
}) => {
  await page.goto('/playground/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  await page.locator('.work-document-editable .ProseMirror p').first().click();
  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '插入表格' }).click();
  await page
    .getByRole('dialog', { name: '选择表格大小' })
    .getByRole('button', { name: '3 行 3 列' })
    .click();
  await page.getByRole('tab', { name: '表格布局' }).click();

  const width = page.getByRole('textbox', { name: '列宽（厘米）' });
  await revealRibbonControl(page, '表格布局', width);
  const initialWidth = await width.inputValue();
  await width.fill('8.5');
  await expect(width).toHaveAttribute('data-office-escape-consumer', 'true');
  await width.press('Escape');
  await expect(width).toHaveValue(initialWidth);
  await expect(width).not.toHaveAttribute('data-office-escape-consumer');
  await expect(width).toBeFocused();

  await width.fill('3.2');
  await width.press('Enter');

  const table = page
    .locator('.work-document-editable .ProseMirror table')
    .first();
  await expect(table).toHaveAttribute('data-office-table-layout', 'fixed');
  await expect
    .poll(async () =>
      Number.parseFloat(
        await table
          .locator('colgroup > col')
          .first()
          .evaluate(
            (column) =>
              (column as HTMLElement).style.width ||
              getComputedStyle(column).width,
          ),
      ),
    )
    .toBeGreaterThan(119);
  await expect
    .poll(async () =>
      Number.parseFloat(
        await table.evaluate((node) => getComputedStyle(node).width),
      ),
    )
    .toBeGreaterThan(450);

  const distributeRows = page.getByRole('button', {
    name: '平均分布行',
  });
  await revealRibbonControl(page, '表格布局', distributeRows);
  await distributeRows.click();
  await expect(table.locator('tr').first()).toHaveAttribute(
    'data-office-row-height',
    /\d+/,
  );

  const autofit = page.getByRole('combobox', { name: '表格自动调整' });
  await revealRibbonControl(page, '表格布局', autofit);
  await autofit.click();
  await page.getByRole('option', { name: '适应内容' }).click();
  await expect(table).toHaveAttribute('data-office-table-layout', 'autofit');
  await expect(table).toHaveAttribute('data-office-table-width-type', 'auto');
  await expect
    .poll(() =>
      table
        .locator('colgroup > col')
        .first()
        .evaluate((column) => (column as HTMLElement).style.width),
    )
    .toBe('');
});

test('table properties keep table, row, column, and cell changes in one workflow', async ({
  page,
}) => {
  const pageErrors: Error[] = [];
  page.on('pageerror', (error) => pageErrors.push(error));
  await page.goto('/playground/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  await page.locator('.work-document-editable .ProseMirror p').first().click();
  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '插入表格' }).click();
  await page
    .getByRole('dialog', { name: '选择表格大小' })
    .getByRole('button', { name: '2 行 2 列' })
    .click();
  await page.getByRole('tab', { name: '表格布局' }).click();

  const properties = page.getByRole('button', { name: '表格属性' });
  await revealRibbonControl(page, '表格布局', properties);
  await properties.click();
  const dialog = page.getByRole('dialog', { name: '表格属性' });
  await expect(dialog).toBeVisible();
  await expect(
    dialog.getByRole('tablist', { name: '表格属性分类' }),
  ).toBeVisible();
  await expect(dialog.getByRole('tab')).toHaveCount(4);

  await page.setViewportSize({ width: 390, height: 844 });
  const dialogBounds = await dialog.boundingBox();
  expect(dialogBounds?.x).toBeGreaterThanOrEqual(0);
  expect(
    (dialogBounds?.x ?? 0) + (dialogBounds?.width ?? 0),
  ).toBeLessThanOrEqual(390);
  for (const name of ['表格', '行', '列', '单元格']) {
    const bounds = await dialog
      .getByRole('tab', { name, exact: true })
      .boundingBox();
    expect(bounds?.height).toBeGreaterThanOrEqual(40);
  }

  await dialog.getByRole('tab', { name: '表格', exact: true }).click();
  await dialog.getByRole('textbox', { name: '表格宽度（百分比）' }).fill('80');

  await dialog.getByRole('tab', { name: '行', exact: true }).click();
  await dialog.getByRole('checkbox', { name: '指定行高' }).check();
  await dialog.getByRole('textbox', { name: '当前行高（厘米）' }).fill('1.2');
  await dialog.getByRole('combobox', { name: '行高规则' }).click();
  await page.getByRole('option', { name: '固定值' }).click();
  await dialog.getByRole('checkbox', { name: '允许跨页断行' }).uncheck();
  await dialog
    .getByRole('checkbox', { name: '在各页顶端重复标题行' })
    .uncheck();

  await dialog.getByRole('tab', { name: '列', exact: true }).click();
  await dialog.getByRole('textbox', { name: '当前列宽（厘米）' }).fill('4');

  await dialog.getByRole('tab', { name: '单元格', exact: true }).click();
  const verticalCenter = dialog.getByRole('radio', { name: '居中' });
  await dialog
    .locator(
      'label:has(input[name="table-properties-vertical-alignment"][value="middle"])',
    )
    .click();
  await expect(verticalCenter).toBeChecked();
  await dialog.getByRole('checkbox', { name: '使用表格默认边距' }).uncheck();
  const leftMargin = dialog.getByRole('textbox', {
    name: '当前单元格左边距（厘米）',
  });
  await leftMargin.fill('0.4');
  const leftMarginBounds = await leftMargin.boundingBox();
  expect(leftMarginBounds?.height).toBeGreaterThanOrEqual(40);

  await dialog.getByRole('button', { name: '确定' }).click();

  const table = page.locator('.work-document-editable table').first();
  const firstRow = table.locator('tr').first();
  const firstCell = firstRow.locator(':is(th, td)').first();
  await expect(table).toHaveAttribute('data-office-table-width', '80');
  await expect(firstRow).toHaveAttribute('data-office-row-height', '45.35');
  await expect(firstRow).toHaveAttribute(
    'data-office-row-height-rule',
    'exact',
  );
  await expect(firstRow).toHaveAttribute('data-office-cant-split', 'true');
  await expect(firstRow).toHaveAttribute('data-office-repeat-header', 'false');
  await expect(firstCell).toHaveAttribute(
    'data-office-cell-vertical-align',
    'middle',
  );
  await expect(firstCell).toHaveAttribute(
    'data-office-cell-margin-left',
    '15.12',
  );
  await expect
    .poll(async () =>
      Number.parseFloat(
        await table
          .locator('colgroup > col')
          .first()
          .evaluate(
            (column) =>
              (column as HTMLElement).style.width ||
              getComputedStyle(column).width,
          ),
      ),
    )
    .toBeCloseTo(151.18, 1);

  await page.keyboard.press('Control+z');
  await expect(table).toHaveAttribute('data-office-table-width', '100');
  await expect(firstRow).not.toHaveAttribute('data-office-row-height');
  await expect(firstRow).not.toHaveAttribute('data-office-cant-split');
  await expect(firstCell).toHaveAttribute(
    'data-office-cell-vertical-align',
    'top',
  );
  await expect(firstCell).not.toHaveAttribute('data-office-cell-margin-left');
  expect(pageErrors).toEqual([]);
});

test('table contextual tabs clean up after deletion through real ribbon overflow', async ({
  page,
}) => {
  await page.goto('/playground/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  const body = page.getByRole('textbox', { name: '文档正文' });
  await page.locator('.work-document-editable .ProseMirror p').first().click();
  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '插入表格' }).click();
  await page
    .getByRole('dialog', { name: '选择表格大小' })
    .getByRole('button', { name: '2 行 2 列' })
    .click();
  await page.getByRole('tab', { name: '表格布局' }).click();

  const deleteTable = page.getByRole('button', { name: '删除表格' });
  await revealRibbonControl(page, '表格布局', deleteTable);
  await deleteTable.click();

  await expect(
    page.locator('.work-document-editable .ProseMirror table'),
  ).toHaveCount(0);
  await expect(page.getByRole('tab', { name: '表格设计' })).toHaveCount(0);
  await expect(page.getByRole('tab', { name: '表格布局' })).toHaveCount(0);
  await expect(page.getByRole('tab', { name: '开始' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(body).toBeFocused();
});

async function revealRibbonControl(
  page: import('@playwright/test').Page,
  tabLabel: string,
  control: import('@playwright/test').Locator,
): Promise<void> {
  const toolbar = page.getByRole('toolbar', { name: `${tabLabel}工具栏` });
  await expect(toolbar).toBeVisible();
  for (let attempt = 0; attempt < 20; attempt += 1) {
    if (await ribbonControlIsReachable(control)) return;
    const next = page.getByRole('button', {
      name: `向右查看更多${tabLabel}工具`,
    });
    await expect(next).toBeVisible();
    const previousScroll = await toolbar.evaluate((element) =>
      Math.round(element.scrollLeft),
    );
    await next.click();
    await expect
      .poll(() => toolbar.evaluate((element) => Math.round(element.scrollLeft)))
      .toBeGreaterThan(previousScroll);
  }
  throw new Error(`${tabLabel}中的目标控件无法通过功能区导航到达。`);
}

async function ribbonControlIsReachable(
  control: import('@playwright/test').Locator,
): Promise<boolean> {
  return control.evaluate((element) => {
    const toolbar = element.closest('.work-office-toolbar');
    if (!(toolbar instanceof HTMLElement)) return false;
    const controlBounds = element.getBoundingClientRect();
    const toolbarBounds = toolbar.getBoundingClientRect();
    const safeInset = toolbar.dataset.hasOverflow ? 34 : 0;
    return (
      controlBounds.left >= toolbarBounds.left + safeInset &&
      controlBounds.right <= toolbarBounds.right - safeInset
    );
  });
}
