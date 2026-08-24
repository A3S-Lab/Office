import { expect, test } from '@playwright/test';

test('Spreadsheet data validation is discoverable from the public Playground', async ({
  page,
}) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));

  await page.goto('/');
  await expect(page.getByRole('heading', { name: '我的文档' })).toBeVisible();
  await page
    .getByRole('button', {
      name: '数据验证 下拉列表、输入提示与错误警告',
    })
    .click();
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();
  await expect(page.getByRole('textbox', { name: '文件名' })).toHaveValue(
    '数据验证示例',
  );

  const ribbon = page.locator('.work-spreadsheet-ribbon');
  await ribbon.getByRole('tab', { name: '数据' }).click();
  await ribbon.getByRole('button', { name: '数据验证' }).click();

  const dialog = page.getByRole('dialog', { name: '数据验证' });
  await expect(dialog).toContainText('Inputs!B2:B6');
  await expect(dialog.getByRole('textbox', { name: '来源' })).toHaveValue(
    "'Lists'!A1:A3",
  );
  await expect(
    dialog.getByRole('checkbox', { name: '忽略空值' }),
  ).not.toBeChecked();
  await expect(
    dialog.getByRole('checkbox', { name: '在单元格内显示下拉箭头' }),
  ).toBeChecked();
  await expect(
    dialog.getByRole('checkbox', { name: '选中单元格时显示输入信息' }),
  ).toBeChecked();
  await expect(
    dialog.getByRole('textbox', { name: '输入信息标题' }),
  ).toHaveValue('Workflow state');
  await expect(
    dialog.getByRole('combobox', { name: '错误警告样式' }),
  ).toHaveValue('stop');
  await expect(
    dialog.getByRole('textbox', { name: '错误警告标题' }),
  ).toHaveValue('Invalid state');
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet data validation stays atomic and accessible at every layout', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.goto('/?e2e=spreadsheet-data-validation');
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();

  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const undo = ribbon.getByRole('button', { name: '撤销' });
  await ribbon.getByRole('tab', { name: '数据' }).click();
  const launcher = ribbon.getByRole('button', { name: '数据验证' });
  await launcher.focus();
  await launcher.click();

  let dialog = page.getByRole('dialog', { name: '数据验证' });
  await expect(dialog).toBeVisible();
  await expect(dialog).toContainText('Inputs!B2:B3,D5:E5');
  await expect(dialog).toContainText('2 个选定区域');
  await expectDialogInsideViewport(page, dialog);
  await page.keyboard.press('Escape');
  await expect(dialog).toHaveCount(0);
  await expect(launcher).toBeFocused();

  await launcher.click();
  dialog = page.getByRole('dialog', { name: '数据验证' });
  const source = dialog.getByRole('textbox', { name: '来源' });
  await source.fill("'Lists'!A1:B2");
  await expect(dialog.getByRole('alert')).toHaveText(
    '下拉列表来源只能是一行或一列连续单元格。',
  );
  await expect(dialog.getByRole('button', { name: '确定' })).toBeDisabled();
  await source.fill("'Lists'!A1:A3");
  await dialog
    .getByRole('checkbox', { name: '选中单元格时显示输入信息' })
    .click();
  await dialog
    .getByRole('textbox', { name: '输入信息标题' })
    .fill('Workflow state');
  await dialog
    .getByRole('textbox', { name: '输入信息', exact: true })
    .fill('Choose a workflow state.');
  await dialog.getByRole('checkbox', { name: '忽略空值' }).click();
  await dialog
    .getByRole('checkbox', { name: '在单元格内显示下拉箭头' })
    .click();
  await dialog
    .getByRole('combobox', { name: '错误警告样式' })
    .selectOption('warning');
  await dialog
    .getByRole('textbox', { name: '错误警告标题' })
    .fill('Invalid state');
  await dialog
    .getByRole('textbox', { name: '错误警告消息' })
    .fill('Choose Ready, Blocked, or In review.');
  await expect(dialog.getByRole('button', { name: '确定' })).toBeEnabled();
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-data-validation-dialog.png'),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '确定' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(launcher).toBeFocused();
  await expect(undo).toBeEnabled();

  await launcher.click();
  dialog = page.getByRole('dialog', { name: '数据验证' });
  await expect(dialog.getByRole('textbox', { name: '来源' })).toHaveValue(
    "'Lists'!A1:A3",
  );
  await expect(
    dialog.getByRole('checkbox', { name: '忽略空值' }),
  ).not.toBeChecked();
  await expect(
    dialog.getByRole('checkbox', { name: '在单元格内显示下拉箭头' }),
  ).not.toBeChecked();
  await expect(
    dialog.getByRole('checkbox', { name: '选中单元格时显示输入信息' }),
  ).toBeChecked();
  await expect(
    dialog.getByRole('textbox', { name: '输入信息标题' }),
  ).toHaveValue('Workflow state');
  await expect(
    dialog.getByRole('textbox', { name: '输入信息', exact: true }),
  ).toHaveValue('Choose a workflow state.');
  await expect(
    dialog.getByRole('combobox', { name: '错误警告样式' }),
  ).toHaveValue('warning');
  await expect(
    dialog.getByRole('textbox', { name: '错误警告标题' }),
  ).toHaveValue('Invalid state');
  await expect(
    dialog.getByRole('textbox', { name: '错误警告消息' }),
  ).toHaveValue('Choose Ready, Blocked, or In review.');
  await expect(dialog.getByRole('button', { name: '全部清除' })).toBeVisible();
  await dialog.getByRole('button', { name: '全部清除' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(launcher).toBeFocused();

  await launcher.click();
  dialog = page.getByRole('dialog', { name: '数据验证' });
  await expect(dialog.getByRole('button', { name: '全部清除' })).toHaveCount(0);
  await page.keyboard.press('Escape');
  await expect(launcher).toBeFocused();

  await undo.click();
  await launcher.click();
  dialog = page.getByRole('dialog', { name: '数据验证' });
  await expect(dialog.getByRole('button', { name: '全部清除' })).toBeVisible();
  await page.keyboard.press('Escape');

  await undo.click();
  await launcher.click();
  dialog = page.getByRole('dialog', { name: '数据验证' });
  await expect(dialog.getByRole('button', { name: '全部清除' })).toHaveCount(0);
  await page.keyboard.press('Escape');
  await expect(undo).toBeDisabled();
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
