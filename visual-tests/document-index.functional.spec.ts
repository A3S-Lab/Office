import { expect, test, type Locator } from '@playwright/test';

test('Writer edits and customizes a native document index responsively', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  if (testInfo.project.name === 'compact-768') {
    await page.setViewportSize({ width: 390, height: 844 });
  }

  await page.goto('/playground/');
  const latestCapabilities = page.getByRole('region', { name: '最新能力' });
  const documentIndexEntry = latestCapabilities.getByRole('button', {
    name: '打开最新能力：原生索引',
  });
  await expect(latestCapabilities).toBeInViewport();
  await expect(documentIndexEntry).toBeInViewport();
  await documentIndexEntry.click();

  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  const documentIndex = editor.locator('.work-document-index');
  await expect(documentIndex).toHaveAttribute('data-index-columns', '2');
  await expect(documentIndex).toHaveAttribute('data-index-format', 'indented');
  await expect(documentIndex).toHaveAttribute('data-index-leader', 'dot');
  await expect(
    documentIndex.locator(
      "[data-index-main-entry='Architecture'][data-index-sub-entry='Runtime']",
    ),
  ).toBeVisible();
  await expect(
    documentIndex.locator(
      "[data-index-main-entry='Collaboration'] .work-document-index-cross-reference",
    ),
  ).toHaveText('参见 Architecture');
  await expect(
    documentIndex.locator(
      "a[data-index-target='index-entry-performance'].italic",
    ),
  ).toHaveText('2');

  await documentIndex
    .locator("a[data-index-target='index-entry-runtime']")
    .click();
  const runtimeMarker = editor.locator(
    ".work-document-index-entry[data-index-entry-id='index-entry-runtime']",
  );
  await expect(runtimeMarker).toHaveClass(/ProseMirror-selectednode/);

  await page.getByRole('tab', { name: '引用' }).click();
  await page.getByRole('button', { name: '标记索引项' }).click();
  const entryDialog = page.getByRole('dialog', { name: '编辑索引项' });
  await expect(entryDialog).toBeVisible();
  await expect(entryDialog.getByLabel('主索引项')).toBeFocused();
  await expect(entryDialog.getByLabel('主索引项')).toHaveValue('Architecture');
  await expect(entryDialog.getByLabel('次索引项')).toHaveValue('Runtime');
  await expectDialogInsideViewport(entryDialog);
  await entryDialog.getByRole('checkbox', { name: '页码倾斜' }).check();
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-document-index-entry-dialog-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await entryDialog.getByRole('button', { name: '应用' }).click();

  await expect(runtimeMarker).toHaveAttribute('data-index-page-bold', 'true');
  await expect(runtimeMarker).toHaveAttribute('data-index-page-italic', 'true');
  await expect(editor).toBeFocused();
  await page.keyboard.press(`${modifier}+z`);
  await expect(runtimeMarker).toHaveAttribute('data-index-page-bold', 'true');
  await expect(runtimeMarker).toHaveAttribute(
    'data-index-page-italic',
    'false',
  );

  await documentIndex.locator('.work-document-index-header').click();
  await expect(documentIndex).toHaveClass(/ProseMirror-selectednode/);
  await page.getByRole('button', { name: '插入或自定义索引' }).click();
  const indexDialog = page.getByRole('dialog', { name: '自定义索引' });
  await expect(indexDialog).toBeVisible();
  await expect(
    indexDialog.getByRole('combobox', { name: '索引栏数' }),
  ).toBeFocused();
  await expectDialogInsideViewport(indexDialog);

  await indexDialog.getByRole('combobox', { name: '索引栏数' }).click();
  await page.getByRole('option', { name: '三栏', exact: true }).click();
  await indexDialog.getByRole('combobox', { name: '索引布局' }).click();
  await page.getByRole('option', { name: '连续式', exact: true }).click();
  await indexDialog.getByRole('combobox', { name: '索引前导符' }).click();
  await page
    .getByRole('option', { name: '短横线（----）', exact: true })
    .click();
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-document-index-dialog-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await indexDialog.getByRole('button', { name: '应用' }).click();

  await expect(documentIndex).toHaveAttribute('data-index-columns', '3');
  await expect(documentIndex).toHaveAttribute('data-index-format', 'run-in');
  await expect(documentIndex).toHaveAttribute('data-index-leader', 'dash');
  await expect(editor).toBeFocused();
  await page.keyboard.press(`${modifier}+z`);
  await expect(documentIndex).toHaveAttribute('data-index-columns', '2');
  await expect(documentIndex).toHaveAttribute('data-index-format', 'indented');
  await expect(documentIndex).toHaveAttribute('data-index-leader', 'dot');
  expect(browserErrors).toEqual([]);
});

async function expectDialogInsideViewport(dialog: Locator): Promise<void> {
  const geometry = await dialog.evaluate((element) => {
    const bounds = element.getBoundingClientRect();
    return {
      bottom: bounds.bottom,
      centerX: bounds.left + bounds.width / 2,
      left: bounds.left,
      right: bounds.right,
      top: bounds.top,
      viewportHeight: document.documentElement.clientHeight,
      viewportWidth: document.documentElement.clientWidth,
    };
  });
  expect(geometry.left).toBeGreaterThanOrEqual(8);
  expect(geometry.right).toBeLessThanOrEqual(geometry.viewportWidth - 8);
  expect(geometry.top).toBeGreaterThanOrEqual(8);
  expect(geometry.bottom).toBeLessThanOrEqual(geometry.viewportHeight - 8);
  expect(geometry.centerX).toBeCloseTo(geometry.viewportWidth / 2, 0);
}
