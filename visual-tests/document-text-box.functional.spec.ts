import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer authors a native text box through the contextual ribbon', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });

  await page.goto('/playground/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  await page.locator('.work-document-editable .ProseMirror p').first().click();
  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '插入文本框' }).click();

  // The pagination engine can recycle the ProseMirror subtree while it
  // measures a changed block. Scope from the stable editable surface so the
  // locator follows that replacement instead of retaining a transient
  // editor element handle.
  const textBox = page.locator(
    '.work-document-editable [data-document-text-box]',
  );
  await expect(textBox).toHaveCount(1);
  await expect(page.getByRole('tab', { name: '文本框' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await page.keyboard.type('季度计划重点');
  await expect(textBox).toContainText('季度计划重点');

  const shapeOptions = [
    ['圆角矩形', 'roundedRectangle'],
    ['椭圆', 'ellipse'],
    ['菱形', 'diamond'],
    ['三角形', 'triangle'],
    ['矩形', 'rectangle'],
  ] as const;
  for (const [label, value] of shapeOptions) {
    await page.getByRole('combobox', { name: '文本框形状' }).click();
    await page.getByRole('option', { name: label, exact: true }).click();
    await expect(textBox).toHaveAttribute('data-text-box-shape', value);
  }
  await page.getByRole('combobox', { name: '文本框形状' }).click();
  await page.getByRole('option', { name: '椭圆', exact: true }).click();
  await expect(textBox).toHaveAttribute('data-text-box-shape', 'ellipse');

  await page
    .getByRole('textbox', { name: '文本框宽度（毫米）', exact: true })
    .fill('76.2');
  await page
    .getByRole('textbox', { name: '文本框高度（毫米）', exact: true })
    .fill('25.4');
  await page.getByRole('button', { name: '浮于文字上方' }).click();
  await page.getByRole('combobox', { name: '文本框水平相对于' }).click();
  await page.getByRole('option', { name: '页面', exact: true }).click();
  await page.getByRole('combobox', { name: '文本框垂直相对于' }).click();
  await page.getByRole('option', { name: '页边距', exact: true }).click();
  await page
    .getByRole('textbox', { name: '文本框水平偏移（毫米）', exact: true })
    .fill('12.7');
  await page
    .getByRole('textbox', { name: '文本框垂直偏移（毫米）', exact: true })
    .fill('-6.35');
  await page.getByRole('button', { name: '底端' }).click();

  await page.getByRole('button', { name: '文本框填充颜色' }).click();
  await page.getByRole('option', { name: '颜色 #d9ead3' }).click();
  await page.getByRole('button', { name: '文本框边框颜色' }).click();
  await page.getByRole('option', { name: '颜色 #0070c0' }).click();
  await page.getByRole('combobox', { name: '文本框边框粗细' }).click();
  await page.getByRole('option', { name: '中（0.7 mm）' }).click();

  await expect(textBox).toHaveAttribute('data-text-box-width', '76.2');
  await expect(textBox).toHaveAttribute('data-text-box-height', '25.4');
  await expect(textBox).toHaveAttribute('data-text-box-layout', 'floating');
  await expect(textBox).toHaveAttribute(
    'data-text-box-horizontal-reference',
    'page',
  );
  await expect(textBox).toHaveAttribute(
    'data-text-box-vertical-reference',
    'margin',
  );
  await expect(textBox).toHaveAttribute(
    'data-text-box-horizontal-offset',
    '12.7',
  );
  await expect(textBox).toHaveAttribute(
    'data-text-box-vertical-offset',
    '-6.35',
  );
  await expect(textBox).toHaveAttribute('data-text-box-fill', '#d9ead3');
  await expect(textBox).toHaveAttribute(
    'data-text-box-border-color',
    '#0070c0',
  );
  await expect(textBox).toHaveAttribute('data-text-box-border-width', '0.7');
  await expect(textBox).toHaveAttribute(
    'data-text-box-vertical-align',
    'bottom',
  );
  await expect(textBox).toBeVisible();

  // Let the asynchronous paginator finish recycling the page subtree before
  // reading geometry from the live DOM node.
  await page.waitForTimeout(250);
  const bounds = await textBox.evaluate((element) => {
    const rect = element.getBoundingClientRect();
    return { left: rect.left, right: rect.right };
  });
  expect(bounds.left).toBeGreaterThanOrEqual(0);
  expect(bounds.right).toBeLessThanOrEqual(page.viewportSize()?.width ?? 1280);

  await page.screenshot({
    path: testInfo.outputPath(`writer-text-box-${testInfo.project.name}.png`),
    animations: 'disabled',
  });

  await page.getByRole('button', { name: '删除文本框' }).click();
  await expect(textBox).toHaveCount(0);
  await expect(page.getByRole('tab', { name: '文本框' })).toHaveCount(0);
  await expect(page.getByRole('tab', { name: '开始' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  expect(browserErrors).toEqual([]);
});

test('Writer keeps imported WPS shape controls discoverable and undoable', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });

  await page.goto('/playground/');
  await page
    .locator('input[aria-label="打开 Office 或 PDF 文件"]')
    .setInputFiles('.a3s-test/fixtures/word-wps-shape.docx');

  const textBox = page.locator(
    '.work-document-editable [data-document-text-box]',
  );
  await expect(textBox).toHaveAttribute(
    'data-text-box-shape',
    'roundedRectangle',
  );
  await expect(textBox).toContainText('WPS native shape');

  await textBox.click();
  await expect(page.getByRole('tab', { name: '文本框' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  const shapeSelect = page.getByRole('combobox', { name: '文本框形状' });
  await expect(shapeSelect).toBeVisible();
  await shapeSelect.click();
  await page.getByRole('option', { name: '椭圆', exact: true }).click();
  await expect(textBox).toHaveAttribute('data-text-box-shape', 'ellipse');

  await page.keyboard.press('Control+z');
  await expect(textBox).toHaveAttribute(
    'data-text-box-shape',
    'roundedRectangle',
  );
  await page.keyboard.press('Control+Shift+z');
  await expect(textBox).toHaveAttribute('data-text-box-shape', 'ellipse');

  await page.screenshot({
    path: testInfo.outputPath(`writer-wps-shape-${testInfo.project.name}.png`),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});
