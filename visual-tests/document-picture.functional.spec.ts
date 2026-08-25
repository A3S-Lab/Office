import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

const pixelPng = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z9WQAAAAASUVORK5CYII=',
  'base64',
);

test('picture contextual ribbon keeps properties, focus, and cleanup coherent', async ({
  page,
}) => {
  const originalViewport = page.viewportSize();
  await page.goto('/playground/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  const body = page.getByRole('textbox', { name: '文档正文' });
  await page.locator('.work-document-editable .ProseMirror p').first().click();
  await page.locator('input[aria-label="插入文档图片"]').setInputFiles({
    name: 'quarterly-plan.png',
    mimeType: 'image/png',
    buffer: pixelPng,
  });

  const imageContainer = page.locator(
    '.work-document-editable [data-resize-container][data-node="image"]',
  );
  await expect(imageContainer).toHaveCount(1);
  const pictureTab = page.getByRole('tab', { name: '图片' });
  await expect(pictureTab).toHaveAttribute('aria-selected', 'true');

  await page.getByRole('button', { name: '四周环绕' }).click();
  await expect(imageContainer).toHaveAttribute(
    'data-office-image-layout',
    'square',
  );
  await page.getByRole('button', { name: '右对齐' }).click();
  await expect(imageContainer).toHaveAttribute(
    'data-office-image-alignment',
    'right',
  );

  const distance = page.getByRole('combobox', { name: '图片与文字距离' });
  await distance.click();
  await page.getByRole('option', { name: '10 毫米' }).click();
  await expect(imageContainer).toHaveAttribute(
    'data-office-image-wrap-distance',
    '10',
  );
  await expect(body).toBeFocused();
  await expect(pictureTab).toHaveAttribute('aria-selected', 'true');

  const propertiesButton = page.getByRole('button', {
    name: '图片属性',
  });
  await propertiesButton.click();
  const dialog = page.getByRole('dialog', { name: '图片属性' });
  const alternativeText = dialog.getByRole('textbox', {
    name: '图片替代文字',
  });
  await expect(
    dialog.getByRole('textbox', { name: '图片宽度（厘米）' }),
  ).toBeFocused();
  await alternativeText.fill('不应保存的说明');
  await dialog.getByRole('button', { name: '取消' }).click();
  await expect(dialog).toBeHidden();
  await expect(propertiesButton).toBeFocused();
  await expect(pictureTab).toHaveAttribute('aria-selected', 'true');
  await expect(
    imageContainer.locator('img[alt="quarterly-plan.png"]'),
  ).toHaveCount(1);

  await propertiesButton.click();
  const width = dialog.getByRole('textbox', {
    name: '图片宽度（厘米）',
  });
  const height = dialog.getByRole('textbox', {
    name: '图片高度（厘米）',
  });
  await width.fill('5');
  await expect(height).toHaveValue('5');
  await dialog.getByRole('checkbox', { name: '锁定纵横比' }).uncheck();
  await height.fill('3');
  await alternativeText.fill('季度计划趋势图');
  await dialog.getByRole('button', { name: '确定' }).click();
  await expect(imageContainer.locator('img[alt="季度计划趋势图"]')).toHaveCount(
    1,
  );
  await expect(imageContainer.locator('img')).toHaveAttribute('width', '189');
  await expect(imageContainer.locator('img')).toHaveAttribute('height', '113');
  await expect(imageContainer).toHaveAttribute(
    'data-office-image-lock-aspect-ratio',
    'false',
  );
  await expect(propertiesButton).toBeFocused();
  await expect(imageContainer).toHaveClass(/ProseMirror-selectednode/);

  await page.setViewportSize({ width: 390, height: 844 });
  await propertiesButton.press('Enter');
  await expect(dialog).toBeVisible();
  const dialogBounds = await dialog.boundingBox();
  expect(dialogBounds?.x).toBeGreaterThanOrEqual(0);
  expect(
    (dialogBounds?.x ?? 0) + (dialogBounds?.width ?? 0),
  ).toBeLessThanOrEqual(390);
  const widthBounds = await width.boundingBox();
  expect(widthBounds?.height).toBeGreaterThanOrEqual(40);
  await alternativeText.fill('仍不应保存的说明');
  await alternativeText.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(propertiesButton).toBeFocused();
  await expect(imageContainer.locator('img[alt="季度计划趋势图"]')).toHaveCount(
    1,
  );

  if (originalViewport) await page.setViewportSize(originalViewport);

  await page.getByRole('button', { name: '删除图片' }).click();
  await expect(imageContainer).toHaveCount(0);
  await expect(pictureTab).toHaveCount(0);
  await expect(page.getByRole('tab', { name: '开始' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(body).toBeFocused();
});
