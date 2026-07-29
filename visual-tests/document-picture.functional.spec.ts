import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

const pixelPng = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z9WQAAAAASUVORK5CYII=',
  'base64',
);

test('picture contextual ribbon keeps selection, dialog focus, and cleanup coherent', async ({
  page,
}) => {
  await page.goto('/');
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

  await page.getByRole('button', { name: '替代文字' }).click();
  const dialog = page.getByRole('dialog', { name: '图片说明' });
  const alternativeText = dialog.getByRole('textbox', {
    name: '图片替代文字',
  });
  await expect(alternativeText).toBeFocused();
  await alternativeText.fill('不应保存的说明');
  await dialog.getByRole('button', { name: '取消' }).click();
  await expect(dialog).toBeHidden();
  await expect(body).toBeFocused();
  await expect(pictureTab).toHaveAttribute('aria-selected', 'true');

  await page.getByRole('button', { name: '替代文字' }).click();
  await alternativeText.fill('季度计划趋势图');
  await dialog.getByRole('button', { name: '保存' }).click();
  await expect(imageContainer.locator('img[alt="季度计划趋势图"]')).toHaveCount(
    1,
  );
  await expect(body).toBeFocused();

  await page.getByRole('button', { name: '删除图片' }).click();
  await expect(imageContainer).toHaveCount(0);
  await expect(pictureTab).toHaveCount(0);
  await expect(page.getByRole('tab', { name: '开始' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(body).toBeFocused();
});
