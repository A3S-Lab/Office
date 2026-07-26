import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  stabilizeVisualSurface,
  waitForDocumentFixture,
} from './visual-test-support';

test('document table Design and Layout stay visual and survive preview', async ({
  page,
}) => {
  await page.goto('/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  await page.locator('.work-document-editable .ProseMirror p').first().click();
  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '插入表格' }).click();
  const picker = page.getByRole('dialog', { name: '选择表格大小' });
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
