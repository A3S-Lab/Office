import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer authors a bounded content control from the Insert ribbon', async ({
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
  await page.getByRole('button', { name: '插入内容控件' }).click();

  const dialog = page.getByRole('dialog', { name: '插入内容控件' });
  await expect(dialog).toBeVisible();
  await dialog
    .getByRole('textbox', { name: '内容控件显示名称' })
    .fill('客户名称');
  await dialog
    .getByRole('textbox', { name: '内容控件程序标签' })
    .fill('customer-name');
  await dialog.getByRole('combobox', { name: '内容控件外观' }).click();
  await page.getByRole('option', { name: '标签', exact: true }).click();
  await dialog.getByRole('button', { name: '插入控件' }).click();

  const control = page.locator(
    '.work-document-editable [data-document-content-control]',
  );
  await expect(control).toHaveCount(1);
  await expect(control).toHaveAttribute(
    'data-content-control-alias',
    '客户名称',
  );
  await expect(control).toHaveAttribute(
    'data-content-control-tag',
    'customer-name',
  );
  await expect(control).toHaveAttribute(
    'data-content-control-appearance',
    'tags',
  );

  await control.click();
  await page.keyboard.type('季度联系人');
  await expect(control).toContainText('季度联系人');
  await expect(control).toBeVisible();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-content-control-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});
