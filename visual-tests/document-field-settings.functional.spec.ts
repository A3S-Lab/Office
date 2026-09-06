import { expect, test } from '@playwright/test';

test('Writer field settings keeps format choices visible and inserts a live field', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });

  await page.goto('/playground/?e2e=word-field-settings');
  await expect(
    page.locator(
      '.work-document-editable .ProseMirror[data-pagination-state="ready"]',
    ),
  ).toBeVisible();
  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '字段设置' }).click();

  const dialog = page.getByRole('dialog', { name: '插入字段' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('combobox', { name: '字段类型' })).toHaveText(
    '页码',
  );
  await dialog.getByRole('combobox', { name: '数字格式' }).click();
  await page.getByRole('option', { name: '小写罗马数字（i）' }).click();
  await expect(dialog.getByRole('status')).toHaveText(/i/);
  await dialog.getByRole('button', { name: '插入字段' }).click();

  const field = page.locator(
    '.work-document-editable .work-document-field[data-field-kind="page"]',
  );
  await expect(field).toHaveAttribute('data-field-display', 'i');
  await expect(field).toHaveAttribute(
    'data-field-instruction',
    'PAGE \\* roman',
  );

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-field-settings-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});
