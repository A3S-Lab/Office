import { expect, test } from '@playwright/test';

test('Writer edits an imported WPS straight connector through its contextual ribbon', async ({
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
    .setInputFiles('.a3s-test/fixtures/word-wps-connector.docx');

  const connector = page.locator(
    '.work-document-editable [data-document-connector]',
  );
  await expect(connector).toHaveCount(1);
  await expect(connector).toHaveAttribute('data-connector-layout', 'floating');
  await connector.click();
  await expect(page.getByRole('tab', { name: '连接符' })).toHaveAttribute(
    'aria-selected',
    'true',
  );

  await page.getByRole('combobox', { name: '连接符终点箭头' }).click();
  await page.getByRole('option', { name: '三角箭头', exact: true }).click();
  await page.getByRole('button', { name: '浮于文字上方' }).click();
  await page
    .getByRole('textbox', { name: '连接符终点 Y（百分比）', exact: true })
    .fill('20');
  await page.getByRole('button', { name: '连接符线条颜色' }).click();
  await page.getByRole('option', { name: '颜色 #0070c0' }).click();

  await expect(connector).toHaveAttribute(
    'data-connector-end-arrow',
    'triangle',
  );
  await expect(connector).toHaveAttribute('data-connector-end-y', '20');
  await expect(connector).toHaveAttribute(
    'data-connector-line-color',
    '#0070c0',
  );

  await page.keyboard.press('Control+z');
  await expect(connector).toHaveAttribute(
    'data-connector-line-color',
    '#c00000',
  );
  await page.keyboard.press('Control+Shift+z');
  await expect(connector).toHaveAttribute(
    'data-connector-line-color',
    '#0070c0',
  );

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-wps-connector-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});
