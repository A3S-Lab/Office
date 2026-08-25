import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('editable Office surfaces claim first-open focus and accept keyboard input', async ({
  page,
}) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));

  await page.goto('/playground/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);
  const documentBody = page.getByRole('textbox', { name: '文档正文' });
  await expect(documentBody).toBeFocused();
  await expect(documentBody).toHaveAttribute('contenteditable', 'true');
  await page.keyboard.type(' FIRST_OPEN_WRITER');
  await expect(documentBody).toContainText('FIRST_OPEN_WRITER');

  await page.goto('/playground/');
  await page.getByRole('button', { name: '# 产品说明 MD · 本次会话' }).click();
  const markdownSource = page.getByRole('textbox', { name: 'Markdown 源码' });
  await expect(markdownSource).toBeFocused();
  await expect(markdownSource).toBeEditable();
  await page.keyboard.type('FIRST_OPEN_MARKDOWN');
  await expect(markdownSource).toHaveValue(/FIRST_OPEN_MARKDOWN/);

  await page.goto('/playground/');
  await page
    .getByRole('button', { name: '季度执行计划 XLSX · 本次会话' })
    .click();
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();
  const spreadsheetGrid = page.locator('.fortune-sheet-overlay');
  await expect(spreadsheetGrid).toBeFocused();
  await page.keyboard.type('FIRST_OPEN_SHEET');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await expect(page.locator('.fortune-fx-input')).toHaveText(
    'FIRST_OPEN_SHEET',
  );

  await page.goto('/playground/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();
  await page.locator('.work-slide-canvas.interactive').waitFor();
  await expect(page.locator('[data-slide-thumbnail].active')).toBeFocused();

  expect(browserErrors).toEqual([]);
});
