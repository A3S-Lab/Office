import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer inserts a page break with WPS Ctrl+Enter', async ({
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

  await page.getByRole('tab', { name: '插入' }).click();
  const insertPageBreak = page.getByRole('button', { name: '插入分页符' });
  await expect(insertPageBreak).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Enter Meta+Enter',
  );

  const editor = page.locator('.work-document-editable .ProseMirror');
  await editor.click();
  await page.keyboard.press('Control+Enter');
  await expect(
    editor.locator('[data-page-break="true"][aria-label="分页符"]'),
  ).toHaveCount(1);
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  await expect(editor).toHaveAttribute('data-pagination-pages', '2');

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-page-break-shortcut-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});
