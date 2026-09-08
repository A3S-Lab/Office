import { expect, type Locator, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer applies double and words underline with WPS shortcuts', async ({
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

  const textRibbon = page.getByRole('region', { name: '文字功能区' });
  await textRibbon.getByRole('button', { name: '更多下划线' }).click();
  const menu = page.getByRole('menu', { name: '下划线样式' });
  await expect(
    menu.getByRole('menuitemradio', { name: '双下划线' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+Shift+D Meta+Shift+D');
  await expect(
    menu.getByRole('menuitemradio', { name: '仅字下划线' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+Shift+W Meta+Shift+W');
  await page.keyboard.press('Escape');

  const title = page.locator('.work-document-editable .ProseMirror h1').first();
  await title.click();
  await selectBlockText(title);
  await page.keyboard.press('Control+Shift+d');
  await expect(
    title.locator(
      "[data-office-underline-style='double'][style*='text-decoration-style: double']",
    ),
  ).toHaveCount(1);

  await page.keyboard.press('Control+Shift+w');
  await expect(
    title.locator("[data-office-underline-style='words']"),
  ).toHaveCount(1);

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-underline-shortcuts-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});

async function selectBlockText(block: Locator): Promise<void> {
  await block.evaluate((element) => {
    const range = document.createRange();
    range.selectNodeContents(element);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
    document.dispatchEvent(new Event('selectionchange', { bubbles: true }));
  });
}
