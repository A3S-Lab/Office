import { expect, type Locator, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer toggles double strikethrough with WPS Ctrl+Shift+X', async ({
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
  await textRibbon.getByRole('button', { name: '更多删除线' }).click();
  const menu = page.getByRole('menu', { name: '删除线样式' });
  await expect(
    menu.getByRole('menuitemradio', { name: '双删除线' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+Shift+X Meta+Shift+X');
  await page.keyboard.press('Escape');

  const title = page.locator('.work-document-editable .ProseMirror h1').first();
  await title.click();
  await selectBlockText(title);
  await page.keyboard.press('Control+Shift+x');
  await expect(
    title.locator('s[data-office-strike-style="double"]'),
  ).toHaveCount(1);

  await page.keyboard.press('Control+Shift+x');
  await expect(
    title.locator('s[data-office-strike-style="double"]'),
  ).toHaveCount(0);
  await expect(page.locator('.work-document-editable .ProseMirror')).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-strike-shortcut-${testInfo.project.name}.png`,
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
