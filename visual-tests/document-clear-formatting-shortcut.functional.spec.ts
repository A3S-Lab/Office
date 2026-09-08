import { expect, type Locator, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer clears formatting with WPS Ctrl+Space', async ({
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
  await expect(
    textRibbon.getByRole('button', { name: '清除格式' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+Space Meta+Space');

  const paragraph = page
    .locator('.work-document-editable .ProseMirror p')
    .first();
  await paragraph.click();
  await selectBlockText(paragraph);
  await page.keyboard.press('Control+b');
  await page.keyboard.press('Control+i');
  await expect(paragraph.locator('strong em, em strong')).toHaveCount(1);

  await page.keyboard.press('Control+Space');
  await expect(paragraph.locator('strong')).toHaveCount(0);
  await expect(paragraph.locator('em')).toHaveCount(0);
  await expect(
    page.locator('.work-document-editable .ProseMirror'),
  ).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-clear-formatting-shortcut-${testInfo.project.name}.png`,
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
