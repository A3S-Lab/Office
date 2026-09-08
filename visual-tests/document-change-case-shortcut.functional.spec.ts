import { expect, type Locator, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer cycles selection case with WPS Shift+F3', async ({
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
    textRibbon.getByRole('button', { name: '大小写效果' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Shift+F3');

  const editor = page.locator('.work-document-editable .ProseMirror');
  const paragraph = editor.locator('p').first();
  await paragraph.click();
  await selectBlockText(paragraph);
  await page.keyboard.type('hello world');
  await selectBlockText(paragraph);

  await page.keyboard.press('Shift+F3');
  await expect(paragraph).toHaveText('HELLO WORLD');

  await selectBlockText(paragraph);
  await page.keyboard.press('Shift+F3');
  await expect(paragraph).toHaveText('Hello World');

  await selectBlockText(paragraph);
  await page.keyboard.press('Shift+F3');
  await expect(paragraph).toHaveText('hello world');
  await expect(editor).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-change-case-shortcut-${testInfo.project.name}.png`,
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
