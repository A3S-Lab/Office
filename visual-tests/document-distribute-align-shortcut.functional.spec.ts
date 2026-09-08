import { expect, type Locator, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer steps distributed alignment with WPS Ctrl+Shift+J', async ({
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
  const distributeButton = textRibbon.getByRole('button', { name: '分散对齐' });
  await expect(distributeButton).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+J Meta+Shift+J',
  );

  const editor = page.locator('.work-document-editable .ProseMirror');
  const paragraph = editor.locator('p').first();
  await paragraph.click();
  await selectBlockText(paragraph);

  await editor.press('Control+Shift+j');
  await expect(paragraph).toHaveAttribute('data-office-text-align', 'distribute');
  await expect(paragraph).toHaveAttribute('style', /text-align:\s*justify/);
  await expect(distributeButton).toHaveClass(/active/);
  await expect(editor).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-distribute-align-shortcut-${testInfo.project.name}.png`,
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
