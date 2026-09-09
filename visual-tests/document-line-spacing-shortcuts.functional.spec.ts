import { expect, type Locator, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer steps line spacing with WPS Ctrl+1 / Ctrl+5 / Ctrl+2', async ({
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
    textRibbon.getByRole('combobox', { name: '行距' }),
  ).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+1 Meta+1 Control+5 Meta+5 Control+2 Meta+2',
  );

  const editor = page.locator('.work-document-editable .ProseMirror');
  const paragraph = editor.locator('p').first();
  await paragraph.click();
  await selectBlockText(paragraph);

  await editor.press('Control+Digit5');
  await expect(paragraph).toHaveAttribute('style', /line-height:\s*1\.5/);

  await editor.press('Control+Digit2');
  await expect(paragraph).toHaveAttribute(
    'style',
    /line-height:\s*2(?:;|$|\s)/,
  );

  await editor.press('Control+Digit1');
  await expect(paragraph).toHaveAttribute(
    'style',
    /line-height:\s*1(?:;|$|\s)/,
  );
  await expect(editor).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-line-spacing-shortcuts-${testInfo.project.name}.png`,
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
