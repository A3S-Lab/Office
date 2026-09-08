import { expect, type Locator, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer steps paragraph alignment with WPS Ctrl+L / Ctrl+E / Ctrl+R / Ctrl+J', async ({
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
    textRibbon.getByRole('button', { name: '左对齐' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+L Meta+L');
  await expect(
    textRibbon.getByRole('button', { name: '居中' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+E Meta+E');
  await expect(
    textRibbon.getByRole('button', { name: '右对齐' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+R Meta+R');
  await expect(
    textRibbon.getByRole('button', { name: '两端对齐' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+J Meta+J');

  const editor = page.locator('.work-document-editable .ProseMirror');
  const paragraph = editor.locator('p').first();
  await paragraph.click();
  await selectBlockText(paragraph);

  await editor.press('Control+e');
  await expect(paragraph).toHaveAttribute('style', /text-align:\s*center/);

  await editor.press('Control+r');
  await expect(paragraph).toHaveAttribute('style', /text-align:\s*right/);

  await editor.press('Control+j');
  await expect(paragraph).toHaveAttribute('style', /text-align:\s*justify/);

  await editor.press('Control+l');
  await expect(paragraph).toHaveAttribute('style', /text-align:\s*left/);
  await expect(editor).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-align-shortcuts-${testInfo.project.name}.png`,
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
