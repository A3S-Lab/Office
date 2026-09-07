import { expect, type Locator, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer keeps WPS shortcuts discoverable and keyboard reachable', async ({
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

  const editor = page.locator('.work-document-editable .ProseMirror');
  const body = page.getByRole('textbox', { name: '文档正文' });
  const textRibbon = page.getByRole('region', { name: '文字功能区' });
  const formatCopy = textRibbon.getByRole('button', { name: '复制格式' });
  await expect(formatCopy).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+C Meta+Shift+C',
  );
  const textCaseTrigger = textRibbon.getByRole('button', {
    name: '大小写效果',
  });
  await textCaseTrigger.click();
  const textCaseMenu = page.getByRole('menu', { name: '大小写效果' });
  await expect(textCaseMenu).toBeVisible();
  await expect(
    textCaseMenu.getByRole('menuitemradio', { name: '全部大写' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+Shift+A Meta+Shift+A');
  await expect(
    textCaseMenu.getByRole('menuitemradio', { name: '小型大写' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+Shift+K Meta+Shift+K');
  await expect(
    textCaseMenu.getByRole('menuitemradio', { name: '常规' }),
  ).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(textCaseTrigger).toBeFocused();

  const source = editor.locator('h1').first();
  await source.click();
  await selectBlockText(source);
  await page.keyboard.press('Control+i');
  await page.keyboard.press('Control+Shift+c');

  const target = editor.locator('p').first();
  await target.click();
  await selectBlockText(target);
  await page.keyboard.press('Control+Shift+v');
  await expect(editor.locator('h1:has(em) ~ h1:has(em)')).toHaveCount(1);

  await page.keyboard.press('Control+Shift+a');
  const allCaps = editor.locator(
    'h1 ~ h1 span[data-office-text-case="all-caps"]',
  );
  await expect(allCaps).toHaveCSS('text-transform', 'uppercase');
  await page.keyboard.press('Control+Shift+k');
  const smallCaps = editor.locator(
    'h1 ~ h1 span[data-office-text-case="small-caps"]',
  );
  await expect(smallCaps).toHaveCSS('font-variant-caps', 'small-caps');
  await page.keyboard.press('Control+z');
  await expect(allCaps).toHaveCount(1);

  await target.click();
  await selectBlockText(target);
  await page.keyboard.press('Control+e');
  await expect(target).toHaveAttribute('style', /text-align:\s*center/);
  await page.keyboard.press('Control+1');
  await expect(target).toHaveAttribute('style', /line-height:\s*1/);
  await page.keyboard.press('Control+Alt+2');
  const headingTwo = editor.locator('h2').last();
  await expect(headingTwo).toBeVisible();

  await page.keyboard.press('F7');
  await expect(body).toHaveAttribute('spellcheck', 'false');
  await headingTwo.click();
  await page.keyboard.press('Control+Shift+e');
  await expect(body).toBeFocused();
  const reviewTab = page.getByRole('tab', { name: '审阅' });
  await reviewTab.click();
  await expect(reviewTab).toHaveAttribute('aria-selected', 'true');
  const trackChanges = page.getByRole('button', { name: '修订模式' });
  await expect(trackChanges).toHaveAttribute('aria-pressed', 'true');
  await expect(page.getByRole('button', { name: '添加批注' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Alt+M Meta+Alt+M',
  );

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-wps-shortcuts-${testInfo.project.name}.png`,
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
