import { expect, type Locator, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer moves paragraphs with WPS Alt+Shift+ArrowUp / ArrowDown', async ({
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
  const scope = editor.getByText('工作范围', { exact: true });
  await scope.click();

  const orderBefore = await paragraphOrder(editor);
  const scopeIndexBefore = orderBefore.indexOf('工作范围');
  expect(scopeIndexBefore).toBeGreaterThan(0);

  await editor.press('Alt+Shift+ArrowUp');
  const orderAfterUp = await paragraphOrder(editor);
  expect(orderAfterUp.indexOf('工作范围')).toBe(scopeIndexBefore - 1);

  await editor.press('Alt+Shift+ArrowDown');
  const orderAfterDown = await paragraphOrder(editor);
  expect(orderAfterDown.indexOf('工作范围')).toBe(scopeIndexBefore);
  await expect(editor).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-move-block-shortcuts-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});

async function paragraphOrder(editor: Locator): Promise<string[]> {
  return editor.locator('p, h1, h2, h3').evaluateAll((nodes) =>
    nodes
      .map((node) => (node.textContent ?? '').trim())
      .filter((text) => text.length > 0),
  );
}
