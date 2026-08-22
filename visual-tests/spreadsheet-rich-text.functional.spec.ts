import { expect, test } from '@playwright/test';

test('Spreadsheet renders native rich text and formats every run through one controlled update', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.goto('/?e2e=spreadsheet-rich-text');

  const grid = page.locator('.fortune-sheet-overlay');
  const status = page.getByTestId('spreadsheet-rich-text-status');
  await expect(grid).toBeVisible();
  await expect(grid).toBeFocused();
  await expect(status).toHaveAttribute('data-revision', '1');
  await expect(status).toHaveAttribute('data-run-count', '3');
  await expect(status).toHaveAttribute('data-run-bold', '1,0,0');
  await expect(status).toHaveAttribute(
    'data-run-origins',
    'theme,indexed,automatic',
  );
  await expect(status).toHaveAttribute('data-run-text', 'Native rich text');

  await page.keyboard.press('Control+Shift+F');
  const dialog = page.getByRole('dialog', { name: '设置单元格格式' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('checkbox', { name: '加粗' }).click();
  await dialog.getByRole('button', { name: '应用' }).click();

  await expect(status).toHaveAttribute('data-revision', '2');
  await expect(status).toHaveAttribute('data-run-count', '3');
  await expect(status).toHaveAttribute('data-run-bold', '1,1,1');
  await expect(status).toHaveAttribute(
    'data-run-origins',
    'theme,indexed,automatic',
  );
  await expect(status).toHaveAttribute('data-run-text', 'Native rich text');
  await expect(grid).toBeFocused();
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-rich-text.png'),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet formats only selected rich text and restores the formula-bar selection', async ({
  page,
}) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.goto('/?e2e=spreadsheet-rich-text');

  const status = page.getByTestId('spreadsheet-rich-text-status');
  const formulaBar = page.locator('.fortune-fx-input');
  await expect(formulaBar).toHaveText('Native rich text');
  await formulaBar.click();
  await formulaBar.evaluate((editor) => {
    const text = editor.textContent ?? '';
    const startOffset = text.indexOf('rich');
    const endOffset = startOffset + 'rich'.length;
    const pointAtOffset = (requestedOffset: number) => {
      const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
      let remaining = requestedOffset;
      let node = walker.nextNode();
      while (node) {
        const length = node.textContent?.length ?? 0;
        if (remaining <= length) return { node, offset: remaining };
        remaining -= length;
        node = walker.nextNode();
      }
      throw new Error(`Cannot resolve formula-bar offset ${requestedOffset}.`);
    };
    if (startOffset < 0) throw new Error('Expected the rich-text fixture.');
    const start = pointAtOffset(startOffset);
    const end = pointAtOffset(endOffset);
    const range = document.createRange();
    range.setStart(start.node, start.offset);
    range.setEnd(end.node, end.offset);
    const selection = window.getSelection();
    if (!selection) throw new Error('Expected a browser text selection.');
    selection.removeAllRanges();
    selection.addRange(range);
    (editor as HTMLElement).focus();
  });
  expect(await page.evaluate(() => window.getSelection()?.toString())).toBe(
    'rich',
  );

  await page.getByRole('button', { name: '加粗' }).click();

  await expect(status).toHaveAttribute('data-revision', '2');
  await expect(status).toHaveAttribute('data-run-count', '4');
  await expect(status).toHaveAttribute('data-run-bold', '1,1,0,0');
  await expect(status).toHaveAttribute(
    'data-run-origins',
    'theme,indexed,indexed,automatic',
  );
  await expect(status).toHaveAttribute('data-run-text', 'Native rich text');
  await expect(formulaBar).toBeFocused();
  expect(await page.evaluate(() => window.getSelection()?.toString())).toBe(
    'rich',
  );
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet formula-bar insertion and deletion preserve native rich-text runs', async ({
  page,
}) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.goto('/?e2e=spreadsheet-rich-text');

  const status = page.getByTestId('spreadsheet-rich-text-status');
  const formulaBar = page.locator('.fortune-fx-input');
  await expect(formulaBar).toHaveText('Native rich text');
  await formulaBar.focus();
  await formulaBar.press('Home');
  await formulaBar.press('ArrowRight');
  await formulaBar.press('ArrowRight');
  await formulaBar.press('ArrowRight');
  await page.keyboard.insertText('X');

  await expect(formulaBar).toHaveText('NatXive rich text');
  await expect(status).toHaveAttribute('data-revision', '1');
  await formulaBar.press('Enter');
  await expect(status).toHaveAttribute('data-revision', '2');
  await expect(status).toHaveAttribute('data-run-count', '3');
  await expect(status).toHaveAttribute('data-run-text', 'NatXive rich text');
  await expect(status).toHaveAttribute('data-cell-text', 'NatXive rich text');
  await expect(status).toHaveAttribute(
    'data-run-origins',
    'theme,indexed,automatic',
  );

  await page.keyboard.press('ArrowUp');
  await expect(formulaBar).toHaveText('NatXive rich text');
  await formulaBar.focus();
  await formulaBar.press('End');
  await formulaBar.press('Backspace');
  await expect(formulaBar).toHaveText('NatXive rich tex');
  await expect(status).toHaveAttribute('data-revision', '2');
  await formulaBar.press('Enter');

  await expect(status).toHaveAttribute('data-revision', '3');
  await expect(status).toHaveAttribute('data-run-count', '3');
  await expect(status).toHaveAttribute('data-run-text', 'NatXive rich tex');
  await expect(status).toHaveAttribute('data-cell-text', 'NatXive rich tex');
  await expect(status).toHaveAttribute(
    'data-run-origins',
    'theme,indexed,automatic',
  );

  await page.keyboard.press('Control+z');
  await expect(status).toHaveAttribute('data-revision', '4');
  await expect(status).toHaveAttribute('data-run-text', 'NatXive rich text');
  await expect(status).toHaveAttribute(
    'data-run-origins',
    'theme,indexed,automatic',
  );
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet in-cell insertion and deletion retain rich-text semantics', async ({
  page,
}) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.goto('/?e2e=spreadsheet-rich-text');

  const grid = page.locator('.fortune-sheet-overlay');
  const cellEditor = page.locator('.luckysheet-cell-input');
  const status = page.getByTestId('spreadsheet-rich-text-status');
  await expect(grid).toBeFocused();
  await page.keyboard.press('F2');
  await expect(cellEditor).toHaveText('Native rich text');
  await page.keyboard.insertText('!');
  await expect(cellEditor).toHaveText('Native rich text!');
  await expect(status).toHaveAttribute('data-revision', '1');
  await page.keyboard.press('Enter');

  await expect(status).toHaveAttribute('data-revision', '2');
  await expect(status).toHaveAttribute('data-run-count', '3');
  await expect(status).toHaveAttribute('data-run-text', 'Native rich text!');
  await expect(status).toHaveAttribute(
    'data-run-origins',
    'theme,indexed,automatic',
  );

  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('F2');
  await page.keyboard.press('Backspace');
  await expect(cellEditor).toHaveText('Native rich text');
  await expect(status).toHaveAttribute('data-revision', '2');
  await page.keyboard.press('Enter');

  await expect(status).toHaveAttribute('data-revision', '3');
  await expect(status).toHaveAttribute('data-run-count', '3');
  await expect(status).toHaveAttribute('data-run-text', 'Native rich text');
  await expect(status).toHaveAttribute(
    'data-run-origins',
    'theme,indexed,automatic',
  );
  expect(browserErrors).toEqual([]);
});
