import { expect, test } from '@playwright/test';

test('Writer keeps WPS numeric fields live, labelled, and responsive', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });

  await page.goto('/playground/?e2e=word-wps-numeric-fields');
  const editor = page.locator(
    '.work-document-editable .ProseMirror[data-pagination-state="ready"]',
  );
  await expect(editor).toBeVisible();
  await expect(editor).toHaveAttribute('data-pagination-engine', 'wasm');
  await expect(page.getByText('Page:', { exact: false })).toBeVisible();

  const pageField = editor.locator(
    '.work-document-field[data-field-kind="page"]',
  );
  await expect(pageField).toHaveAttribute('data-field-display', 'I');
  await expect(pageField).toHaveAttribute(
    'data-field-instruction',
    'PAGE \\* ROMAN \\* MERGEFORMAT',
  );
  await expect(pageField).toHaveAttribute('aria-label', '当前页码');

  const totalPagesField = editor.locator(
    '.work-document-field[data-field-kind="numPages"]',
  );
  await expect(totalPagesField).toHaveAttribute('data-field-display', 'A');
  await expect(totalPagesField).toHaveAttribute(
    'data-field-instruction',
    'NUMPAGES \\* ALPHABETIC \\* MERGEFORMAT',
  );
  await expect(totalPagesField).toHaveAttribute('aria-label', '总页数');

  const sectionField = editor.locator(
    '.work-document-field[data-field-kind="section"]',
  );
  await expect(sectionField).toHaveAttribute('data-field-display', '1st');
  await expect(sectionField).toHaveAttribute('aria-label', '当前节号');

  const referenceField = editor.locator(
    '.work-document-field[data-field-kind="pageReference"]',
  );
  await expect(referenceField).toHaveAttribute('data-field-display', '1st');
  await expect(referenceField).toHaveAttribute(
    'data-field-target-name',
    'Fields_target',
  );
  await expect(referenceField).toHaveAttribute('aria-label', '目标页码');

  await expect(page.locator('.work-document-page-stack')).toHaveAttribute(
    'data-page-count',
    '1',
  );
  await editor.focus();
  await page.keyboard.press('F9');
  await expect(editor).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-wps-numeric-fields-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});
