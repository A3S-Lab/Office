import { expect, type Page, test } from '@playwright/test';

async function openWpsDocument(
  page: Page,
  fixture: string,
  readyText: string,
): Promise<void> {
  await page.goto('/playground/');
  await page
    .locator('input[aria-label="打开 Office 或 PDF 文件"]')
    .setInputFiles(`.a3s-test/fixtures/${fixture}`);
  await expect(page.getByText(readyText, { exact: false })).toBeVisible();
  await expect(
    page.locator('.work-document-editable .ProseMirror'),
  ).toHaveAttribute('data-pagination-state', 'ready');
}

function collectBrowserErrors(page: Page): string[] {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });
  return errors;
}

test('Writer preserves imported WPS page layout and pagination metadata', async ({
  page,
}, testInfo) => {
  const browserErrors = collectBrowserErrors(page);
  await openWpsDocument(
    page,
    'word-wps-layout.docx',
    'Quarterly revenue report',
  );

  const editor = page.locator('.work-document-editable .ProseMirror');
  await expect(editor).toHaveAttribute('data-pagination-engine', 'wasm');
  await expect(editor).toHaveAttribute('data-pagination-text-engine', 'dom');
  await expect(page.locator('.work-document-page-stack')).toHaveAttribute(
    'data-page-count',
    '1',
  );
  await expect(
    page.locator('.work-document-page[data-pdf-page-count="1"]'),
  ).toHaveAttribute('data-pdf-page-size', 'a4');
  await expect(
    page.locator('.work-document-page[data-pdf-page-count="1"]'),
  ).toHaveAttribute('data-pdf-orientation', 'portrait');
  await expect(
    editor
      .locator(
        'p[data-office-line-rule="auto"][data-office-auto-line-height="1.84"]',
      )
      .first(),
  ).toBeVisible();
  await expect(
    editor
      .locator(
        'p[data-office-line-rule="auto"][data-office-auto-line-height="1.2075"]',
      )
      .first(),
  ).toBeVisible();
  await expect(
    page.locator('.tableWrapper > table[data-office-table-imported="true"]'),
  ).toBeVisible();
  await page.screenshot({
    path: testInfo.outputPath(`writer-wps-layout-${testInfo.project.name}.png`),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});

test('Writer keeps WPS Latin and CJK font metrics on one page', async ({
  page,
}, testInfo) => {
  const browserErrors = collectBrowserErrors(page);
  await openWpsDocument(
    page,
    'word-wps-font-matrix.docx',
    'Microsoft YaHei 16pt 3 line 360',
  );

  const editor = page.locator('.work-document-editable .ProseMirror');
  await expect(page.locator('.work-document-page-stack')).toHaveAttribute(
    'data-page-count',
    '1',
  );
  await expect(
    editor.locator('[data-office-word-line-height-factor="1.2207"]').first(),
  ).toBeVisible();
  await expect(
    editor.locator('[data-office-word-line-height-factor="1.3301"]').first(),
  ).toBeVisible();
  await expect(
    editor.locator('[data-office-word-line-height-factor="1.7143"]').first(),
  ).toBeVisible();
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-wps-font-matrix-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});

test('Writer preserves the WPS document-grid line pitch', async ({
  page,
}, testInfo) => {
  const browserErrors = collectBrowserErrors(page);
  await openWpsDocument(
    page,
    'word-wps-grid-matrix.docx',
    'Microsoft YaHei 16pt grid row 3',
  );

  const section = page.locator(
    '.work-document-section[data-section-document-grid-type="lines"]',
  );
  await expect(section).toHaveAttribute(
    'data-section-document-grid-line-pitch',
    '12',
  );
  await expect(section).toHaveAttribute(
    'style',
    /--work-document-word-grid-line-pitch/,
  );
  await expect(page.locator('.work-document-page-stack')).toHaveAttribute(
    'data-page-count',
    '1',
  );
  await page.screenshot({
    path: testInfo.outputPath(`writer-wps-grid-${testInfo.project.name}.png`),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});

test('Writer preserves WPS mixed-script direction and line metrics', async ({
  page,
}, testInfo) => {
  const browserErrors = collectBrowserErrors(page);
  await openWpsDocument(page, 'word-wps-script-matrix.docx', 'Mixed 6');

  const editor = page.locator('.work-document-editable .ProseMirror');
  await expect(editor.locator('p[dir="rtl"]').first()).toBeVisible();
  await expect(
    editor
      .locator('[data-office-word-line-height-factor="1.15"][style*="Arial"]')
      .first(),
  ).toBeVisible();
  await expect(
    editor
      .locator(
        '[data-office-word-line-height-factor="1.3301"][style*="Segoe UI"]',
      )
      .first(),
  ).toBeVisible();
  await expect(
    editor
      .locator(
        '[data-office-word-line-height-factor="1.7143"][style*="Microsoft YaHei"]',
      )
      .first(),
  ).toBeVisible();
  await expect(editor.locator('span strong').first()).toBeVisible();
  await expect(editor.locator('span em').first()).toBeVisible();
  await expect(page.locator('.work-document-page-stack')).toHaveAttribute(
    'data-page-count',
    '1',
  );
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-wps-script-matrix-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});
