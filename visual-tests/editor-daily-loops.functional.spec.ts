import { readFile } from 'node:fs/promises';
import {
  type Download,
  expect,
  type Page,
  type TestInfo,
  test,
} from '@playwright/test';
import JSZip from 'jszip';
import { PDFDocument } from 'pdf-lib';
import * as XLSX from 'xlsx';
import { openPdfFixture, waitForPdfFixture } from './pdf-test-support';

// Each editor's create → edit → save → reopen contract. The downloaded bytes
// are checked independently of the editor before the file is reopened.

function desktopOnly(testInfo: TestInfo) {
  test.skip(testInfo.project.name !== 'desktop-1280');
  test.setTimeout(180_000);
}

test('Writer daily loop survives a DOCX save and reopen', async ({
  page,
}, testInfo) => {
  desktopOnly(testInfo);
  const browserErrors = collectBrowserErrors(page);
  await page.goto('/playground/');
  await page.locator("button[data-template-id='blank-document']").click();
  const body = page.locator(
    ".work-document-editable .ProseMirror[contenteditable='true']",
  );
  await expect(body).toBeFocused();

  await page.keyboard.type('WPSDAILY alpha token');
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+End');
  await page.getByRole('button', { name: '加粗' }).click();
  await page.getByRole('button', { name: '增大字号' }).click();
  await page.getByRole('button', { name: '居中' }).click();
  const centered = body.locator("p[style*='text-align: center']").first();
  await expect(centered.locator("[style*='font-size: 12pt']")).toBeVisible();
  await expect(centered.locator('strong')).toBeVisible();

  await page.getByRole('button', { name: '替换' }).click();
  const findPanel = page.locator('.work-document-find-panel');
  await findPanel.locator("input[aria-label='查找内容']").fill('alpha');
  await findPanel.locator("input[aria-label='替换为']").fill('beta');
  await expect(page.getByText('1 个匹配')).toBeVisible();
  await page.getByRole('button', { name: '全部替换' }).click();
  await expect(centered).toHaveText('WPSDAILY beta token');
  await page.getByRole('button', { name: '关闭查找' }).click();

  await centered.locator('strong').dblclick();
  await page.getByRole('tab', { name: '审阅' }).click();
  await page.getByRole('button', { name: '添加批注' }).click();
  const composer = page.locator('.work-document-comment-composer');
  await composer
    .locator("textarea[aria-label='批注内容']")
    .fill('WPSDAILY review note');
  await composer.locator('footer button:last-child').click();
  await expect(page.getByText('WPSDAILY review note')).toBeVisible();

  // Append the table after the commented paragraph so export keeps one body
  // paragraph, then a table (not a split run around the table).
  await body.evaluate((root) => {
    const selection = window.getSelection();
    if (!selection) return;
    const range = document.createRange();
    range.selectNodeContents(root);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
  });
  await settleSelection(page);
  await page.keyboard.press('Enter');
  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '插入表格' }).click();
  await page.getByRole('button', { name: '2 行 2 列' }).click();
  await expect(body.locator('table')).toBeVisible();

  const docx = await exportFile(page, '无标题文字.docx', async () => {
    await page.getByRole('button', { name: '导出' }).click();
    await page
      .locator(".work-export-menu-panel [role='menuitem']")
      .first()
      .click();
  });
  const archive = await JSZip.loadAsync(await downloadBytes(docx));
  const documentXml =
    (await archive.file('word/document.xml')?.async('text')) ?? '';
  expect(docxBodyOutline(documentXml)).toEqual([
    'WPSDAILY beta token',
    '<table>',
  ]);
  expect(await archive.file('word/comments.xml')?.async('text')).toContain(
    'WPSDAILY review note',
  );

  await reopenExportedFile(page, '无标题文字.docx');
  const reopened = page.locator('.work-document-editable .ProseMirror');
  const reopenedCentered = reopened
    .locator("p[style*='text-align: center']")
    .first();
  await expect(reopenedCentered).toHaveText('WPSDAILY beta token');
  await expect(
    reopenedCentered.locator("[style*='font-size: 12pt']").first(),
  ).toBeVisible();
  await expect(reopenedCentered.locator('strong').first()).toBeVisible();
  await expect(reopened.locator('table')).toBeVisible();
  await page.getByRole('tab', { name: '审阅' }).click();
  await page.locator("button[aria-label^='查看批注']").click();
  await expect(page.getByText('WPSDAILY review note')).toBeVisible();
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet daily loop survives an XLSX save and reopen', async ({
  page,
}, testInfo) => {
  desktopOnly(testInfo);
  const browserErrors = collectBrowserErrors(page);
  await page.goto('/playground/');
  await page.locator("button[data-template-id='blank-spreadsheet']").click();
  const grid = page.locator('.work-spreadsheet-editor .fortune-sheet-overlay');
  await expect(grid).toBeVisible();
  const formulaBar = page.locator('.fortune-fx-input');
  const selection = page.locator("output[aria-label='表格选区状态']");
  const result = page.locator("output[aria-label='当前单元格结果']");

  const nameBox = page.locator('.fortune-name-box');
  await grid.focus();
  await page.keyboard.press('Control+Home');
  for (const [address, value] of [
    ['A1', '10'],
    ['A3', '20'],
    ['A5', '=SUM(A1:A3)'],
  ] as const) {
    for (
      let step = 0;
      step < 2 && ((await nameBox.textContent()) ?? '').trim() !== address;
      step += 1
    ) {
      await page.keyboard.press('ArrowDown');
    }
    await expect(nameBox).toHaveText(address);
    await page.keyboard.type(value);
    await page.keyboard.press('Enter');
    await expect(nameBox).not.toHaveText(address);
  }
  await page.keyboard.press('ArrowUp');
  await expect(nameBox).toHaveText('A5');
  await expect(formulaBar).toHaveText('=SUM(A1:A3)');
  await expect(result).toHaveText('30');
  const sumAddress = (await selection.textContent())?.trim() ?? '';

  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  await page
    .locator(".work-spreadsheet-ribbon [role='tab'][data-tab-id='data']")
    .click();
  await page.getByRole('button', { name: '自动筛选' }).click();
  await expect(
    page.locator(".work-spreadsheet-editor[data-auto-filter='active']"),
  ).toBeVisible();

  const xlsx = await exportFile(page, '无标题表格.xlsx', () =>
    page.getByRole('button', { name: '导出' }).click(),
  );
  const workbook = XLSX.read(await downloadBytes(xlsx), {
    cellFormula: true,
    cellStyles: true,
  });
  const sheet = workbook.Sheets[workbook.SheetNames[0] ?? ''];
  const sumCell = Object.entries(sheet ?? {}).find(
    ([address, cell]) =>
      !address.startsWith('!') &&
      typeof cell === 'object' &&
      String((cell as XLSX.CellObject).f ?? '').toUpperCase() === 'SUM(A1:A3)',
  );
  expect(sumCell?.[0]).toBe(sumAddress || 'A5');
  expect((sumCell?.[1] as XLSX.CellObject | undefined)?.v).toBe(30);
  expect(sheet?.['!autofilter']).toBeTruthy();

  await reopenExportedFile(page, '无标题表格.xlsx');
  await expect(grid).toBeVisible();
  await expect(
    page.locator(".work-spreadsheet-editor[data-auto-filter='active']"),
  ).toBeVisible();
  await grid.focus();
  await page.keyboard.press('Control+Home');
  await expect(formulaBar).toHaveText('10');
  await expect(nameBox).toHaveText('A1');
  // Escape dismisses filter chrome; Fortune parks focus in the hidden cell
  // editor unless we restore the overlay. ArrowDown must still reach A2.
  await page.keyboard.press('Escape');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('A2');
  expect(browserErrors).toEqual([]);
});

test('Presentation daily loop survives a PPTX save and reopen', async ({
  page,
}, testInfo) => {
  desktopOnly(testInfo);
  const browserErrors = collectBrowserErrors(page);
  await page.goto('/playground/');
  await page.locator("button[data-template-id='blank-presentation']").click();
  const strip = page.locator('.work-slide-strip');
  await expect(
    strip.locator("[data-slide-thumbnail][data-slide-index='0']"),
  ).toBeVisible();

  await strip.locator('[data-slide-thumbnail].active').click();
  // Meta+M is intercepted by macOS; use the ribbon command instead.
  await page.getByRole('button', { name: '新建幻灯片' }).click();
  await expect(
    strip.locator("[data-slide-thumbnail][data-slide-index='1'].active"),
  ).toBeVisible();
  await page
    .locator('.work-slide-canvas.interactive .work-slide-placeholder-text')
    .first()
    .dblclick();
  const textEngine = page.locator(
    ".work-slide-canvas.interactive [data-presentation-text-engine='tiptap']",
  );
  await expect(textEngine).toBeFocused();
  await page.keyboard.type('WPSDAILY slide text');
  await expect(textEngine).toHaveText('WPSDAILY slide text');
  await page.keyboard.press('Escape');

  await page
    .locator("[aria-label='演讲者备注']")
    .fill('WPSDAILY speaker notes');
  await expect(page.getByText('已添加演讲者备注')).toBeVisible();
  await strip.locator('[data-slide-thumbnail].active').click();
  await page.getByRole('button', { name: '复制幻灯片' }).click();
  await expect(
    strip.locator("[data-slide-thumbnail][data-slide-index='2'].active"),
  ).toBeVisible();
  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '形状' }).click();
  await expect(
    page.locator('.work-slide-canvas.interactive .work-slide-element.shape'),
  ).toBeVisible();
  await page.keyboard.press('Escape');

  const pptx = await exportFile(page, '无标题演示.pptx', () =>
    page.locator("button.work-export-button[aria-label='导出']").click(),
  );
  const archive = await JSZip.loadAsync(await downloadBytes(pptx));
  const slides = Object.keys(archive.files).filter((name) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(name),
  );
  expect(slides).toHaveLength(3);
  const notes = await Promise.all(
    Object.keys(archive.files)
      .filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name))
      .map((name) => archive.file(name)?.async('text')),
  );
  expect(
    notes.filter((xml) => xml?.includes('WPSDAILY speaker notes')),
  ).toHaveLength(2);

  await reopenExportedFile(page, '无标题演示.pptx');
  for (const index of ['1', '2']) {
    await strip
      .locator(`[data-slide-thumbnail][data-slide-index='${index}']`)
      .click();
    await expect(
      page.locator('.work-slide-canvas.interactive .work-slide-rich-text', {
        hasText: 'WPSDAILY slide text',
      }),
    ).toBeVisible();
    await expect(page.getByText('已添加演讲者备注')).toBeVisible();
  }
  expect(browserErrors).toEqual([]);
});

test('PDF daily loop survives an organize, save, and reopen', async ({
  page,
}, testInfo) => {
  desktopOnly(testInfo);
  const browserErrors = collectBrowserErrors(page);
  await page.goto('/playground/');
  await openPdfFixture(page, { pageCount: 4 });
  await waitForPdfFixture(page);
  const rail = page.locator('.work-pdf-thumbnail-rail');
  await expect(rail).toHaveAttribute('data-pdf-page-count', '4');

  await page
    .locator("button[aria-label='组织 PDF 页面']:not(:disabled)")
    .click();
  await page.getByRole('button', { name: '插入空白页' }).click();
  await expect(rail).toHaveAttribute('data-pdf-page-count', '5');

  const pdf = await exportFile(page, /\.pdf 已下载/, () =>
    page.locator("button.work-export-button[aria-label='下载 PDF']").click(),
  );
  const saved = await PDFDocument.load(await downloadBytes(pdf));
  expect(saved.getPageCount()).toBe(5);

  await page.getByRole('button', { name: '重新打开已导出文件' }).click();
  await expect(page.getByText(/\.pdf 已打开/)).toBeVisible();
  await expect(page.locator('.work-pdf-embed')).toHaveAttribute(
    'data-ready',
    'true',
    { timeout: 50_000 },
  );
  await expect(rail).toHaveAttribute('data-pdf-page-count', '5');
  expect(browserErrors).toEqual([]);
});

test('Markdown daily loop survives an MD save and reopen', async ({
  page,
}, testInfo) => {
  desktopOnly(testInfo);
  const browserErrors = collectBrowserErrors(page);
  await page.goto('/playground/');
  await page.getByText('产品说明', { exact: true }).first().click();
  const source = page.locator('.work-markdown-pane.source textarea');
  const preview = page.locator('.work-markdown-pane.visual');
  await expect(source).toBeVisible();

  await source.fill('# MDDAILY heading\n\n**MDDAILY alpha**\n\n| Item | Owner |\n| --- | --- |\n| MDDAILY row | Ops |');
  await expect(preview.locator('strong')).toHaveText('MDDAILY alpha');
  await expect(preview.locator('table')).toBeVisible();

  const markdown = await exportFile(page, '产品说明.md', () =>
    page.getByRole('button', { name: '导出' }).click(),
  );
  const saved = new TextDecoder().decode(await downloadBytes(markdown));
  expect(saved).toContain('# MDDAILY heading');
  expect(saved).toContain('**MDDAILY alpha**');
  expect(saved).toMatch(/\|\s*MDDAILY row\s*\|\s*Ops\s*\|/);

  await reopenExportedFile(page, '产品说明.md');
  await expect(preview.locator('h1')).toHaveText('MDDAILY heading');
  await expect(preview.locator('strong')).toHaveText('MDDAILY alpha');
  await expect(preview.getByText('MDDAILY row')).toBeVisible();
  expect(browserErrors).toEqual([]);
});

async function exportFile(
  page: Page,
  downloadedNotice: string | RegExp,
  trigger: () => Promise<unknown>,
): Promise<Download> {
  const download = page.waitForEvent('download');
  await trigger();
  const file = await download;
  await expect(
    page.getByText(
      typeof downloadedNotice === 'string'
        ? `${downloadedNotice} 已下载`
        : downloadedNotice,
    ),
  ).toBeVisible();
  return file;
}

/** Lets the browser deliver `selectionchange` so the editor adopts the caret. */
async function settleSelection(page: Page) {
  await page.evaluate(
    () =>
      new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
      ),
  );
}

/** Top-level body blocks: paragraph text with runs joined, or a table marker. */
function docxBodyOutline(documentXml: string): string[] {
  const body = documentXml.slice(
    documentXml.indexOf('<w:body>'),
    documentXml.indexOf('</w:body>'),
  );
  const outline: string[] = [];
  for (const [block] of body.matchAll(
    /<w:tbl>[\s\S]*?<\/w:tbl>|<w:p[ >][\s\S]*?<\/w:p>/g,
  )) {
    if (block.startsWith('<w:tbl>')) {
      outline.push('<table>');
      continue;
    }
    const text = Array.from(block.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g))
      .map((match) => match[1])
      .join('');
    if (text) outline.push(text);
  }
  return outline;
}

async function reopenExportedFile(page: Page, fileName: string) {
  await page.getByRole('button', { name: '重新打开已导出文件' }).click();
  await expect(page.getByText(`${fileName} 已打开`)).toBeVisible();
}

async function downloadBytes(download: Download): Promise<Uint8Array> {
  const path = await download.path();
  if (!path) throw new Error('The download did not produce a local file.');
  return new Uint8Array(await readFile(path));
}

function collectBrowserErrors(page: Page): string[] {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });
  return errors;
}
