import { readFile } from 'node:fs/promises';
import { expect, type Download, type Page, test } from '@playwright/test';
import { PDFDocument } from 'pdf-lib';
import {
  createPdfFixture,
  openPdfFixture,
  waitForPdfFixture,
} from './pdf-test-support';

test('PDF page organization mutates, exports, saves, and restores exact history', async ({
  page,
}, testInfo) => {
  test.skip(testInfo.project.name !== 'desktop-1280');
  test.setTimeout(180_000);
  const browserErrors = collectBrowserErrors(page);

  await page.goto('/playground/');
  await openPdfFixture(page, { pageCount: 4 });
  await waitForPdfFixture(page);
  await expectPdfPageCount(page, 4);

  let organizer = await openPageOrganizer(page);
  await expect(organizer.getByText('已选择 1 / 4 页')).toBeVisible();
  await organizer.getByRole('button', { name: '插入空白页' }).click();
  await expectPdfPageCount(page, 5);

  await runPdfToolbarAction(page, '撤销');
  await expectPdfPageCount(page, 4);
  await runPdfToolbarAction(page, '重做');
  await expectPdfPageCount(page, 5);

  organizer = await openPageOrganizer(page);
  await organizer.getByRole('button', { name: '向右旋转所选页' }).click();
  await expectPdfPageCount(page, 5);

  organizer = await openPageOrganizer(page);
  await organizer.getByRole('button', { name: '选择第 2 页' }).click();
  await organizer.getByRole('button', { name: '删除所选页' }).click();
  await expectPdfPageCount(page, 4);

  organizer = await openPageOrganizer(page);
  await organizer
    .getByRole('button', { name: '选择第 1 页' })
    .dragTo(organizer.getByRole('button', { name: '选择第 3 页' }));
  await expectPdfPageCount(page, 4);
  expect(await pdfPageSummary(await savePdfAndDownload(page))).toEqual({
    pageCount: 4,
    rotations: [0, 90, 0, 0],
  });
  await page.getByRole('textbox', { name: '页码' }).fill('2');
  await page.getByRole('textbox', { name: '页码' }).press('Enter');
  await expect(page.getByRole('textbox', { name: '页码' })).toHaveValue('2');

  organizer = await openPageOrganizer(page);
  const mergeChooser = page.waitForEvent('filechooser');
  await organizer.getByRole('button', { name: '合并另一个 PDF' }).click();
  await (await mergeChooser).setFiles({
    name: 'merge-fixture.pdf',
    mimeType: 'application/pdf',
    buffer: createPdfFixture(2),
  });
  await expectPdfPageCount(page, 6);

  await runPdfToolbarAction(page, '撤销');
  await expectPdfPageCount(page, 4);
  await runPdfToolbarAction(page, '重做');
  await expectPdfPageCount(page, 6);

  organizer = await openPageOrganizer(page);
  await organizer.getByRole('button', { name: '选择第 2 页' }).click();
  await page.screenshot({
    path: testInfo.outputPath('pdf-page-organization-desktop.png'),
    fullPage: false,
  });
  const extractedDownload = page.waitForEvent('download');
  await organizer.getByRole('button', { name: '抽取所选页' }).click();
  expect(await pdfPageSummary(await extractedDownload)).toEqual({
    pageCount: 1,
    rotations: [90],
  });

  const splitDownloads: Download[] = [];
  const collectSplitDownload = (download: Download) =>
    splitDownloads.push(download);
  page.on('download', collectSplitDownload);
  await organizer.getByRole('button', { name: '在所选页后拆分' }).click();
  await expect.poll(() => splitDownloads.length).toBe(2);
  page.off('download', collectSplitDownload);
  const splitSummaries = await Promise.all(splitDownloads.map(pdfPageSummary));
  expect(splitSummaries.map(({ pageCount }) => pageCount).sort()).toEqual([
    2, 4,
  ]);

  await organizer.getByRole('button', { name: '完成' }).click();
  await expect(
    page.getByRole('button', { name: '组织 PDF 页面' }),
  ).toBeFocused();

  expect(await pdfPageSummary(await savePdfAndDownload(page))).toEqual({
    pageCount: 6,
    rotations: [0, 90, 0, 0, 0, 0],
  });

  expect(browserErrors).toEqual([]);
});

test('compact PDF page organization opens through overflow and restores focus', async ({
  page,
}, testInfo) => {
  test.skip(testInfo.project.name !== 'compact-768');
  const browserErrors = collectBrowserErrors(page);
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/playground/');
  await openPdfFixture(page, { pageCount: 3 });
  await waitForPdfFixture(page);

  const directTrigger = page.getByRole('button', { name: '组织 PDF 页面' });
  await expect(directTrigger).toBeHidden();
  const organizer = await openPageOrganizer(page);
  await expect(organizer).toBeVisible();
  await expect(
    organizer.getByRole('button', { name: '向右旋转所选页' }),
  ).toBeVisible();
  await expect(
    organizer.getByRole('button', { name: '删除所选页' }),
  ).toBeVisible();
  await page.screenshot({
    path: testInfo.outputPath('pdf-page-organization-compact.png'),
    fullPage: false,
  });
  await organizer.getByRole('button', { name: '完成' }).click();

  const overflowTrigger = page.getByRole('button', { name: '更多 PDF 工具' });
  await expect(overflowTrigger).toBeFocused();
  expect(browserErrors).toEqual([]);
});

async function openPageOrganizer(page: Page) {
  const directTrigger = page.getByRole('button', { name: '组织 PDF 页面' });
  if (await directTrigger.isVisible()) {
    await directTrigger.click();
  } else {
    await page.getByRole('button', { name: '更多 PDF 工具' }).click();
    await page
      .getByRole('menu', { name: '更多 PDF 工具' })
      .getByRole('menuitem', { name: '组织页面' })
      .click();
  }
  const dialog = page.getByRole('dialog', { name: '组织 PDF 页面' });
  await expect(dialog).toBeVisible();
  return dialog;
}

async function runPdfToolbarAction(page: Page, name: string): Promise<void> {
  const directAction = page.getByRole('button', { name, exact: true });
  if (await directAction.isVisible()) {
    await directAction.click();
    return;
  }
  await page.getByRole('button', { name: '更多 PDF 工具' }).click();
  await page
    .getByRole('menu', { name: '更多 PDF 工具' })
    .getByRole('menuitem', { name, exact: true })
    .click();
}

async function expectPdfPageCount(
  page: Page,
  pageCount: number,
): Promise<void> {
  await expect(page.locator('.work-pdf-page-total')).toHaveText(
    `/ ${pageCount}`,
    {
      timeout: 50_000,
    },
  );
  await expect(page.locator('.work-pdf-embed')).toHaveAttribute(
    'data-ready',
    'true',
    { timeout: 50_000 },
  );
}

async function savePdfAndDownload(page: Page): Promise<Download> {
  await page.getByRole('button', { name: '保存' }).click();
  await expect(page.getByText('PDF 批注已保存到当前浏览器会话')).toBeVisible();
  await expect(page.locator('.work-pdf-embed')).toHaveAttribute(
    'data-ready',
    'true',
    { timeout: 50_000 },
  );
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: '下载 PDF' }).click();
  return download;
}

async function pdfPageSummary(download: Download): Promise<{
  pageCount: number;
  rotations: number[];
}> {
  const downloadPath = await download.path();
  if (!downloadPath)
    throw new Error('PDF download did not produce a local file.');
  const pdf = await PDFDocument.load(await readFile(downloadPath));
  return {
    pageCount: pdf.getPageCount(),
    rotations: pdf.getPages().map((page) => page.getRotation().angle),
  };
}

function collectBrowserErrors(page: Page): string[] {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });
  return errors;
}
