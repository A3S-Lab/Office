import { expect, type Locator, type Page, test } from '@playwright/test';
import { openPdfFixture, waitForPdfFixture } from './pdf-test-support';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

interface WheelZoomFixture {
  name: string;
  open: (page: Page) => Promise<void>;
  ready: (page: Page) => Promise<void>;
  surface: (page: Page) => Locator;
  zoomLabel: string;
}

const fixtures: WheelZoomFixture[] = [
  {
    name: 'document',
    open: openDocumentFixture,
    ready: waitForDocumentFixture,
    surface: (page) => page.locator('.work-document-scroll'),
    zoomLabel: '文档缩放比例',
  },
  {
    name: 'markdown',
    open: (page) =>
      page.getByRole('button', { name: '# 产品说明 MD · 本次会话' }).click(),
    ready: (page) =>
      page.locator('.work-markdown-editor .ProseMirror').waitFor(),
    surface: (page) => page.locator('.work-markdown-workspace'),
    zoomLabel: 'Markdown 缩放比例',
  },
  {
    name: 'spreadsheet',
    open: (page) =>
      page
        .getByRole('button', { name: '季度执行计划 XLSX · 本次会话' })
        .click(),
    ready: (page) =>
      page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor(),
    surface: (page) => page.locator('.fortune-sheet-overlay'),
    zoomLabel: '表格缩放比例',
  },
  {
    name: 'presentation',
    open: (page) =>
      page
        .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
        .click(),
    ready: async (page) => {
      await page.locator('.work-slide-canvas.interactive').waitFor();
      await expect(page.locator('.work-presentation-editor')).toHaveAttribute(
        'data-presentation-geometry-state',
        'idle',
      );
    },
    surface: (page) => page.locator('.work-slide-stage'),
    zoomLabel: '演示缩放比例',
  },
  {
    name: 'pdf',
    open: openPdfFixture,
    ready: waitForPdfFixture,
    surface: (page) => page.locator('.work-pdf-embed'),
    zoomLabel: 'PDF 缩放比例',
  },
];

test('every editor routes Ctrl/Cmd + wheel to its own bounded zoom', async ({
  page,
}, testInfo) => {
  test.skip(
    testInfo.project.name !== 'desktop-1280',
    'The gesture contract only needs one desktop browser project.',
  );

  for (const fixture of fixtures) {
    await test.step(fixture.name, async () => {
      await page.goto('/');
      await fixture.open(page);
      await fixture.ready(page);

      const output = page.locator(`output[aria-label="${fixture.zoomLabel}"]`);
      await expect(output).toBeVisible();
      const initialZoom = await readZoom(output);
      const initialBrowserScale = await browserScale(page);

      await modifiedWheel(page, fixture.surface(page), 'Control', -120);
      await expect.poll(() => readZoom(output)).toBeGreaterThan(initialZoom);
      expect(await browserScale(page)).toEqual(initialBrowserScale);

      const zoomedIn = await readZoom(output);
      await modifiedWheel(page, fixture.surface(page), 'Meta', 120);
      await expect.poll(() => readZoom(output)).toBeLessThan(zoomedIn);
      expect(await browserScale(page)).toEqual(initialBrowserScale);
    });
  }
});

async function modifiedWheel(
  page: Page,
  surface: Locator,
  modifier: 'Control' | 'Meta',
  deltaY: number,
): Promise<void> {
  await surface.hover();
  await page.keyboard.down(modifier);
  try {
    await page.mouse.wheel(0, deltaY);
  } finally {
    await page.keyboard.up(modifier);
  }
}

async function readZoom(output: Locator): Promise<number> {
  const value = Number.parseInt((await output.textContent()) ?? '', 10);
  if (!Number.isFinite(value)) throw new Error('Editor zoom is unavailable.');
  return value;
}

function browserScale(page: Page): Promise<{ dpr: number; viewport: number }> {
  return page.evaluate(() => ({
    dpr: window.devicePixelRatio,
    viewport: window.visualViewport?.scale ?? 1,
  }));
}
