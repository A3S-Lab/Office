import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  type Browser,
  type CDPSession,
  chromium,
  type Page,
} from '@playwright/test';
import {
  browserLongTasks,
  installPerformanceObserver,
  resetBrowserMeasurements,
  waitForVisualQuiet,
} from './benchmark-browser';
import {
  collectMetrics,
  type LongTaskSummary,
  type MetricSnapshot,
  messageOf,
} from './benchmark-metrics';

type BenchmarkMode = 'cold' | 'preloaded';

interface PdfMilestones {
  firstBitmapMs: number;
  readyMs: number;
  shellMountedMs: number;
  thumbnailReadyMs: number;
}

interface PdfSurfaceStats {
  currentPage: number | null;
  loadedBitmapCount: number;
  mountedPageCount: number;
  mountedThumbnailCount: number;
  readyThumbnailCount: number;
  thumbnailWindowEnd: number | null;
  thumbnailWindowStart: number | null;
  totalPages: number | null;
}

interface PdfBenchmarkRun {
  browserErrors: string[];
  consoleErrors: string[];
  failure?: string;
  importLongTasks?: LongTaskSummary;
  jumpLongTasks?: LongTaskSummary;
  jumpMs?: number;
  metrics?: {
    afterImport: MetricSnapshot;
    beforeImport: MetricSnapshot;
    retained: MetricSnapshot;
  };
  milestones?: PdfMilestones;
  mode: BenchmarkMode;
  run: number;
  submittedAt: string;
  surfaceAfterImport?: PdfSurfaceStats;
  surfaceAfterJump?: PdfSurfaceStats;
  workers: string[];
}

interface PdfBrowserBenchmarkState {
  firstBitmapMs: number | null;
  jumpMs: number | null;
  readyMs: number | null;
  shellMountedMs: number | null;
  startedAt: number | null;
  thumbnailReadyMs: number | null;
}

declare global {
  interface Window {
    __a3sPdfBenchmark?: PdfBrowserBenchmarkState;
  }
}

const targetUrl = argument('--url') ?? 'http://127.0.0.1:4175/playground/';
const repeatCount = positiveInteger(argument('--runs'), 3);
const timeoutMs = positiveInteger(argument('--timeout-ms'), 120_000);
const fixture = argument('--fixture') ?? './fixtures/pdf-pages-1000.pdf';
const requestedModes = (argument('--modes') ?? 'cold,preloaded')
  .split(',')
  .map((value) => value.trim())
  .filter((value): value is BenchmarkMode =>
    ['cold', 'preloaded'].includes(value),
  );

if (!requestedModes.length) throw new Error('No valid PDF modes selected.');

const timestamp = new Date().toISOString().replaceAll(':', '-');
const resultDirectory = fileURLToPath(new URL('./results/', import.meta.url));
const resultPath = resolve(resultDirectory, `${timestamp}-pdf.json`);
const report: {
  completedAt?: string;
  environment: Record<string, unknown>;
  runs: PdfBenchmarkRun[];
  startedAt: string;
} = {
  environment: {
    fixture,
    headless: true,
    modes: requestedModes,
    playwright: '1.61.1',
    repeatCount,
    targetUrl,
    timeoutMs,
  },
  runs: [],
  startedAt: new Date().toISOString(),
};

await mkdir(resultDirectory, { recursive: true });

for (const mode of requestedModes) {
  for (let run = 1; run <= repeatCount; run += 1) {
    console.log(JSON.stringify({ event: 'run-start', mode, run }));
    const result = await benchmarkPdf(mode, run);
    report.runs.push(result);
    await saveReport();
    console.log(
      JSON.stringify({
        event: 'run-finish',
        failure: result.failure ?? null,
        milestones: result.milestones ?? null,
        mode,
        run,
      }),
    );
  }
}

report.completedAt = new Date().toISOString();
await saveReport();
console.log(JSON.stringify({ event: 'complete', resultPath }));

async function benchmarkPdf(
  mode: BenchmarkMode,
  run: number,
): Promise<PdfBenchmarkRun> {
  let browser: Browser | undefined;
  let cdp: CDPSession | undefined;
  const result: PdfBenchmarkRun = {
    browserErrors: [],
    consoleErrors: [],
    mode,
    run,
    submittedAt: new Date().toISOString(),
    workers: [],
  };

  try {
    browser = await chromium.launch({
      args: [
        '--disable-background-timer-throttling',
        '--disable-renderer-backgrounding',
        '--enable-precise-memory-info',
      ],
      headless: true,
    });
    const context = await browser.newContext({
      deviceScaleFactor: 1,
      viewport: { height: 1_000, width: 1_440 },
    });
    const page = await context.newPage();
    page.on('console', (message) => {
      if (message.type() === 'error') result.consoleErrors.push(message.text());
    });
    page.on('crash', () => result.browserErrors.push('Renderer crashed.'));
    page.on('pageerror', (error) => result.browserErrors.push(error.message));
    page.on('worker', (worker) => result.workers.push(worker.url()));
    await installPerformanceObserver(page);
    cdp = await context.newCDPSession(page);
    await cdp.send('Performance.enable');
    await page.goto(targetUrl, {
      timeout: 120_000,
      waitUntil: 'domcontentloaded',
    });
    const fileInput = page.locator(
      'input[aria-label="打开 Office 或 PDF 文件"]',
    );
    await fileInput.waitFor({ state: 'attached', timeout: 60_000 });
    report.environment.browser ??= await page.evaluate(() => ({
      deviceMemory: (navigator as Navigator & { deviceMemory?: number })
        .deviceMemory,
      hardwareConcurrency: navigator.hardwareConcurrency,
      userAgent: navigator.userAgent,
    }));

    if (mode === 'preloaded') {
      await page.getByRole('button', { exact: true, name: 'PDF 打开' }).hover();
      await waitForPreloadMark(
        page,
        'a3s-office.playground.pdf.runtime-preload-ready',
      );
    }

    await installPdfMilestoneClock(page);
    const beforeImport = await collectMetrics(cdp);
    await resetBrowserMeasurements(page);
    await fileInput.setInputFiles(fixturePath(), { timeout: timeoutMs });
    await page.waitForFunction(
      () => {
        const state = window.__a3sPdfBenchmark;
        return Boolean(
          state?.firstBitmapMs !== null &&
            state?.readyMs !== null &&
            state?.thumbnailReadyMs !== null,
        );
      },
      undefined,
      { polling: 'raf', timeout: timeoutMs },
    );
    result.milestones = await readPdfMilestones(page);
    await waitForVisualQuiet(page);
    result.importLongTasks = await browserLongTasks(page);
    result.surfaceAfterImport = await collectPdfSurfaceStats(page);
    const afterImport = await collectMetrics(cdp);

    await resetBrowserMeasurements(page);
    await installPdfEndJumpClock(page);
    await page.locator('[data-pdf-page-index="0"]').focus();
    await page.keyboard.press('End');
    await page.waitForFunction(
      () => window.__a3sPdfBenchmark?.jumpMs !== null,
      undefined,
      { polling: 'raf', timeout: timeoutMs },
    );
    result.jumpMs = await page.evaluate(
      () => window.__a3sPdfBenchmark?.jumpMs ?? Number.NaN,
    );
    await page.waitForFunction(
      () => {
        const rail = document.querySelector<HTMLElement>(
          '.work-pdf-thumbnail-rail',
        );
        const pageCount = Number(rail?.dataset.pdfPageCount);
        return Number(rail?.dataset.pdfThumbnailWindowEnd) === pageCount;
      },
      undefined,
      { polling: 'raf', timeout: timeoutMs },
    );
    result.jumpLongTasks = await browserLongTasks(page);
    result.surfaceAfterJump = await collectPdfSurfaceStats(page);

    await cdp.send('HeapProfiler.collectGarbage');
    result.metrics = {
      afterImport,
      beforeImport,
      retained: await collectMetrics(cdp),
    };
  } catch (error) {
    result.failure = messageOf(error);
  } finally {
    await browser?.close().catch(() => undefined);
  }
  return result;
}

async function waitForPreloadMark(page: Page, name: string): Promise<void> {
  await page.waitForFunction(
    (entryName) => performance.getEntriesByName(entryName).length > 0,
    name,
    { polling: 'raf', timeout: timeoutMs },
  );
}

async function installPdfMilestoneClock(page: Page): Promise<void> {
  await page.evaluate(() => {
    const state: PdfBrowserBenchmarkState = {
      firstBitmapMs: null,
      jumpMs: null,
      readyMs: null,
      shellMountedMs: null,
      startedAt: null,
      thumbnailReadyMs: null,
    };
    window.__a3sPdfBenchmark = state;
    const input = document.querySelector<HTMLInputElement>(
      'input[aria-label="打开 Office 或 PDF 文件"]',
    );
    if (!input) throw new Error('The Office file input is unavailable.');
    input.addEventListener(
      'change',
      () => {
        state.startedAt = performance.now();
        const inspect = () => {
          const startedAt = state.startedAt;
          if (startedAt === null) return;
          const elapsed = performance.now() - startedAt;
          if (
            state.shellMountedMs === null &&
            document.querySelector('.work-pdf-viewer')
          ) {
            state.shellMountedMs = elapsed;
          }
          if (
            state.readyMs === null &&
            document.querySelector('.work-pdf-embed[data-ready="true"]')
          ) {
            state.readyMs = elapsed;
          }
          if (
            state.thumbnailReadyMs === null &&
            document.querySelector('[data-pdf-thumbnail-state="ready"]')
          ) {
            state.thumbnailReadyMs = elapsed;
          }
          if (state.firstBitmapMs === null && loadedPdfBitmaps().length > 0) {
            state.firstBitmapMs = elapsed;
          }
          if (
            state.readyMs !== null &&
            state.thumbnailReadyMs !== null &&
            state.firstBitmapMs !== null
          ) {
            return;
          }
          requestAnimationFrame(inspect);
        };
        requestAnimationFrame(inspect);
      },
      { capture: true, once: true },
    );

    function loadedPdfBitmaps(): HTMLImageElement[] {
      const root = document.querySelector('.work-pdf-native-viewer');
      if (!root) return [];
      return deepElements(root, 'img').filter(
        (element): element is HTMLImageElement =>
          element instanceof HTMLImageElement &&
          element.complete &&
          element.naturalWidth > 0,
      );
    }

    function deepElements(root: ParentNode, selector: string): Element[] {
      const matches = Array.from(root.querySelectorAll(selector));
      for (const element of root.querySelectorAll('*')) {
        if (element.shadowRoot) {
          matches.push(...deepElements(element.shadowRoot, selector));
        }
      }
      return [...new Set(matches)];
    }
  });
}

async function installPdfEndJumpClock(page: Page): Promise<void> {
  await page.evaluate(() => {
    const state = window.__a3sPdfBenchmark;
    const rail = document.querySelector<HTMLElement>(
      '.work-pdf-thumbnail-rail',
    );
    const firstPage = document.querySelector<HTMLElement>(
      '[data-pdf-page-index="0"]',
    );
    const totalPages = Number(rail?.dataset.pdfPageCount);
    if (!state || !firstPage || !Number.isFinite(totalPages)) {
      throw new Error('PDF jump instrumentation is unavailable.');
    }
    state.jumpMs = null;
    firstPage.addEventListener(
      'keydown',
      (event) => {
        if (event.key !== 'End') return;
        const startedAt = performance.now();
        const inspect = () => {
          const last = document.querySelector<HTMLElement>(
            `[data-pdf-page-index="${totalPages - 1}"]`,
          );
          if (last?.getAttribute('aria-current') === 'page') {
            state.jumpMs =
              Math.round((performance.now() - startedAt) * 10) / 10;
            return;
          }
          requestAnimationFrame(inspect);
        };
        requestAnimationFrame(inspect);
      },
      { capture: true, once: true },
    );
  });
}

async function readPdfMilestones(page: Page): Promise<PdfMilestones> {
  return page.evaluate(() => {
    const state = window.__a3sPdfBenchmark;
    if (
      !state ||
      state.firstBitmapMs === null ||
      state.readyMs === null ||
      state.shellMountedMs === null ||
      state.thumbnailReadyMs === null
    ) {
      throw new Error('PDF milestones are incomplete.');
    }
    return {
      firstBitmapMs: round(state.firstBitmapMs),
      readyMs: round(state.readyMs),
      shellMountedMs: round(state.shellMountedMs),
      thumbnailReadyMs: round(state.thumbnailReadyMs),
    };

    function round(value: number): number {
      return Math.round(value * 10) / 10;
    }
  });
}

async function collectPdfSurfaceStats(page: Page): Promise<PdfSurfaceStats> {
  return page.evaluate(() => {
    const rail = document.querySelector<HTMLElement>(
      '.work-pdf-thumbnail-rail',
    );
    const pageInput = document.querySelector<HTMLInputElement>(
      'input[aria-label="页码"]',
    );
    const root = document.querySelector('.work-pdf-native-viewer');
    const loadedBitmaps = root
      ? deepElements(root, 'img').filter(
          (element): element is HTMLImageElement =>
            element instanceof HTMLImageElement &&
            element.complete &&
            element.naturalWidth > 0,
        )
      : [];
    const mountedPages = root
      ? deepElements(root, 'div[style*="touch-action: none"]')
      : [];
    return {
      currentPage: numeric(pageInput?.value),
      loadedBitmapCount: loadedBitmaps.length,
      mountedPageCount: mountedPages.length,
      mountedThumbnailCount: document.querySelectorAll(
        '[data-pdf-page-thumbnail]',
      ).length,
      readyThumbnailCount: document.querySelectorAll(
        '[data-pdf-thumbnail-state="ready"]',
      ).length,
      thumbnailWindowEnd: numeric(rail?.dataset.pdfThumbnailWindowEnd),
      thumbnailWindowStart: numeric(rail?.dataset.pdfThumbnailWindowStart),
      totalPages: numeric(rail?.dataset.pdfPageCount),
    };

    function deepElements(
      rootElement: ParentNode,
      selector: string,
    ): Element[] {
      const matches = Array.from(rootElement.querySelectorAll(selector));
      for (const element of rootElement.querySelectorAll('*')) {
        if (element.shadowRoot) {
          matches.push(...deepElements(element.shadowRoot, selector));
        }
      }
      return [...new Set(matches)];
    }

    function numeric(value: string | undefined): number | null {
      const parsed = Number(value);
      return Number.isFinite(parsed) ? parsed : null;
    }
  });
}

function fixturePath(): string {
  return fileURLToPath(new URL(fixture, import.meta.url));
}

function argument(name: string): string | undefined {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function positiveInteger(value: string | undefined, fallback: number): number {
  const numeric = Number(value);
  return Number.isInteger(numeric) && numeric > 0 ? numeric : fallback;
}

async function saveReport(): Promise<void> {
  await writeFile(resultPath, `${JSON.stringify(report, null, 2)}\n`);
}
