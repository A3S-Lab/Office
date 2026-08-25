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
  round,
  subtractMetrics,
} from './benchmark-metrics';

type BenchmarkMode = 'default' | 'content-visibility';

interface PresentationMilestones {
  editorReadyMs: number;
  sceneReadyMs: number;
  shellMountedMs: number;
  thumbnailSceneReadyMs: number;
}

interface PresentationSurfaceStats {
  activeSceneElementCount: number;
  mountedThumbnailCount: number;
  renderedThumbnailCount: number;
  selectedSlideIndex: number | null;
  slideCount: number | null;
  thumbnailSceneElementCount: number;
  thumbnailWindowEnd: number | null;
  thumbnailWindowStart: number | null;
  thumbnailWindowed: boolean;
}

interface PresentationNudgeResult {
  commitMs: number;
  durationMs: number;
  fromLeft: string;
  longTasks: LongTaskSummary;
  metrics: MetricSnapshot;
  toLeft: string;
}

interface PresentationBenchmarkRun {
  browserErrors: string[];
  consoleErrors: string[];
  failure?: string;
  importLongTasks?: LongTaskSummary;
  jumpLongTasks?: LongTaskSummary;
  jumpMs?: number;
  metrics?: {
    afterImport: MetricSnapshot;
    afterJump: MetricSnapshot;
    beforeImport: MetricSnapshot;
    retained: MetricSnapshot;
  };
  milestones?: PresentationMilestones;
  mode: BenchmarkMode;
  nudges?: PresentationNudgeResult[];
  run: number;
  submittedAt: string;
  surfaceAfterImport?: PresentationSurfaceStats;
  surfaceAfterJump?: PresentationSurfaceStats;
}

interface PresentationBrowserBenchmarkState {
  editorReadyMs: number | null;
  jumpMs: number | null;
  sceneReadyMs: number | null;
  shellMountedMs: number | null;
  startedAt: number | null;
  thumbnailSceneReadyMs: number | null;
}

interface PresentationNudgeBenchmarkState {
  commitMs: number | null;
  startedAt: number | null;
}

declare global {
  interface Window {
    __a3sPresentationBenchmark?: PresentationBrowserBenchmarkState;
    __a3sPresentationNudgeBenchmark?: PresentationNudgeBenchmarkState;
  }
}

const targetUrl = argument('--url') ?? 'http://127.0.0.1:4175/playground/';
const repeatCount = positiveInteger(argument('--runs'), 3);
const timeoutMs = positiveInteger(argument('--timeout-ms'), 120_000);
const fixture =
  argument('--fixture') ?? './fixtures/presentation-slides-1000.pptx';
const requestedModes = (argument('--modes') ?? 'default')
  .split(',')
  .map((value) => value.trim())
  .filter((value): value is BenchmarkMode =>
    ['content-visibility', 'default'].includes(value),
  );

if (!requestedModes.length) {
  throw new Error('No valid presentation benchmark modes selected.');
}

const timestamp = new Date().toISOString().replaceAll(':', '-');
const resultDirectory = fileURLToPath(new URL('./results/', import.meta.url));
const resultPath = resolve(resultDirectory, `${timestamp}-presentation.json`);
const report: {
  completedAt?: string;
  environment: Record<string, unknown>;
  runs: PresentationBenchmarkRun[];
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

for (let run = 1; run <= repeatCount; run += 1) {
  const modes = run % 2 === 1 ? requestedModes : [...requestedModes].reverse();
  for (const mode of modes) {
    console.log(JSON.stringify({ event: 'run-start', mode, run }));
    const result = await benchmarkPresentation(mode, run);
    report.runs.push(result);
    await saveReport();
    console.log(
      JSON.stringify({
        event: 'run-finish',
        failure: result.failure ?? null,
        jumpMs: result.jumpMs ?? null,
        milestones: result.milestones ?? null,
        mode,
        nudges: result.nudges ?? null,
        run,
      }),
    );
  }
}

report.completedAt = new Date().toISOString();
await saveReport();
console.log(JSON.stringify({ event: 'complete', resultPath }));

async function benchmarkPresentation(
  mode: BenchmarkMode,
  run: number,
): Promise<PresentationBenchmarkRun> {
  let browser: Browser | undefined;
  let cdp: CDPSession | undefined;
  const result: PresentationBenchmarkRun = {
    browserErrors: [],
    consoleErrors: [],
    mode,
    run,
    submittedAt: new Date().toISOString(),
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
    if (mode === 'content-visibility') {
      await page.addStyleTag({
        content: [
          '.work-slide-thumbnail-list > button,',
          '.work-presentation-sorter-grid > button {',
          '  content-visibility: auto;',
          '  contain-intrinsic-block-size: auto 88px;',
          '}',
        ].join('\n'),
      });
    }

    await installPresentationMilestoneClock(page);
    const beforeImport = await collectMetrics(cdp);
    await resetBrowserMeasurements(page);
    await fileInput.setInputFiles(fixturePath(), { timeout: timeoutMs });
    await page.waitForFunction(
      () => window.__a3sPresentationBenchmark?.editorReadyMs !== null,
      undefined,
      { polling: 'raf', timeout: timeoutMs },
    );
    await waitForVisualQuiet(page);
    result.milestones = await readPresentationMilestones(page);
    result.importLongTasks = await browserLongTasks(page);
    result.surfaceAfterImport = await collectPresentationSurfaceStats(page);
    const afterImport = await collectMetrics(cdp);

    await resetBrowserMeasurements(page);
    await installPresentationEndJumpClock(page);
    await page.locator('[data-slide-thumbnail][data-slide-index="0"]').focus();
    await page.keyboard.press('End');
    await page.waitForFunction(
      () => window.__a3sPresentationBenchmark?.jumpMs !== null,
      undefined,
      { polling: 'raf', timeout: timeoutMs },
    );
    await waitForVisualQuiet(page);
    result.jumpMs = await page.evaluate(
      () => window.__a3sPresentationBenchmark?.jumpMs ?? Number.NaN,
    );
    result.jumpLongTasks = await browserLongTasks(page);
    result.surfaceAfterJump = await collectPresentationSurfaceStats(page);
    const afterJump = await collectMetrics(cdp);

    const object = page
      .locator('.work-slide-canvas.interactive [data-slide-element-id]')
      .first();
    await object.focus();
    await page.waitForFunction(
      () =>
        document.activeElement?.getAttribute('data-slide-element-selected') ===
        'true',
      undefined,
      { polling: 'raf', timeout: 30_000 },
    );
    result.nudges = [
      await nudgeSelectedObject(page, cdp),
      await nudgeSelectedObject(page, cdp),
    ];

    await cdp.send('HeapProfiler.collectGarbage');
    result.metrics = {
      afterImport,
      afterJump,
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

async function installPresentationMilestoneClock(page: Page): Promise<void> {
  await page.evaluate(() => {
    const state: PresentationBrowserBenchmarkState = {
      editorReadyMs: null,
      jumpMs: null,
      sceneReadyMs: null,
      shellMountedMs: null,
      startedAt: null,
      thumbnailSceneReadyMs: null,
    };
    window.__a3sPresentationBenchmark = state;
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
            document.querySelector('.work-presentation-editor')
          ) {
            state.shellMountedMs = elapsed;
          }
          const activeScene = document.querySelector(
            '.work-slide-canvas.interactive',
          );
          if (
            state.sceneReadyMs === null &&
            activeScene?.querySelectorAll('[data-slide-element-id]').length ===
              9 &&
            activeScene.textContent?.includes('presentation slide 1')
          ) {
            state.sceneReadyMs = elapsed;
          }
          if (
            state.thumbnailSceneReadyMs === null &&
            document.querySelectorAll(
              '[data-slide-thumbnail-rendered="true"] .work-slide-element',
            ).length >= 9
          ) {
            state.thumbnailSceneReadyMs = elapsed;
          }
          const rail = document.querySelector<HTMLElement>(
            '.work-slide-strip, .work-presentation-sorter',
          );
          if (
            state.editorReadyMs === null &&
            state.shellMountedMs !== null &&
            state.sceneReadyMs !== null &&
            state.thumbnailSceneReadyMs !== null &&
            Number(rail?.dataset.slideCount) === 1_000
          ) {
            state.editorReadyMs = elapsed;
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

async function installPresentationEndJumpClock(page: Page): Promise<void> {
  await page.evaluate(() => {
    const state = window.__a3sPresentationBenchmark;
    const first = document.querySelector<HTMLElement>(
      '[data-slide-thumbnail][data-slide-index="0"]',
    );
    if (!state || !first) {
      throw new Error('Presentation jump instrumentation is unavailable.');
    }
    state.jumpMs = null;
    first.addEventListener(
      'keydown',
      (event) => {
        if (event.key !== 'End') return;
        const startedAt = performance.now();
        const inspect = () => {
          const selected = document.querySelector<HTMLElement>(
            '[data-slide-thumbnail].active',
          );
          const activeScene = document.querySelector(
            '.work-slide-canvas.interactive',
          );
          if (
            selected?.dataset.slideIndex === '999' &&
            activeScene?.textContent?.includes('presentation slide 1000') &&
            activeScene.querySelectorAll('[data-slide-element-id]').length === 9
          ) {
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

async function nudgeSelectedObject(
  page: Page,
  cdp: CDPSession,
): Promise<PresentationNudgeResult> {
  const object = page.locator(
    '.work-slide-canvas.interactive [data-slide-element-selected="true"]',
  );
  const fromLeft = (await object.getAttribute('style')) ?? '';
  await installPresentationNudgeClock(page, fromLeft);
  await resetBrowserMeasurements(page);
  const beforeNudge = await collectMetrics(cdp);
  await page.keyboard.press('ArrowRight');
  await page.waitForFunction(
    () => window.__a3sPresentationNudgeBenchmark?.commitMs !== null,
    undefined,
    { polling: 'raf', timeout: 30_000 },
  );
  await waitForFrames(page, 2);
  const timing = await page.evaluate(() => {
    const state = window.__a3sPresentationNudgeBenchmark;
    if (state?.commitMs === null || state?.startedAt === null || !state) {
      throw new Error('Presentation nudge timing is incomplete.');
    }
    return {
      commitMs: state.commitMs,
      durationMs: performance.now() - state.startedAt,
    };
  });
  const afterNudge = await collectMetrics(cdp);
  return {
    commitMs: round(timing.commitMs),
    durationMs: round(timing.durationMs),
    fromLeft,
    longTasks: await browserLongTasks(page),
    metrics: subtractMetrics(afterNudge, beforeNudge),
    toLeft: (await object.getAttribute('style')) ?? '',
  };
}

async function installPresentationNudgeClock(
  page: Page,
  previousStyle: string,
): Promise<void> {
  await page.evaluate((styleBeforeNudge) => {
    const selected = document.querySelector<HTMLElement>(
      '.work-slide-canvas.interactive [data-slide-element-selected="true"]',
    );
    if (!selected) throw new Error('The selected presentation object is gone.');
    const state: PresentationNudgeBenchmarkState = {
      commitMs: null,
      startedAt: null,
    };
    window.__a3sPresentationNudgeBenchmark = state;
    const observer = new MutationObserver(() => {
      if (
        state.startedAt === null ||
        selected.getAttribute('style') === styleBeforeNudge
      ) {
        return;
      }
      state.commitMs = performance.now() - state.startedAt;
      observer.disconnect();
    });
    selected.addEventListener(
      'keydown',
      (event) => {
        if (event.key !== 'ArrowRight') return;
        state.startedAt = performance.now();
        observer.observe(selected, {
          attributeFilter: ['style'],
          attributes: true,
        });
      },
      { capture: true, once: true },
    );
  }, previousStyle);
}

async function readPresentationMilestones(
  page: Page,
): Promise<PresentationMilestones> {
  return page.evaluate(() => {
    const state = window.__a3sPresentationBenchmark;
    if (
      !state ||
      state.editorReadyMs === null ||
      state.sceneReadyMs === null ||
      state.shellMountedMs === null ||
      state.thumbnailSceneReadyMs === null
    ) {
      throw new Error('Presentation milestones are incomplete.');
    }
    return {
      editorReadyMs: round(state.editorReadyMs),
      sceneReadyMs: round(state.sceneReadyMs),
      shellMountedMs: round(state.shellMountedMs),
      thumbnailSceneReadyMs: round(state.thumbnailSceneReadyMs),
    };

    function round(value: number): number {
      return Math.round(value * 10) / 10;
    }
  });
}

async function collectPresentationSurfaceStats(
  page: Page,
): Promise<PresentationSurfaceStats> {
  return page.evaluate(() => {
    const rail = document.querySelector<HTMLElement>(
      '.work-slide-strip, .work-presentation-sorter',
    );
    const selected = rail?.querySelector<HTMLElement>(
      '[data-slide-thumbnail].active',
    );
    return {
      activeSceneElementCount: document.querySelectorAll(
        '.work-slide-canvas.interactive [data-slide-element-id]',
      ).length,
      mountedThumbnailCount:
        rail?.querySelectorAll('[data-slide-thumbnail]').length ?? 0,
      renderedThumbnailCount:
        rail?.querySelectorAll('[data-slide-thumbnail-rendered="true"]')
          .length ?? 0,
      selectedSlideIndex: numeric(selected?.dataset.slideIndex),
      slideCount: numeric(rail?.dataset.slideCount),
      thumbnailSceneElementCount:
        rail?.querySelectorAll(
          '[data-slide-thumbnail-rendered="true"] .work-slide-element',
        ).length ?? 0,
      thumbnailWindowEnd: numeric(rail?.dataset.slideWindowEnd),
      thumbnailWindowStart: numeric(rail?.dataset.slideWindowStart),
      thumbnailWindowed: rail?.dataset.slideWindowed === 'true',
    };

    function numeric(value: string | undefined): number | null {
      const parsed = Number(value);
      return Number.isFinite(parsed) ? parsed : null;
    }
  });
}

async function waitForFrames(page: Page, count: number): Promise<void> {
  await page.evaluate(
    (frames) =>
      new Promise<void>((resolvePromise) => {
        const next = (remaining: number) => {
          if (remaining <= 0) {
            resolvePromise();
            return;
          }
          requestAnimationFrame(() => next(remaining - 1));
        };
        next(frames);
      }),
    count,
  );
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
