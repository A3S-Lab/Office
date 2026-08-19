import { mkdir, writeFile } from 'node:fs/promises';
import { basename, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  type Browser,
  type CDPSession,
  chromium,
  type Page,
} from '@playwright/test';
import {
  browserLongTaskRecords,
  browserLongTasks,
  collectFailureState,
  collectModelStats,
  collectPaginationStats,
  elapsedMs,
  installPerformanceObserver,
  type JumpProfile,
  measureContinuousScroll,
  measureJump,
  resetBrowserMeasurements,
  type ScenarioKind,
  type ScrollProfile,
  validateEndPosition,
  waitForStableDocumentPagination,
  waitForVisibleSelector,
  waitForVisualQuiet,
} from './benchmark-browser';
import {
  collectMetrics,
  type LongTaskSummary,
  type MetricSnapshot,
  messageOf,
  subtractMetrics,
  within,
} from './benchmark-metrics';

interface Scenario {
  fixture: string;
  kind: ScenarioKind;
  readySelector: string;
}

interface BenchmarkRun {
  browserErrors: string[];
  consoleErrors: string[];
  failure?: string;
  failureMetrics?: MetricSnapshot;
  failureState?: Record<string, unknown>;
  fixture: string;
  kind: ScenarioKind;
  milestones?: {
    contentMountedMs: number;
    paginationReadyMs: number | null;
    paginationTerminalMs: number | null;
    shellMountedMs: number;
  };
  importLongTasks?: LongTaskSummary;
  importLongTasksAtEditorVisible?: LongTaskSummary;
  importLongTaskRecords?: Array<{
    durationMs: number;
    phases: string[];
    startTimeMs: number;
  }>;
  metrics?: {
    afterImport: MetricSnapshot;
    afterScroll: MetricSnapshot;
    beforeImport: MetricSnapshot;
    scrollDelta: MetricSnapshot;
  };
  model?: Record<string, unknown>;
  pagination?: Record<string, string>;
  retainedMetrics?: MetricSnapshot;
  run: number;
  scroll?: {
    endJump: JumpProfile;
    longTasks: LongTaskSummary;
    middleJump: JumpProfile;
    profile: ScrollProfile;
  };
  submittedAt: string;
  validation?: Record<string, unknown>;
  workers?: string[];
}
const scenarios: Record<ScenarioKind, Scenario> = {
  'document-text': {
    fixture: 'fixtures/document-text-100k.docx',
    kind: 'document-text',
    readySelector:
      ".work-document-editable .ProseMirror[data-document-windowed='true'][data-document-chunk-count='782']",
  },
  'document-table': {
    fixture: 'fixtures/document-table-100k.docx',
    kind: 'document-table',
    readySelector:
      ".work-document-editable .ProseMirror[data-document-windowed='true'][data-document-chunk-count='6250']",
  },
  spreadsheet: {
    fixture: 'fixtures/spreadsheet-table-100k-x10.xlsx',
    kind: 'spreadsheet',
    readySelector: '.work-spreadsheet-canvas > .fortune-container',
  },
};

const targetUrl = argument('--url') ?? 'https://a3s-lab.github.io/Office/';
const repeatCount = positiveInteger(argument('--runs'), 3);
const timeoutMs = positiveInteger(argument('--timeout-ms'), 180_000);
const requestedKinds = (
  argument('--scenarios') ?? Object.keys(scenarios).join(',')
)
  .split(',')
  .map((value) => value.trim())
  .filter((value): value is ScenarioKind => value in scenarios);

if (!requestedKinds.length) {
  throw new Error('No valid scenarios were selected.');
}

const timestamp = new Date().toISOString().replaceAll(':', '-');
const resultDirectory = fileURLToPath(new URL('./results/', import.meta.url));
const resultPath = resolve(resultDirectory, `${timestamp}.json`);
const report: {
  completedAt?: string;
  environment: Record<string, unknown>;
  runs: BenchmarkRun[];
  startedAt: string;
} = {
  environment: {
    headless: true,
    playwright: '1.61.1',
    repeatCount,
    scenarios: requestedKinds,
    targetUrl,
    timeoutMs,
  },
  runs: [],
  startedAt: new Date().toISOString(),
};

await mkdir(resultDirectory, { recursive: true });

for (const kind of requestedKinds) {
  for (let run = 1; run <= repeatCount; run += 1) {
    const scenario = scenarios[kind];
    console.log(
      JSON.stringify({
        event: 'run-start',
        fixture: scenario.fixture,
        kind,
        run,
      }),
    );
    const result = await benchmarkScenario(scenario, run);
    report.runs.push(result);
    await saveReport();
    console.log(
      JSON.stringify({
        event: 'run-finish',
        failure: result.failure ?? null,
        kind,
        milestones: result.milestones ?? null,
        run,
      }),
    );
  }
}

report.completedAt = new Date().toISOString();
await saveReport();
console.log(JSON.stringify({ event: 'complete', resultPath }));

async function benchmarkScenario(
  scenario: Scenario,
  run: number,
): Promise<BenchmarkRun> {
  let browser: Browser | undefined;
  let diagnosticCdp: CDPSession | undefined;
  let diagnosticPage: Page | undefined;
  const browserErrors: string[] = [];
  const consoleErrors: string[] = [];
  const workers: string[] = [];
  const result: BenchmarkRun = {
    browserErrors,
    consoleErrors,
    fixture: basename(scenario.fixture),
    kind: scenario.kind,
    run,
    submittedAt: new Date().toISOString(),
    workers,
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
      viewport: { height: 1000, width: 1440 },
    });
    const page = await context.newPage();
    diagnosticPage = page;
    page.on('console', (message) => {
      if (message.type() === 'error') consoleErrors.push(message.text());
    });
    page.on('crash', () => browserErrors.push('Renderer crashed.'));
    page.on('pageerror', (error) => browserErrors.push(error.message));
    page.on('worker', (worker) => workers.push(worker.url()));
    await installPerformanceObserver(page);
    const cdp = await context.newCDPSession(page);
    diagnosticCdp = cdp;
    await cdp.send('Performance.enable');
    await page.goto(targetUrl, {
      timeout: 120_000,
      waitUntil: 'domcontentloaded',
    });
    await page
      .locator('input[aria-label="打开 Office 或 PDF 文件"]')
      .waitFor({ state: 'attached', timeout: 60_000 });
    report.environment.browser = await page.evaluate(() => ({
      deviceMemory: (navigator as Navigator & { deviceMemory?: number })
        .deviceMemory,
      hardwareConcurrency: navigator.hardwareConcurrency,
      userAgent: navigator.userAgent,
    }));

    const beforeImport = await collectMetrics(cdp);
    await resetBrowserMeasurements(page);
    const fixturePath = fileURLToPath(
      new URL(scenario.fixture, new URL('./', import.meta.url)),
    );
    await page
      .locator('input[aria-label="打开 Office 或 PDF 文件"]')
      .setInputFiles(fixturePath, { timeout: timeoutMs });
    await waitForVisibleSelector(page, scenario.readySelector, timeoutMs);
    const shellMountedMs = await elapsedMs(page);
    if (scenario.kind === 'spreadsheet') {
      await page.waitForFunction(
        () => {
          const scrollbar = document.querySelector<HTMLElement>(
            '.luckysheet-scrollbar-y',
          );
          return (scrollbar?.scrollHeight ?? 0) >= 2_000_000;
        },
        undefined,
        { timeout: timeoutMs },
      );
    }
    const contentMountedMs = await elapsedMs(page);
    result.importLongTasksAtEditorVisible = await browserLongTasks(page);

    let paginationReadyMs: number | null = null;
    let paginationTerminalMs: number | null = null;
    if (scenario.kind.startsWith('document')) {
      try {
        const state = await waitForStableDocumentPagination(page, timeoutMs);
        paginationTerminalMs = await elapsedMs(page);
        if (state === 'ready') paginationReadyMs = paginationTerminalMs;
        else browserErrors.push('Pagination entered the error state.');
      } catch (error) {
        browserErrors.push(
          `Pagination did not reach a terminal state: ${messageOf(error)}`,
        );
      }
    } else {
      await waitForVisualQuiet(page);
    }

    result.milestones = {
      contentMountedMs,
      paginationReadyMs,
      paginationTerminalMs,
      shellMountedMs,
    };
    result.importLongTasks = await browserLongTasks(page);
    result.importLongTaskRecords = await browserLongTaskRecords(page);
    result.model = await collectModelStats(page, scenario.kind);
    result.pagination = await collectPaginationStats(page);
    const afterImport = await collectMetrics(cdp);

    // Memory.getDOMCounters can enqueue V8 cleanup. Drain instrumentation work
    // before resetting the observer so the scroll profile measures the product.
    await waitForVisualQuiet(page);
    await resetBrowserMeasurements(page);
    const scrollSelector = scenario.kind.startsWith('document')
      ? '.work-document-scroll'
      : '.luckysheet-scrollbar-y';
    const middleJump = await measureJump(page, scrollSelector, 0.5);
    const endJump = await measureJump(page, scrollSelector, 1);
    const profile = await measureContinuousScroll(page, scrollSelector, 120);
    const scrollLongTasks = await browserLongTasks(page);
    result.validation = await validateEndPosition(page, scenario.kind);
    const afterScroll = await collectMetrics(cdp);
    result.metrics = {
      afterImport,
      afterScroll,
      beforeImport,
      scrollDelta: subtractMetrics(afterScroll, afterImport),
    };
    result.scroll = {
      endJump,
      longTasks: scrollLongTasks,
      middleJump,
      profile,
    };

    try {
      await cdp.send('HeapProfiler.collectGarbage');
      result.retainedMetrics = await collectMetrics(cdp);
    } catch (error) {
      browserErrors.push(
        `Retained heap measurement failed: ${messageOf(error)}`,
      );
    }
  } catch (error) {
    result.failure = messageOf(error);
    if (diagnosticPage && !diagnosticPage.isClosed()) {
      result.importLongTasks ??= await within(
        browserLongTasks(diagnosticPage),
        10_000,
      ).catch(() => undefined);
      result.failureState = await within(
        collectFailureState(diagnosticPage),
        10_000,
      ).catch((diagnosticError) => ({
        diagnosticError: messageOf(diagnosticError),
      }));
    }
    if (diagnosticCdp) {
      result.failureMetrics = await within(
        collectMetrics(diagnosticCdp),
        10_000,
      ).catch(() => undefined);
    }
  } finally {
    await browser?.close().catch(() => undefined);
  }
  return result;
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
