import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium, type Page } from '@playwright/test';

interface EditLongTask {
  duration: number;
  startTime: number;
}

declare global {
  interface Window {
    __a3sSpreadsheetEditLongTasks?: EditLongTask[];
  }
}

const targetUrl = process.argv[2] ?? 'http://127.0.0.1:4175/';
const repeatCount = positiveInteger(process.argv[3], 3);
const timeoutMs = 180_000;
const fixturePath = fileURLToPath(
  new URL('./fixtures/spreadsheet-table-100k-x10.xlsx', import.meta.url),
);
const results: Array<Record<string, unknown>> = [];

for (let run = 1; run <= repeatCount; run += 1) {
  const browser = await chromium.launch({
    args: [
      '--disable-background-timer-throttling',
      '--disable-renderer-backgrounding',
      '--enable-precise-memory-info',
    ],
    headless: true,
  });
  try {
    const context = await browser.newContext({
      deviceScaleFactor: 1,
      viewport: { height: 1_000, width: 1_440 },
    });
    const page = await context.newPage();
    const browserErrors: string[] = [];
    const consoleErrors: string[] = [];
    page.on('console', (message) => {
      if (message.type() === 'error') consoleErrors.push(message.text());
    });
    page.on('pageerror', (error) => browserErrors.push(error.message));
    await installLongTaskObserver(page);
    await page.goto(targetUrl, {
      timeout: 120_000,
      waitUntil: 'domcontentloaded',
    });
    const input = page.locator('input[aria-label="打开 Office 或 PDF 文件"]');
    await input.waitFor({ state: 'attached', timeout: 60_000 });
    await input.setInputFiles(fixturePath, { timeout: timeoutMs });
    await page
      .locator('.work-spreadsheet-canvas > .fortune-container')
      .waitFor({ state: 'visible', timeout: timeoutMs });
    await page.waitForFunction(
      () =>
        (document.querySelector<HTMLElement>('.luckysheet-scrollbar-y')
          ?.scrollHeight ?? 0) >= 2_000_000,
      undefined,
      { timeout: timeoutMs },
    );
    await waitForFrames(page, 2);

    const edits = [];
    edits.push(await editCell(page, 'J100000', `A3S_EDIT_${run}_A`));
    edits.push(await editCell(page, 'I100000', `A3S_EDIT_${run}_B`));
    const persistedValues = {
      I100000: await readCellFromFormulaBar(page, 'I100000'),
      J100000: await readCellFromFormulaBar(page, 'J100000'),
    };
    results.push({
      browserErrors,
      consoleErrors,
      edits,
      persistedValues,
      run,
    });
    console.log(JSON.stringify({ event: 'run-finish', persistedValues, run }));
  } finally {
    await browser.close();
  }
}

const outputDirectory = fileURLToPath(new URL('./results/', import.meta.url));
await mkdir(outputDirectory, { recursive: true });
const timestamp = new Date().toISOString().replaceAll(':', '-');
const outputPath = resolve(
  outputDirectory,
  `${timestamp}-spreadsheet-edit.json`,
);
await writeFile(
  outputPath,
  `${JSON.stringify({ repeatCount, results, targetUrl }, null, 2)}\n`,
);
console.log(JSON.stringify({ event: 'complete', outputPath }));

async function editCell(
  page: Page,
  reference: string,
  value: string,
): Promise<Record<string, unknown>> {
  await selectCell(page, reference);
  await page.evaluate(() => {
    performance.clearMeasures('a3s-office.spreadsheet.controlled-projection');
    performance.clearMeasures('a3s-office.spreadsheet.fortune-projection');
    window.__a3sSpreadsheetEditLongTasks = [];
  });
  const formulaBar = page.locator('[aria-label="当前单元格输入"]');
  await formulaBar.fill(value);
  const startedAt = await page.evaluate(() => performance.now());
  await page.keyboard.press('Enter');
  await page.waitForFunction(
    (started) => {
      const controlled = performance.getEntriesByName(
        'a3s-office.spreadsheet.controlled-projection',
      );
      const projected = performance.getEntriesByName(
        'a3s-office.spreadsheet.fortune-projection',
      );
      return (
        controlled.some((entry) => entry.startTime >= started) &&
        projected.some((entry) => entry.startTime >= started)
      );
    },
    startedAt,
    { timeout: 30_000 },
  );
  await waitForFrames(page, 2);
  return page.evaluate(
    ({ editStartedAt, expectedReference }) => {
      const controlled = performance
        .getEntriesByName('a3s-office.spreadsheet.controlled-projection')
        .filter((entry) => entry.startTime >= editStartedAt)
        .at(-1) as PerformanceMeasure | undefined;
      const projected = performance
        .getEntriesByName('a3s-office.spreadsheet.fortune-projection')
        .filter((entry) => entry.startTime >= editStartedAt)
        .at(-1) as PerformanceMeasure | undefined;
      const longTasks = (window.__a3sSpreadsheetEditLongTasks ?? []).filter(
        (entry) => entry.startTime >= editStartedAt,
      );
      const nameBox =
        document.querySelector<HTMLInputElement>('.fortune-name-box');
      return {
        commitToStableFramesMs:
          Math.round((performance.now() - editStartedAt) * 10) / 10,
        controlledProjection: controlled
          ? {
              detail: controlled.detail,
              durationMs: Math.round(controlled.duration * 10) / 10,
            }
          : null,
        expectedReference,
        fortuneProjectionMs: Math.round((projected?.duration ?? 0) * 10) / 10,
        location: nameBox?.value ?? nameBox?.textContent?.trim() ?? null,
        longTaskCount: longTasks.length,
        longestTaskMs: Math.round(
          Math.max(0, ...longTasks.map((entry) => entry.duration)),
        ),
      };
    },
    { editStartedAt: startedAt, expectedReference: reference },
  );
}

async function readCellFromFormulaBar(
  page: Page,
  reference: string,
): Promise<string> {
  await selectCell(page, reference);
  const formulaBar = page.locator('[aria-label="当前单元格输入"]');
  return (await formulaBar.textContent())?.trim() ?? '';
}

async function selectCell(page: Page, reference: string): Promise<void> {
  const match = /^([A-Z]+)(\d+)$/.exec(reference);
  if (!match || Number(match[2]) !== 100_000) {
    throw new Error(`Unsupported benchmark cell reference: ${reference}`);
  }
  const targetColumn = columnIndex(match[1]);
  const overlay = page.locator('.fortune-sheet-overlay');
  await overlay.focus();
  await page.keyboard.press('Control+End');
  for (let column = 9; column > targetColumn; column -= 1) {
    await page.keyboard.press('ArrowLeft');
  }
  await page.waitForFunction((expected) => {
    const element =
      document.querySelector<HTMLInputElement>('.fortune-name-box');
    return (element?.value ?? element?.textContent?.trim()) === expected;
  }, reference);
  await waitForFrames(page, 1);
}

function columnIndex(label: string): number {
  let value = 0;
  for (const character of label) {
    value = value * 26 + character.charCodeAt(0) - 64;
  }
  return value - 1;
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

async function installLongTaskObserver(page: Page): Promise<void> {
  await page.addInitScript(() => {
    window.__a3sSpreadsheetEditLongTasks = [];
    new PerformanceObserver((list) => {
      for (const entry of list.getEntries()) {
        window.__a3sSpreadsheetEditLongTasks?.push({
          duration: entry.duration,
          startTime: entry.startTime,
        });
      }
    }).observe({ entryTypes: ['longtask'] });
  });
}

function positiveInteger(value: string | undefined, fallback: number): number {
  const parsed = Number.parseInt(value ?? '', 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}
