import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from '@playwright/test';

const targetUrl = process.argv[2] ?? 'http://127.0.0.1:4175/';
const runs = Number.parseInt(process.argv[3] ?? '3', 10);
const scenarios = [
  { chunkId: 'document-chunk-1-782', kind: 'text' },
  { chunkId: 'document-chunk-1-6250', kind: 'table' },
] as const;
const results: Array<Record<string, unknown>> = [];

for (const scenario of scenarios) {
  for (let run = 1; run <= runs; run += 1) {
    const browser = await chromium.launch({
      args: [
        '--disable-background-timer-throttling',
        '--disable-renderer-backgrounding',
      ],
      headless: true,
    });
    try {
      const context = await browser.newContext({
        deviceScaleFactor: 1,
        viewport: { height: 1_000, width: 1_440 },
      });
      const page = await context.newPage();
      const errors: string[] = [];
      page.on('console', (message) => {
        if (message.type() === 'error') errors.push(message.text());
      });
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(targetUrl, {
        timeout: 120_000,
        waitUntil: 'domcontentloaded',
      });
      const fixture = fileURLToPath(
        new URL(
          `./fixtures/document-${scenario.kind}-100k.docx`,
          import.meta.url,
        ),
      );
      await page
        .locator('input[aria-label="打开 Office 或 PDF 文件"]')
        .setInputFiles(fixture, { timeout: 180_000 });
      const editor = page.locator('.work-document-editable .ProseMirror');
      await editor.waitFor({ state: 'visible', timeout: 180_000 });
      await page.waitForFunction(
        () =>
          document.querySelector<HTMLElement>(
            '.work-document-editable .ProseMirror',
          )?.dataset.paginationState === 'ready',
        undefined,
        { timeout: 180_000 },
      );
      await editor.focus();
      await page.keyboard.press('Control+End');
      await page
        .locator(
          `[data-document-chunk-id="${scenario.chunkId}"][data-document-chunk-mounted="true"]:not([data-document-lazy-preview])`,
        )
        .waitFor({ state: 'visible', timeout: 30_000 });

      const edits = [];
      for (const marker of [' FIRST-EDIT', ' SECOND-EDIT']) {
        const before = await controlledPublishCount(page);
        await installPageEditTiming(page, before);
        await page.keyboard.insertText(marker);
        await page.waitForFunction(
          (count) =>
            Number(
              document.querySelector<HTMLElement>(
                '.work-document-editable .ProseMirror',
              )?.dataset.documentControlledPublishCount,
            ) > count,
          before,
          { timeout: 30_000 },
        );
        edits.push(
          await page.evaluate(() => {
            const element = document.querySelector<HTMLElement>(
              '.work-document-editable .ProseMirror',
            );
            const timing = (
              globalThis as typeof globalThis & {
                __a3sDocumentEditTiming?: {
                  publishedAt: number | null;
                  startedAt: number | null;
                };
              }
            ).__a3sDocumentEditTiming;
            const wallMs =
              timing?.startedAt !== null &&
              timing?.startedAt !== undefined &&
              timing.publishedAt !== null
                ? timing.publishedAt - timing.startedAt
                : Number.NaN;
            return {
              mode: element?.dataset.documentControlledPublishMode ?? null,
              publishMs: Number(
                element?.dataset.documentControlledPublishMs ?? Number.NaN,
              ),
              wallMs,
            };
          }),
        );
      }
      const finalText = await page
        .locator(`[data-document-chunk-id="${scenario.chunkId}"]`)
        .textContent();
      results.push({
        kind: scenario.kind,
        run,
        edits,
        errors,
        finalTextContainsBothEdits:
          finalText?.includes('FIRST-EDIT') === true &&
          finalText.includes('SECOND-EDIT'),
      });
    } finally {
      await browser.close();
    }
  }
}

const timestamp = new Date().toISOString().replaceAll(':', '-');
const resultPath = resolve(
  fileURLToPath(new URL('./results/', import.meta.url)),
  `${timestamp}-document-edit.json`,
);
await mkdir(fileURLToPath(new URL('./results/', import.meta.url)), {
  recursive: true,
});
await writeFile(
  resultPath,
  `${JSON.stringify({ completedAt: new Date().toISOString(), results }, null, 2)}\n`,
);
console.log(JSON.stringify({ resultPath, results }, null, 2));

async function controlledPublishCount(
  page: import('@playwright/test').Page,
): Promise<number> {
  return page.evaluate(
    () =>
      Number(
        document.querySelector<HTMLElement>(
          '.work-document-editable .ProseMirror',
        )?.dataset.documentControlledPublishCount,
      ) || 0,
  );
}

async function installPageEditTiming(
  page: import('@playwright/test').Page,
  publishCount: number,
): Promise<void> {
  await page.evaluate((count) => {
    const editor = document.querySelector<HTMLElement>(
      '.work-document-editable .ProseMirror',
    );
    if (!editor) throw new Error('The document editor is unavailable.');
    const timing = { publishedAt: null, startedAt: null } as {
      publishedAt: number | null;
      startedAt: number | null;
    };
    editor.addEventListener(
      'beforeinput',
      () => {
        timing.startedAt = performance.now();
      },
      { capture: true, once: true },
    );
    const observer = new MutationObserver(() => {
      if (Number(editor.dataset.documentControlledPublishCount) <= count) {
        return;
      }
      timing.publishedAt = performance.now();
      observer.disconnect();
    });
    observer.observe(editor, {
      attributeFilter: ['data-document-controlled-publish-count'],
      attributes: true,
    });
    Object.assign(globalThis, { __a3sDocumentEditTiming: timing });
  }, publishCount);
}
