import type { Page } from '@playwright/test';
import {
  type LongTaskSummary,
  percentile,
  round,
  summarizeDurations,
} from './benchmark-metrics';

export type ScenarioKind = 'document-table' | 'document-text' | 'spreadsheet';

export interface ScrollProfile {
  actualEnd: number;
  durationMs: number;
  effectiveFps: number;
  frameCount: number;
  frameIntervalAverageMs: number;
  frameIntervalMaximumMs: number;
  frameIntervalP95Ms: number;
  framesOver20Ms: number;
  framesOver33Ms: number;
  framesOver50Ms: number;
  maximumScrollTop: number;
}

export interface JumpProfile {
  actualScrollTop: number;
  firstFrameMs: number;
  maximumScrollTop: number;
  ratio: number;
  secondFrameMs: number;
}

interface BrowserBenchmarkState {
  lastLongTaskEnd: number;
  longTasks: Array<{ duration: number; startTime: number }>;
  startedAt: number;
}

declare global {
  interface Window {
    __a3sOfficeBenchmark?: BrowserBenchmarkState;
  }
}

export async function waitForStableDocumentPagination(
  page: Page,
  timeout: number,
): Promise<'error' | 'ready'> {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeout) {
    const remaining = Math.max(1, timeout - (Date.now() - startedAt));
    await page.waitForFunction(
      () => {
        const dataset = document.querySelector<HTMLElement>(
          '.work-document-editable .ProseMirror',
        )?.dataset;
        return (
          dataset?.paginationActive !== 'true' &&
          (dataset?.paginationState === 'ready' ||
            dataset?.paginationState === 'error')
        );
      },
      undefined,
      { timeout: remaining },
    );
    await waitForVisualQuiet(page);
    const state = await page.evaluate(() => {
      const dataset = document.querySelector<HTMLElement>(
        '.work-document-editable .ProseMirror',
      )?.dataset;
      return dataset?.paginationActive !== 'true' &&
        (dataset?.paginationState === 'ready' ||
          dataset?.paginationState === 'error')
        ? dataset.paginationState
        : null;
    });
    if (state === 'ready' || state === 'error') return state;
  }
  throw new Error(`Pagination did not stabilize within ${timeout}ms.`);
}

export async function waitForVisibleSelector(
  page: Page,
  selector: string,
  timeout: number,
): Promise<void> {
  await page.waitForFunction(
    (candidate) => {
      const element = document.querySelector(candidate);
      if (!(element instanceof Element)) return false;
      const style = getComputedStyle(element);
      return (
        style.display !== 'none' &&
        style.visibility !== 'hidden' &&
        style.visibility !== 'collapse' &&
        element.getClientRects().length > 0
      );
    },
    selector,
    { polling: 'raf', timeout },
  );
}

export async function installPerformanceObserver(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const state: BrowserBenchmarkState = {
      lastLongTaskEnd: 0,
      longTasks: [],
      startedAt: 0,
    };
    window.__a3sOfficeBenchmark = state;
    try {
      const observer = new PerformanceObserver((list) => {
        for (const entry of list.getEntries()) {
          state.longTasks.push({
            duration: entry.duration,
            startTime: entry.startTime,
          });
          state.lastLongTaskEnd = entry.startTime + entry.duration;
        }
      });
      observer.observe({ buffered: true, type: 'longtask' });
    } catch {
      // Long Task API support is recorded as an empty collection.
    }
  });
}

export async function resetBrowserMeasurements(page: Page): Promise<void> {
  await page.evaluate(() => {
    const state = window.__a3sOfficeBenchmark;
    if (!state) throw new Error('Benchmark observer was not installed.');
    performance.clearMeasures();
    state.lastLongTaskEnd = 0;
    state.longTasks.length = 0;
    state.startedAt = performance.now();
  });
}

export async function elapsedMs(page: Page): Promise<number> {
  return round(
    await page.evaluate(() => {
      const startedAt = window.__a3sOfficeBenchmark?.startedAt;
      if (startedAt === undefined)
        throw new Error('Benchmark was not started.');
      return performance.now() - startedAt;
    }),
  );
}

export async function browserLongTasks(page: Page): Promise<LongTaskSummary> {
  const durations = await page.evaluate(
    () =>
      window.__a3sOfficeBenchmark?.longTasks.map(({ duration }) => duration) ??
      [],
  );
  return summarizeDurations(durations);
}

export async function browserLongTaskRecords(
  page: Page,
): Promise<
  Array<{ durationMs: number; phases: string[]; startTimeMs: number }>
> {
  return page.evaluate(() => {
    const state = window.__a3sOfficeBenchmark;
    if (!state) return [];
    const measures = performance
      .getEntriesByType('measure')
      .filter(
        (entry) =>
          entry.name.startsWith('a3s-office.spreadsheet.') &&
          !entry.name.endsWith('.import-total') &&
          !entry.name.includes('worker') &&
          !(
            entry.name.endsWith('.style-gate') &&
            (entry as PerformanceMeasure).detail?.worker === true
          ),
      );
    return state.longTasks.map(({ duration, startTime }) => {
      const endTime = startTime + duration;
      return {
        durationMs: Math.round(duration * 10) / 10,
        phases: measures
          .filter(
            (measure) =>
              measure.startTime < endTime &&
              measure.startTime + measure.duration > startTime,
          )
          .sort((left, right) => left.duration - right.duration)
          .map((measure) =>
            measure.name.replace('a3s-office.spreadsheet.', ''),
          ),
        startTimeMs: Math.round((startTime - state.startedAt) * 10) / 10,
      };
    });
  });
}

export async function waitForVisualQuiet(page: Page): Promise<void> {
  await page.evaluate(
    () =>
      new Promise<void>((resolvePromise) => {
        const state = window.__a3sOfficeBenchmark;
        const startedAt = performance.now();
        const check = () => {
          const now = performance.now();
          const lastTask = state?.lastLongTaskEnd ?? 0;
          if (
            now - startedAt >= 300 &&
            (lastTask === 0 || now - lastTask >= 300)
          ) {
            resolvePromise();
            return;
          }
          if (now - startedAt >= 5_000) {
            resolvePromise();
            return;
          }
          requestAnimationFrame(check);
        };
        requestAnimationFrame(check);
      }),
  );
}

export async function collectModelStats(
  page: Page,
  kind: ScenarioKind,
): Promise<Record<string, unknown>> {
  if (kind === 'spreadsheet') {
    return page.evaluate(() => {
      const benchmarkStartedAt = window.__a3sOfficeBenchmark?.startedAt ?? 0;
      const scrollbar = document.querySelector<HTMLElement>(
        '.luckysheet-scrollbar-y',
      );
      const canvas = document.querySelector<HTMLCanvasElement>(
        '.fortune-sheet-canvas',
      );
      return {
        canvasHeight: canvas?.height ?? null,
        canvasWidth: canvas?.width ?? null,
        importMeasures: performance
          .getEntriesByType('measure')
          .filter((entry) => entry.name.startsWith('a3s-office.spreadsheet.'))
          .map((entry) => ({
            detail: (entry as PerformanceMeasure).detail,
            durationMs: Math.round(entry.duration * 10) / 10,
            name: entry.name,
            startTimeMs:
              Math.round((entry.startTime - benchmarkStartedAt) * 10) / 10,
          })),
        scrollClientHeight: scrollbar?.clientHeight ?? null,
        scrollHeight: scrollbar?.scrollHeight ?? null,
      };
    });
  }
  return page.evaluate(() => {
    const benchmarkStartedAt = window.__a3sOfficeBenchmark?.startedAt ?? 0;
    const editor = document.querySelector<HTMLElement>(
      '.work-document-editable .ProseMirror',
    );
    const pageStack = document.querySelector<HTMLElement>(
      '.work-document-page-stack',
    );
    const firstTableBody =
      editor?.querySelector<HTMLTableSectionElement>('tbody');
    return {
      documentChunkCount: Number(editor?.dataset.documentChunkCount) || 0,
      documentEditorMountMs:
        Number(editor?.dataset.documentEditorMountMs) || null,
      documentExternalApplyMs:
        Number(editor?.dataset.documentExternalApplyMs) || null,
      documentMountedChunkCount:
        Number(editor?.dataset.documentMountedChunkCount) || 0,
      documentWindowed: editor?.dataset.documentWindowed === 'true',
      editorChildElements: editor?.childElementCount ?? null,
      importMeasures: performance
        .getEntriesByType('measure')
        .filter((entry) => entry.name.startsWith('a3s-office.document.'))
        .map((entry) => ({
          detail: (entry as PerformanceMeasure).detail,
          durationMs: Math.round(entry.duration * 10) / 10,
          name: entry.name,
          startTimeMs:
            Math.round((entry.startTime - benchmarkStartedAt) * 10) / 10,
        })),
      mountedPageSheets:
        pageStack?.querySelectorAll('[data-work-document-page-sheet]').length ??
        null,
      pageCount: Number(pageStack?.dataset.pageCount) || null,
      pageWindowEnd: Number(pageStack?.dataset.pageWindowEnd) || null,
      pageWindowStart: Number(pageStack?.dataset.pageWindowStart) || null,
      paragraphs: editor?.getElementsByTagName('p').length ?? null,
      tableRows: firstTableBody?.rows.length ?? 0,
      textLength: editor?.innerText.length ?? null,
    };
  });
}

export async function collectPaginationStats(
  page: Page,
): Promise<Record<string, string>> {
  return page.evaluate(() => {
    const dataset = document.querySelector<HTMLElement>(
      '.work-document-editable .ProseMirror',
    )?.dataset;
    if (!dataset) return {};
    return Object.fromEntries(
      Object.entries(dataset).filter(
        (entry): entry is [string, string] =>
          entry[0].startsWith('pagination') && entry[1] !== undefined,
      ),
    );
  });
}

export async function collectFailureState(
  page: Page,
): Promise<Record<string, unknown>> {
  return page.evaluate(() => {
    const progress = document.querySelector<HTMLProgressElement>(
      '.playground-import-progress progress',
    );
    const progressOutput = document.querySelector<HTMLOutputElement>(
      '.playground-import-progress output',
    );
    const editor = document.querySelector<HTMLElement>(
      '.work-document-editable .ProseMirror',
    );
    return {
      editorMounted: editor !== null,
      elapsedMs: Math.round(
        performance.now() - (window.__a3sOfficeBenchmark?.startedAt ?? 0),
      ),
      importProgress: progress?.value ?? null,
      importStage: progressOutput?.textContent?.trim() ?? null,
      paginationState: editor?.dataset.paginationState ?? null,
    };
  });
}

export async function measureJump(
  page: Page,
  selector: string,
  ratio: number,
): Promise<JumpProfile> {
  return page.evaluate(
    async ({ ratio: targetRatio, selector: targetSelector }) => {
      const scroll = document.querySelector<HTMLElement>(targetSelector);
      if (!scroll)
        throw new Error(`Scroll container not found: ${targetSelector}`);
      scroll.style.setProperty('scroll-behavior', 'auto', 'important');
      const maximumScrollTop = Math.max(
        0,
        scroll.scrollHeight - scroll.clientHeight,
      );
      const startedAt = performance.now();
      scroll.scrollTop = maximumScrollTop * targetRatio;
      const firstFrame = await new Promise<number>((resolveFrame) =>
        requestAnimationFrame(() => resolveFrame(performance.now())),
      );
      const secondFrame = await new Promise<number>((resolveFrame) =>
        requestAnimationFrame(() => resolveFrame(performance.now())),
      );
      return {
        actualScrollTop: Math.round(scroll.scrollTop),
        firstFrameMs: Math.round((firstFrame - startedAt) * 10) / 10,
        maximumScrollTop: Math.round(maximumScrollTop),
        ratio: targetRatio,
        secondFrameMs: Math.round((secondFrame - startedAt) * 10) / 10,
      };
    },
    { ratio, selector },
  );
}

export async function measureContinuousScroll(
  page: Page,
  selector: string,
  frameCount: number,
): Promise<ScrollProfile> {
  const raw = await page.evaluate(
    async ({ frames, targetSelector }) => {
      const scroll = document.querySelector<HTMLElement>(targetSelector);
      if (!scroll)
        throw new Error(`Scroll container not found: ${targetSelector}`);
      scroll.style.setProperty('scroll-behavior', 'auto', 'important');
      const maximumScrollTop = Math.max(
        0,
        scroll.scrollHeight - scroll.clientHeight,
      );
      scroll.scrollTop = 0;
      await new Promise<void>((resolveFrame) =>
        requestAnimationFrame(() => resolveFrame()),
      );
      let previous = performance.now();
      const intervals: number[] = [];
      const startedAt = previous;
      for (let index = 1; index <= frames; index += 1) {
        scroll.scrollTop = (maximumScrollTop * index) / frames;
        const current = await new Promise<number>((resolveFrame) =>
          requestAnimationFrame(() => resolveFrame(performance.now())),
        );
        intervals.push(current - previous);
        previous = current;
      }
      await new Promise<void>((resolveFrame) =>
        requestAnimationFrame(() => resolveFrame()),
      );
      return {
        actualEnd: scroll.scrollTop,
        durationMs: previous - startedAt,
        intervals,
        maximumScrollTop,
      };
    },
    { frames: frameCount, targetSelector: selector },
  );
  const sorted = [...raw.intervals].sort((left, right) => left - right);
  const average =
    raw.intervals.reduce((sum, value) => sum + value, 0) /
    Math.max(1, raw.intervals.length);
  return {
    actualEnd: Math.round(raw.actualEnd),
    durationMs: round(raw.durationMs),
    effectiveFps: round(1000 / average),
    frameCount: raw.intervals.length,
    frameIntervalAverageMs: round(average),
    frameIntervalMaximumMs: round(sorted.at(-1) ?? 0),
    frameIntervalP95Ms: round(percentile(sorted, 0.95)),
    framesOver20Ms: raw.intervals.filter((value) => value > 20).length,
    framesOver33Ms: raw.intervals.filter((value) => value > 33).length,
    framesOver50Ms: raw.intervals.filter((value) => value > 50).length,
    maximumScrollTop: Math.round(raw.maximumScrollTop),
  };
}

export async function validateEndPosition(
  page: Page,
  kind: ScenarioKind,
): Promise<Record<string, unknown>> {
  if (kind === 'spreadsheet') {
    await page.locator('.fortune-sheet-overlay').focus();
    const startedAt = await page.evaluate(() => performance.now());
    await page.keyboard.press('Control+End');
    await page.evaluate(
      () =>
        new Promise<void>((resolvePromise) =>
          requestAnimationFrame(() =>
            requestAnimationFrame(() => resolvePromise()),
          ),
        ),
    );
    return page.evaluate((started) => {
      const nameBox =
        document.querySelector<HTMLInputElement>('.fortune-name-box');
      const scrollbar = document.querySelector<HTMLElement>(
        '.luckysheet-scrollbar-y',
      );
      return {
        controlEndMs: Math.round((performance.now() - started) * 10) / 10,
        location: nameBox?.value ?? nameBox?.textContent?.trim() ?? null,
        scrollTop: scrollbar?.scrollTop ?? null,
      };
    }, startedAt);
  }
  return page.evaluate((scenarioKind) => {
    const editor = document.querySelector<HTMLElement>(
      '.work-document-editable .ProseMirror',
    );
    const scroll = document.querySelector<HTMLElement>('.work-document-scroll');
    const scrollRect = scroll?.getBoundingClientRect();
    const windowDiagnostics = {
      documentMountedChunkCount:
        Number(editor?.dataset.documentMountedChunkCount) || 0,
      documentWindowLastMs: Number(editor?.dataset.documentWindowLastMs) || 0,
      documentWindowMaxMs: Number(editor?.dataset.documentWindowMaxMs) || 0,
      documentWindowTotalMs: Number(editor?.dataset.documentWindowTotalMs) || 0,
      documentWindowUpdates: Number(editor?.dataset.documentWindowUpdates) || 0,
      paginationIndexedChunks:
        Number(editor?.dataset.paginationIndexedChunks) || 0,
      paginationVisibleBreaks:
        Number(editor?.dataset.paginationVisibleBreaks) || 0,
      paginationVisibleDecorations:
        Number(editor?.dataset.paginationVisibleDecorations) || 0,
      paginationVisualBreaks:
        Number(editor?.dataset.paginationVisualBreaks) || 0,
    };
    if (scenarioKind === 'document-table') {
      const mountedRows = editor?.querySelectorAll<HTMLTableRowElement>(
        'tbody > tr:not(.work-document-table-page-break)',
      );
      const lastRow = mountedRows?.item(Math.max(0, mountedRows.length - 1));
      const lastRowRect = lastRow?.getBoundingClientRect();
      const finalChunk = Array.from(
        editor?.querySelectorAll<HTMLElement>('[data-document-chunk-id]') ?? [],
      ).at(-1);
      const finalChunkRect = finalChunk?.getBoundingClientRect();
      const containers = Array.from(
        editor?.querySelectorAll<HTMLElement>(
          '[data-document-chunk-window-container="true"]',
        ) ?? [],
      );
      const finalContainer = finalChunk?.closest<HTMLElement>(
        '[data-document-chunk-window-container="true"]',
      );
      const finalContainerRect = finalContainer?.getBoundingClientRect();
      const paginationWidgets = Array.from(
        editor?.querySelectorAll<HTMLElement>(
          '.work-document-table-page-break, .work-document-table-cell-page-break',
        ) ?? [],
      );
      return {
        ...windowDiagnostics,
        containerCount: containers.length,
        containerOffsetHeightSum: containers.reduce(
          (height, container) => height + container.offsetHeight,
          0,
        ),
        containerSamples: [
          containers[0],
          containers[Math.floor(containers.length / 2)],
          containers.at(-1),
        ]
          .filter((container): container is HTMLElement => Boolean(container))
          .map((container) => ({
            height: container.style.height,
            id: container.dataset.documentChunkId,
            offsetHeight: container.offsetHeight,
            offsetTop: container.offsetTop,
          })),
        endVisible:
          lastRowRect !== undefined &&
          scrollRect !== undefined &&
          lastRowRect.bottom >= scrollRect.top &&
          lastRowRect.top <= scrollRect.bottom,
        finalChunkBottom: finalChunkRect?.bottom ?? null,
        finalChunkHeight: finalChunkRect?.height ?? null,
        finalChunkMounted: finalChunk?.dataset.documentChunkMounted === 'true',
        finalChunkTop: finalChunkRect?.top ?? null,
        finalContainerBottom: finalContainerRect?.bottom ?? null,
        finalContainerHeight: finalContainerRect?.height ?? null,
        finalContainerChildren: Array.from(finalContainer?.children ?? []).map(
          (child) => {
            const element = child as HTMLElement;
            const rect = element.getBoundingClientRect();
            const style = getComputedStyle(element);
            return {
              className: element.className,
              display: style.display,
              left: rect.left,
              id: element.dataset.documentChunkId ?? null,
              mounted: element.dataset.documentChunkMounted ?? null,
              offsetHeight: element.offsetHeight,
              offsetParentClassName:
                (element.offsetParent as HTMLElement | null)?.className ?? null,
              offsetParentId:
                (element.offsetParent as HTMLElement | null)?.dataset
                  .documentChunkId ?? null,
              offsetTop: element.offsetTop,
              paginationExtraHeight:
                Number(element.dataset.documentChunkPaginationExtraHeight) || 0,
              tagName: element.tagName,
              top: rect.top,
              width: rect.width,
              widgetCount: element.querySelectorAll(
                '.work-document-table-page-break, .work-document-table-cell-page-break',
              ).length,
            };
          },
        ),
        finalContainerPaginationWidgetCount:
          finalContainer?.querySelectorAll(
            '.work-document-table-page-break, .work-document-table-cell-page-break',
          ).length ?? 0,
        finalContainerTop: finalContainerRect?.top ?? null,
        sectionGeometry: (() => {
          const section = finalContainer?.closest<HTMLElement>(
            '.work-document-section',
          );
          if (!section) return null;
          const rect = section.getBoundingClientRect();
          const style = getComputedStyle(section);
          return {
            clientHeight: section.clientHeight,
            clientWidth: section.clientWidth,
            columnCount: style.columnCount,
            columnFill: style.columnFill,
            height: style.height,
            offsetHeight: section.offsetHeight,
            rectHeight: rect.height,
            rectLeft: rect.left,
            rectTop: rect.top,
            scrollHeight: section.scrollHeight,
            scrollWidth: section.scrollWidth,
            width: style.width,
          };
        })(),
        finalLeafOffsetTop: finalChunk?.offsetTop ?? null,
        lastRowBottom: lastRowRect?.bottom ?? null,
        lastRowText:
          lastRow?.textContent?.replaceAll(/\s+/g, ' ').trim() ?? null,
        lastRowTop: lastRowRect?.top ?? null,
        paginationWidgetCount: paginationWidgets.length,
        paginationWidgetSamples: [
          paginationWidgets[0],
          paginationWidgets[Math.floor(paginationWidgets.length / 2)],
          paginationWidgets.at(-1),
        ]
          .filter((widget): widget is HTMLElement => Boolean(widget))
          .map((widget) => ({
            chunkId:
              widget.closest<HTMLElement>('[data-document-chunk-id]')?.dataset
                .documentChunkId ?? null,
            offsetHeight: widget.offsetHeight,
            offsetTop: widget.offsetTop,
            pageIndex: widget.dataset.pageIndex ?? null,
            parentClassName:
              (widget.parentElement as HTMLElement | null)?.className ?? null,
            parentTagName: widget.parentElement?.tagName ?? null,
          })),
        scrollBottom: scrollRect?.bottom ?? null,
        scrollHeight: scroll?.scrollHeight ?? null,
        scrollTop: scroll?.scrollTop ?? null,
        scrollViewportTop: scrollRect?.top ?? null,
      };
    }
    const paragraphs = editor?.getElementsByTagName('p');
    const lastParagraph = paragraphs?.item(Math.max(0, paragraphs.length - 1));
    const lastParagraphRect = lastParagraph?.getBoundingClientRect();
    const chunks = Array.from(
      editor?.querySelectorAll<HTMLElement>('[data-document-chunk-id]') ?? [],
    );
    const finalChunk = chunks.at(-1);
    const finalChunkRect = finalChunk?.getBoundingClientRect();
    const page = editor?.closest<HTMLElement>('.work-document-page');
    const pageRect = page?.getBoundingClientRect();
    const stack = page?.querySelector<HTMLElement>('.work-document-page-stack');
    return {
      ...windowDiagnostics,
      endVisible:
        lastParagraphRect !== undefined &&
        scrollRect !== undefined &&
        lastParagraphRect.bottom >= scrollRect.top &&
        lastParagraphRect.top <= scrollRect.bottom,
      finalChunkBottom: finalChunkRect?.bottom ?? null,
      finalChunkExtraHeight:
        Number(finalChunk?.dataset.documentChunkPaginationExtraHeight) || 0,
      finalChunkHeight: finalChunkRect?.height ?? null,
      finalChunkTop: finalChunkRect?.top ?? null,
      chunkCount: chunks.length,
      chunkOffsetHeightSum: chunks.reduce(
        (height, chunk) => height + chunk.offsetHeight,
        0,
      ),
      chunkSamples: [
        chunks[0],
        chunks[Math.floor(chunks.length / 2)],
        finalChunk,
      ]
        .filter((chunk): chunk is HTMLElement => Boolean(chunk))
        .map((chunk) => ({
          display: getComputedStyle(chunk).display,
          height: chunk.style.height,
          id: chunk.dataset.documentChunkId,
          mounted: chunk.dataset.documentChunkMounted,
          offsetHeight: chunk.offsetHeight,
          offsetTop: chunk.offsetTop,
          paginationExtraHeight:
            Number(chunk.dataset.documentChunkPaginationExtraHeight) || 0,
          position: getComputedStyle(chunk).position,
        })),
      lastParagraphText: lastParagraph?.textContent ?? null,
      lastParagraphBottom: lastParagraphRect?.bottom ?? null,
      lastParagraphTop: lastParagraphRect?.top ?? null,
      pageBottom: pageRect?.bottom ?? null,
      pageHeight: pageRect?.height ?? null,
      pageSurfaceHeight: Number(stack?.dataset.pageSurfaceHeight) || null,
      pageTop: pageRect?.top ?? null,
      scrollBottom: scrollRect?.bottom ?? null,
      scrollClientHeight: scroll?.clientHeight ?? null,
      scrollHeight: scroll?.scrollHeight ?? null,
      scrollTop: scroll?.scrollTop ?? null,
      scrollViewportTop: scrollRect?.top ?? null,
    };
  }, kind);
}
