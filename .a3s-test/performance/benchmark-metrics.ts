import type { CDPSession } from '@playwright/test';

export interface MetricSnapshot {
  documents: number;
  frames: number;
  jsEventListeners: number;
  jsHeapTotalBytes: number;
  jsHeapUsedBytes: number;
  layoutCount: number;
  layoutDurationMs: number;
  nodes: number;
  recalcStyleCount: number;
  recalcStyleDurationMs: number;
  scriptDurationMs: number;
  taskDurationMs: number;
}

export interface LongTaskSummary {
  count: number;
  maximumMs: number;
  p95Ms: number;
  totalMs: number;
}

export async function collectMetrics(cdp: CDPSession): Promise<MetricSnapshot> {
  const [performanceMetrics, domCounters] = await Promise.all([
    cdp.send('Performance.getMetrics') as Promise<{
      metrics: Array<{ name: string; value: number }>;
    }>,
    cdp.send('Memory.getDOMCounters') as Promise<{
      documents: number;
      jsEventListeners: number;
      nodes: number;
    }>,
  ]);
  const metrics = Object.fromEntries(
    performanceMetrics.metrics.map(({ name, value }) => [name, value]),
  );
  return {
    documents: domCounters.documents,
    frames: metrics.Frames ?? 0,
    jsEventListeners: domCounters.jsEventListeners,
    jsHeapTotalBytes: Math.round(metrics.JSHeapTotalSize ?? 0),
    jsHeapUsedBytes: Math.round(metrics.JSHeapUsedSize ?? 0),
    layoutCount: metrics.LayoutCount ?? 0,
    layoutDurationMs: round((metrics.LayoutDuration ?? 0) * 1000),
    nodes: domCounters.nodes,
    recalcStyleCount: metrics.RecalcStyleCount ?? 0,
    recalcStyleDurationMs: round((metrics.RecalcStyleDuration ?? 0) * 1000),
    scriptDurationMs: round((metrics.ScriptDuration ?? 0) * 1000),
    taskDurationMs: round((metrics.TaskDuration ?? 0) * 1000),
  };
}

export function subtractMetrics(
  after: MetricSnapshot,
  before: MetricSnapshot,
): MetricSnapshot {
  return Object.fromEntries(
    Object.entries(after).map(([key, value]) => [
      key,
      round(value - before[key as keyof MetricSnapshot]),
    ]),
  ) as unknown as MetricSnapshot;
}

export function summarizeDurations(durations: number[]): LongTaskSummary {
  const sorted = [...durations].sort((left, right) => left - right);
  return {
    count: sorted.length,
    maximumMs: round(sorted.at(-1) ?? 0),
    p95Ms: round(percentile(sorted, 0.95)),
    totalMs: round(sorted.reduce((sum, value) => sum + value, 0)),
  };
}

export function percentile(sorted: number[], ratio: number): number {
  if (!sorted.length) return 0;
  const index = Math.min(
    sorted.length - 1,
    Math.ceil(sorted.length * ratio) - 1,
  );
  return sorted[Math.max(0, index)] ?? 0;
}

export function round(value: number): number {
  return Math.round(value * 10) / 10;
}

export function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export async function within<T>(
  promise: Promise<T>,
  timeout: number,
): Promise<T> {
  let timeoutIdentifier: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      promise,
      new Promise<T>((_, reject) => {
        timeoutIdentifier = setTimeout(
          () => reject(new Error(`Diagnostic timed out after ${timeout}ms.`)),
          timeout,
        );
      }),
    ]);
  } finally {
    if (timeoutIdentifier !== undefined) clearTimeout(timeoutIdentifier);
  }
}
