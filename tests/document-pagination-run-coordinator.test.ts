import { expect, test } from '@rstest/core';
import { createDocumentPaginationRunCoordinator } from '../src/internal/features/work/editors/document-pagination-run-coordinator';

test('coalesces repeated layout observations into one follow-up run', async () => {
  const frames = createTestFrames();
  const runs: AbortSignal[] = [];
  const completions: Array<() => void> = [];
  let coalescedRequests = 0;
  const coordinator = createDocumentPaginationRunCoordinator({
    ...frames.options,
    onCoalescedRequest: () => {
      coalescedRequests += 1;
    },
    onError: (error) => {
      throw error;
    },
    run: (signal) => {
      runs.push(signal);
      return new Promise<void>((resolve) => completions.push(resolve));
    },
  });

  coordinator.request();
  frames.flushAll();
  expect(runs).toHaveLength(1);

  for (let index = 0; index < 2_000; index += 1) coordinator.request();

  expect(coalescedRequests).toBe(2_000);
  expect(runs[0].aborted).toBe(false);
  expect(frames.pending()).toBe(0);

  completions[0]();
  await settlePromises();
  expect(frames.pending()).toBe(1);

  frames.flushAll();
  expect(runs).toHaveLength(2);
  completions[1]();
  await settlePromises();
  expect(frames.pending()).toBe(0);
  coordinator.dispose();
});

test('aborts stale work only when an active document is invalidated', async () => {
  const frames = createTestFrames();
  const runs: AbortSignal[] = [];
  const abortReasons: string[] = [];
  const coordinator = createDocumentPaginationRunCoordinator({
    ...frames.options,
    onAbort: (reason) => abortReasons.push(reason),
    onError: (error) => {
      throw error;
    },
    run: (signal) => {
      runs.push(signal);
      return new Promise<void>((resolve) => {
        signal.addEventListener('abort', () => resolve(), { once: true });
      });
    },
  });

  coordinator.request();
  frames.flushAll();
  coordinator.request();
  expect(runs[0].aborted).toBe(false);

  coordinator.request({ invalidateActive: true });
  coordinator.request({ invalidateActive: true });
  expect(runs[0].aborted).toBe(true);
  expect(abortReasons).toEqual(['invalidation']);

  await settlePromises();
  frames.flushAll();
  expect(runs).toHaveLength(2);
  coordinator.dispose();
  expect(runs[1].aborted).toBe(true);
  expect(abortReasons).toEqual(['invalidation', 'disposal']);
});

function createTestFrames() {
  let nextFrame = 1;
  const callbacks = new Map<number, () => void>();
  return {
    options: {
      cancelFrame: (frame: number) => {
        callbacks.delete(frame);
      },
      requestFrame: (callback: () => void) => {
        const frame = nextFrame;
        nextFrame += 1;
        callbacks.set(frame, callback);
        return frame;
      },
    },
    flushAll: () => {
      while (callbacks.size) {
        const [frame, callback] = callbacks.entries().next().value as [
          number,
          () => void,
        ];
        callbacks.delete(frame);
        callback();
      }
    },
    pending: () => callbacks.size,
  };
}

async function settlePromises(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}
