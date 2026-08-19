export type WorkFileImportStage =
  | 'reading'
  | 'parsing'
  | 'analyzing'
  | 'finalizing';

export interface WorkFileImportProgress {
  stage: WorkFileImportStage;
  stageProgress: number;
  progress: number;
  bytesRead: number;
  totalBytes: number;
}

export interface WorkFileImportOptions {
  /** Reuses a host-reserved identity so an editor shell can mount in parallel. */
  artifactId?: string;
  /** Reuses host-reserved worksheet identities during a spreadsheet import. */
  spreadsheetSheetIds?: readonly string[];
  signal?: AbortSignal;
  onProgress?: (progress: WorkFileImportProgress) => void;
}

export interface WorkFileImportContext {
  bytes: ArrayBuffer;
  controller: WorkFileImportController;
  spreadsheetSheetIds?: readonly string[];
}

const stageRanges: Record<WorkFileImportStage, readonly [number, number]> = {
  reading: [0, 0.2],
  parsing: [0.2, 0.75],
  analyzing: [0.75, 0.95],
  finalizing: [0.95, 1],
};

export class WorkFileImportController {
  private progress = 0;

  constructor(
    private readonly options: WorkFileImportOptions,
    readonly totalBytes: number,
  ) {}

  get signal(): AbortSignal | undefined {
    return this.options.signal;
  }

  throwIfAborted(): void {
    if (!this.signal?.aborted) return;
    throw workFileImportAbortError(this.signal.reason);
  }

  report(
    stage: WorkFileImportStage,
    stageProgress: number,
    bytesRead = this.totalBytes,
  ): void {
    this.throwIfAborted();
    const boundedStageProgress = Math.max(0, Math.min(1, stageProgress));
    const [start, end] = stageRanges[stage];
    this.progress = Math.max(
      this.progress,
      start + (end - start) * boundedStageProgress,
    );
    this.options.onProgress?.({
      stage,
      stageProgress: boundedStageProgress,
      progress: this.progress,
      bytesRead: Math.max(0, Math.min(this.totalBytes, bytesRead)),
      totalBytes: this.totalBytes,
    });
    this.throwIfAborted();
  }

  async checkpoint(
    stage: WorkFileImportStage,
    stageProgress: number,
  ): Promise<void> {
    this.report(stage, stageProgress);
    await this.yieldToMainThread();
  }

  async yieldToMainThread(): Promise<void> {
    this.throwIfAborted();
    const taskScheduler = (
      globalThis as typeof globalThis & {
        scheduler?: { yield?: () => Promise<void> };
      }
    ).scheduler;
    if (typeof taskScheduler?.yield === 'function') {
      await taskScheduler.yield();
    } else {
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
    }
    this.throwIfAborted();
  }

  complete(): void {
    this.report('finalizing', 1);
  }
}

function workFileImportAbortError(reason: unknown): Error {
  if (reason instanceof Error && reason.name === 'AbortError') return reason;
  if (typeof DOMException !== 'undefined') {
    return new DOMException('Office file import was cancelled.', 'AbortError');
  }
  const error = new Error('Office file import was cancelled.');
  error.name = 'AbortError';
  return error;
}
