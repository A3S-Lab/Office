export interface DocumentPaginationRunRequest {
  invalidateActive?: boolean;
}

export interface DocumentPaginationRunCoordinatorOptions {
  cancelFrame: (frame: number) => void;
  onAbort?: (reason: 'invalidation' | 'disposal') => void;
  onCoalescedRequest?: () => void;
  onError: (error: unknown) => void;
  onRunFinish?: (aborted: boolean) => void;
  onRunStart?: () => void;
  requestFrame: (callback: () => void) => number;
  run: (signal: AbortSignal) => Promise<void>;
}

export interface DocumentPaginationRunCoordinator {
  dispose: () => void;
  request: (request?: DocumentPaginationRunRequest) => void;
}

/**
 * Runs document pagination as a single-flight operation.
 *
 * Layout observations that arrive during a run are merged into one follow-up
 * pass. Only document invalidation aborts active work, so observer and font
 * events cannot continuously starve a large document before it reaches a
 * stable layout.
 */
export function createDocumentPaginationRunCoordinator({
  cancelFrame,
  onAbort,
  onCoalescedRequest,
  onError,
  onRunFinish,
  onRunStart,
  requestFrame,
  run,
}: DocumentPaginationRunCoordinatorOptions): DocumentPaginationRunCoordinator {
  let activeController: AbortController | null = null;
  let disposed = false;
  let frame = 0;
  let rerunRequested = false;
  let running = false;

  const queue = () => {
    if (disposed || frame || running) return;
    frame = requestFrame(() => {
      if (disposed) return;
      frame = requestFrame(() => {
        frame = 0;
        if (!disposed) void execute();
      });
    });
  };

  const execute = async () => {
    if (disposed || running) return;
    running = true;
    const controller = new AbortController();
    activeController = controller;
    onRunStart?.();
    try {
      await run(controller.signal);
    } catch (error) {
      if (!controller.signal.aborted) onError(error);
    } finally {
      if (activeController === controller) activeController = null;
      running = false;
      onRunFinish?.(controller.signal.aborted);
      if (!disposed && rerunRequested) {
        rerunRequested = false;
        queue();
      }
    }
  };

  const request = ({
    invalidateActive = false,
  }: DocumentPaginationRunRequest = {}) => {
    if (disposed) return;
    if (!running) {
      queue();
      return;
    }

    rerunRequested = true;
    onCoalescedRequest?.();
    if (
      invalidateActive &&
      activeController &&
      !activeController.signal.aborted
    ) {
      onAbort?.('invalidation');
      activeController.abort();
    }
  };

  const dispose = () => {
    if (disposed) return;
    disposed = true;
    rerunRequested = false;
    if (frame) {
      cancelFrame(frame);
      frame = 0;
    }
    if (activeController && !activeController.signal.aborted) {
      onAbort?.('disposal');
      activeController.abort();
    }
    activeController = null;
  };

  return { dispose, request };
}
