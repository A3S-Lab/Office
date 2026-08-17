export interface DocumentPaginationMeasurementPass {
  readonly from: number;
}

export interface DocumentPaginationMeasurementRange {
  begin: () => DocumentPaginationMeasurementPass;
  commit: (pass: DocumentPaginationMeasurementPass) => void;
  ensureDirty: () => void;
  invalidate: (position: number) => void;
  reset: () => void;
  restore: (pass: DocumentPaginationMeasurementPass) => void;
  restoreActive: () => void;
}

/**
 * Tracks the earliest document position that the next pagination pass must
 * measure. An active pass consumes that position temporarily, then either
 * commits it or restores it when cooperative work is aborted.
 */
export function createDocumentPaginationMeasurementRange(): DocumentPaginationMeasurementRange {
  let activePass: DocumentPaginationMeasurementPass | null = null;
  let dirtyFrom = 0;

  const restore = (pass: DocumentPaginationMeasurementPass) => {
    if (activePass !== pass) return;
    dirtyFrom = Math.min(dirtyFrom, pass.from);
    activePass = null;
  };

  return {
    begin: () => {
      const pass = {
        from: Number.isFinite(dirtyFrom) ? dirtyFrom : 0,
      };
      activePass = pass;
      dirtyFrom = Number.POSITIVE_INFINITY;
      return pass;
    },
    commit: (pass) => {
      if (activePass === pass) activePass = null;
    },
    ensureDirty: () => {
      if (Number.isFinite(dirtyFrom)) return;
      dirtyFrom = activePass?.from ?? 0;
    },
    invalidate: (position) => {
      const normalizedPosition = Number.isFinite(position)
        ? Math.max(0, position)
        : 0;
      dirtyFrom = Math.min(dirtyFrom, normalizedPosition);
    },
    reset: () => {
      activePass = null;
      dirtyFrom = 0;
    },
    restore,
    restoreActive: () => {
      if (activePass) restore(activePass);
    },
  };
}
