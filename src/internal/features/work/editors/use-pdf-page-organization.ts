import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type {
  PdfPageOrganizationDiagnostic,
  PdfPageOrganizationErrorCode,
  PdfPageOrganizationExportOperation,
  PdfPageOrganizationJob,
  PdfPageOrganizationMutation,
  PdfPageOrganizationResult,
} from './pdf-page-organization';
import { runPdfPageOrganizationJob } from './pdf-page-organization-worker-client';

export interface PdfPageOrganizationExport {
  fileName: string;
  pageCount: number;
  pdf: Blob;
}

export interface PdfPageOrganizationControllerError {
  code: PdfPageOrganizationErrorCode;
  message: string;
}

export interface PdfPageOrganizationControllerState {
  available: boolean;
  busy: boolean;
  canRedo: boolean;
  canUndo: boolean;
  diagnostics: readonly PdfPageOrganizationDiagnostic[];
  error: PdfPageOrganizationControllerError | null;
  revision: number;
}

export interface PdfPageOrganizationController {
  state: PdfPageOrganizationControllerState;
  dismissError: () => void;
  exportPages: (
    operation: PdfPageOrganizationExportOperation,
  ) => Promise<boolean>;
  mutate: (
    mutation: PdfPageOrganizationMutation,
    importSource?: Blob,
  ) => Promise<boolean>;
  redo: () => void;
  undo: () => void;
}

interface PdfPageOrganizationHistoryEntry {
  after: Blob;
  before: Blob;
}

interface UsePdfPageOrganizationOptions {
  enabled: boolean;
  fileName: string;
  onExport?: (
    files: readonly PdfPageOrganizationExport[],
  ) => boolean | Promise<boolean>;
  readCurrentSource: () => Promise<Blob>;
  replaceSource: (source: Blob) => void;
  resetKey?: string;
}

type RunPdfPageOrganizationJob = (
  job: PdfPageOrganizationJob,
  signal?: AbortSignal,
) => Promise<PdfPageOrganizationResult>;

const initialState: PdfPageOrganizationControllerState = {
  available: false,
  busy: false,
  canRedo: false,
  canUndo: false,
  diagnostics: [],
  error: null,
  revision: 0,
};

export function usePdfPageOrganization(
  options: UsePdfPageOrganizationOptions,
  runJob: RunPdfPageOrganizationJob = runPdfPageOrganizationJob,
): PdfPageOrganizationController {
  const [state, setState] = useState<PdfPageOrganizationControllerState>({
    ...initialState,
    available: options.enabled,
  });
  const undoRef = useRef<PdfPageOrganizationHistoryEntry[]>([]);
  const redoRef = useRef<PdfPageOrganizationHistoryEntry[]>([]);
  const abortRef = useRef<AbortController | null>(null);
  const optionsRef = useRef(options);
  const runJobRef = useRef(runJob);
  optionsRef.current = options;
  runJobRef.current = runJob;

  useEffect(() => {
    abortRef.current?.abort();
    abortRef.current = null;
    undoRef.current = [];
    redoRef.current = [];
    setState({ ...initialState, available: options.enabled });
  }, [options.enabled, options.resetKey]);

  useEffect(
    () => () => {
      abortRef.current?.abort();
    },
    [],
  );

  const mutate = useCallback(
    async (
      mutation: PdfPageOrganizationMutation,
      importSource?: Blob,
    ): Promise<boolean> => {
      const currentOptions = optionsRef.current;
      if (!currentOptions.enabled || abortRef.current) return false;
      const abort = new AbortController();
      abortRef.current = abort;
      setState((current) => ({
        ...current,
        busy: true,
        diagnostics: [],
        error: null,
      }));
      try {
        const before = await currentOptions.readCurrentSource();
        const job: PdfPageOrganizationJob = {
          ...(importSource
            ? { importSource: await importSource.arrayBuffer() }
            : {}),
          kind: 'mutate',
          mutation,
          source: await before.arrayBuffer(),
        };
        const result = await runJobRef.current(job, abort.signal);
        if (result.kind !== 'mutated') {
          throw new Error('PDF page mutation returned export files.');
        }
        const after = pdfBlob(result.source);
        const entry = { after, before };
        undoRef.current = [...undoRef.current, entry];
        redoRef.current = [];
        currentOptions.replaceSource(after);
        setState((current) => ({
          ...current,
          busy: false,
          canRedo: false,
          canUndo: true,
          diagnostics: result.diagnostics,
          error: null,
          revision: current.revision + 1,
        }));
        return true;
      } catch (error) {
        if (abort.signal.aborted) return false;
        setState((current) => ({
          ...current,
          busy: false,
          diagnostics: [],
          error: controllerError(error),
        }));
        return false;
      } finally {
        if (abortRef.current === abort) abortRef.current = null;
      }
    },
    [],
  );

  const exportPages = useCallback(
    async (operation: PdfPageOrganizationExportOperation): Promise<boolean> => {
      const currentOptions = optionsRef.current;
      if (!currentOptions.enabled || abortRef.current) return false;
      const abort = new AbortController();
      abortRef.current = abort;
      setState((current) => ({
        ...current,
        busy: true,
        diagnostics: [],
        error: null,
      }));
      try {
        const before = await currentOptions.readCurrentSource();
        const result = await runJobRef.current(
          {
            kind: 'export',
            operation,
            source: await before.arrayBuffer(),
          },
          abort.signal,
        );
        if (result.kind !== 'exported') {
          throw new Error('PDF page export returned a mutated source.');
        }
        const files = namedPdfExports(
          currentOptions.fileName,
          operation.kind,
          result,
        );
        const accepted = currentOptions.onExport
          ? await currentOptions.onExport(files)
          : downloadPdfExports(files);
        if (!accepted)
          throw new Error('The host rejected the PDF page export.');
        setState((current) => ({
          ...current,
          busy: false,
          diagnostics: result.diagnostics,
          error: null,
        }));
        return true;
      } catch (error) {
        if (abort.signal.aborted) return false;
        setState((current) => ({
          ...current,
          busy: false,
          diagnostics: [],
          error: controllerError(error),
        }));
        return false;
      } finally {
        if (abortRef.current === abort) abortRef.current = null;
      }
    },
    [],
  );

  const undo = useCallback(() => {
    if (!optionsRef.current.enabled || abortRef.current) return;
    const entry = undoRef.current.at(-1);
    if (!entry) return;
    undoRef.current = undoRef.current.slice(0, -1);
    redoRef.current = [...redoRef.current, entry];
    optionsRef.current.replaceSource(entry.before);
    setState((current) => ({
      ...current,
      canRedo: true,
      canUndo: undoRef.current.length > 0,
      diagnostics: [],
      error: null,
      revision: current.revision + 1,
    }));
  }, []);

  const redo = useCallback(() => {
    if (!optionsRef.current.enabled || abortRef.current) return;
    const entry = redoRef.current.at(-1);
    if (!entry) return;
    redoRef.current = redoRef.current.slice(0, -1);
    undoRef.current = [...undoRef.current, entry];
    optionsRef.current.replaceSource(entry.after);
    setState((current) => ({
      ...current,
      canRedo: redoRef.current.length > 0,
      canUndo: true,
      diagnostics: [],
      error: null,
      revision: current.revision + 1,
    }));
  }, []);

  const dismissError = useCallback(() => {
    setState((current) => ({ ...current, error: null }));
  }, []);

  return useMemo(
    () => ({ state, dismissError, exportPages, mutate, redo, undo }),
    [dismissError, exportPages, mutate, redo, state, undo],
  );
}

function namedPdfExports(
  fileName: string,
  kind: PdfPageOrganizationExportOperation['kind'],
  result: Extract<PdfPageOrganizationResult, { kind: 'exported' }>,
): PdfPageOrganizationExport[] {
  const stem = fileName.replace(/\.pdf$/i, '') || 'document';
  return result.files.map((file, index) => ({
    fileName:
      kind === 'extract'
        ? `${stem}-extracted.pdf`
        : `${stem}-part-${index + 1}.pdf`,
    pageCount: file.pageCount,
    pdf: pdfBlob(file.source),
  }));
}

function pdfBlob(source: Uint8Array): Blob {
  const copy = new Uint8Array(source.byteLength);
  copy.set(source);
  return new Blob([copy.buffer], { type: 'application/pdf' });
}

function downloadPdfExports(
  files: readonly PdfPageOrganizationExport[],
): boolean {
  if (
    typeof document === 'undefined' ||
    typeof URL.createObjectURL !== 'function'
  ) {
    return false;
  }
  for (const file of files) {
    const url = URL.createObjectURL(file.pdf);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = file.fileName;
    anchor.hidden = true;
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
    queueMicrotask(() => URL.revokeObjectURL(url));
  }
  return true;
}

function controllerError(error: unknown): PdfPageOrganizationControllerError {
  const candidate = error as { code?: unknown; message?: unknown };
  return {
    code:
      typeof candidate?.code === 'string' &&
      candidate.code.startsWith('pdf.pages.')
        ? (candidate.code as PdfPageOrganizationErrorCode)
        : 'pdf.pages.invalid-source',
    message:
      typeof candidate?.message === 'string' && candidate.message
        ? candidate.message.slice(0, 512)
        : 'Unable to organize this PDF.',
  };
}
