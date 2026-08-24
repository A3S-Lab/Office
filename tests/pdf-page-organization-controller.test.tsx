import { expect, rstest, test } from '@rstest/core';
import { act, renderHook } from '@testing-library/react';
import {
  usePdfPageOrganization,
  type PdfPageOrganizationExport,
} from '../src/internal/features/work/editors/use-pdf-page-organization';
import type {
  PdfPageOrganizationJob,
  PdfPageOrganizationResult,
} from '../src/internal/features/work/editors/pdf-page-organization';

test('commits one source replacement and one undo record per page intent', async () => {
  let current = pdfBlob('source');
  const applied: Blob[] = [];
  const run = rstest.fn(async (_job: PdfPageOrganizationJob) =>
    mutationResult('after'),
  );
  const { result } = renderHook(() =>
    usePdfPageOrganization(
      {
        enabled: true,
        fileName: 'report.pdf',
        readCurrentSource: () => Promise.resolve(current),
        replaceSource: (source) => {
          current = source;
          applied.push(source);
        },
        resetKey: 'source-1',
      },
      run,
    ),
  );

  await act(() =>
    result.current.mutate({
      degrees: 90,
      kind: 'rotate',
      pageIndexes: [1, 2],
    }),
  );
  expect(run).toHaveBeenCalledTimes(1);
  expect(result.current.state).toMatchObject({
    busy: false,
    canRedo: false,
    canUndo: true,
    revision: 1,
  });
  expect(await applied[0]?.text()).toBe('after');

  act(() => result.current.undo());
  expect(await applied[1]?.text()).toBe('source');
  expect(result.current.state).toMatchObject({ canRedo: true, canUndo: false });

  act(() => result.current.redo());
  expect(await applied[2]?.text()).toBe('after');
  expect(result.current.state).toMatchObject({ canRedo: false, canUndo: true });
});

test('exports exact worker files without adding a history record', async () => {
  const exported: PdfPageOrganizationExport[][] = [];
  const run = rstest.fn(async () => exportResult(['part-1', 'part-2']));
  const { result } = renderHook(() =>
    usePdfPageOrganization(
      {
        enabled: true,
        fileName: 'report.pdf',
        onExport: async (files) => {
          exported.push([...files]);
          return true;
        },
        readCurrentSource: () => Promise.resolve(pdfBlob('source')),
        replaceSource: () => undefined,
        resetKey: 'source-1',
      },
      run,
    ),
  );

  await act(() =>
    result.current.exportPages({
      kind: 'split',
      splitAfterPageIndexes: [0],
    }),
  );
  expect(result.current.state).toMatchObject({ canUndo: false, revision: 0 });
  expect(exported[0]?.map((file) => file.fileName)).toEqual([
    'report-part-1.pdf',
    'report-part-2.pdf',
  ]);
  expect(await exported[0]?.[0]?.pdf.text()).toBe('part-1');
});

test('clears page history when the authoritative host source changes', async () => {
  let current = pdfBlob('source');
  const { result, rerender } = renderHook(
    ({ resetKey }) =>
      usePdfPageOrganization(
        {
          enabled: true,
          fileName: 'report.pdf',
          readCurrentSource: () => Promise.resolve(current),
          replaceSource: (source) => {
            current = source;
          },
          resetKey,
        },
        async () => mutationResult('after'),
      ),
    { initialProps: { resetKey: 'source-1' } },
  );
  await act(() => result.current.mutate({ kind: 'insert-blank', index: 1 }));
  expect(result.current.state.canUndo).toBe(true);
  rerender({ resetKey: 'source-2' });
  expect(result.current.state).toMatchObject({
    canRedo: false,
    canUndo: false,
    revision: 0,
  });
});

test('keeps the current source unchanged and exposes fail-closed diagnostics', async () => {
  const replaceSource = rstest.fn();
  const { result } = renderHook(() =>
    usePdfPageOrganization(
      {
        enabled: true,
        fileName: 'report.pdf',
        readCurrentSource: () => Promise.resolve(pdfBlob('source')),
        replaceSource,
        resetKey: 'source-1',
      },
      async () => {
        throw Object.assign(new Error('Signed source is immutable.'), {
          code: 'pdf.pages.signed-source',
        });
      },
    ),
  );
  await act(() =>
    result.current.mutate({ degrees: 90, kind: 'rotate', pageIndexes: [0] }),
  );
  expect(replaceSource).not.toHaveBeenCalled();
  expect(result.current.state.error).toMatchObject({
    code: 'pdf.pages.signed-source',
    message: 'Signed source is immutable.',
  });
  act(() => result.current.dismissError());
  expect(result.current.state.error).toBeNull();
});

function pdfBlob(text: string): Blob {
  return new Blob([text], { type: 'application/pdf' });
}

function mutationResult(text: string): PdfPageOrganizationResult {
  return {
    diagnostics: [],
    kind: 'mutated',
    pageCount: 2,
    source: new TextEncoder().encode(text),
  };
}

function exportResult(texts: string[]): PdfPageOrganizationResult {
  return {
    diagnostics: [],
    files: texts.map((text) => ({
      pageCount: 1,
      source: new TextEncoder().encode(text),
    })),
    kind: 'exported',
  };
}
