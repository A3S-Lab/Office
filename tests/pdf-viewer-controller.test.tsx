import type { PluginRegistry } from '@embedpdf/react-pdf-viewer';
import { expect, test } from '@rstest/core';
import { act, renderHook, waitFor } from '@testing-library/react';
import { usePdfViewerController } from '../src/internal/features/work/editors/pdf-viewer-controller';

test('routes PDF commands through typed plugin capabilities', async () => {
  let currentPage = 2;
  const calls: string[] = [];
  const pageChange = createEvent();
  const layoutReady = createEvent();
  const zoomChange = createEvent();
  const searchChange = createEvent();
  const historyChange = createEvent();
  const activeDocumentChange = createEvent();
  const documentOpened = createEvent();
  const documentClosed = createEvent();

  const capabilities = {
    'document-manager': {
      getActiveDocumentId: () => 'document-1',
      onActiveDocumentChanged: activeDocumentChange.subscribe,
      onDocumentOpened: documentOpened.subscribe,
      onDocumentClosed: documentClosed.subscribe,
    },
    scroll: {
      forDocument: () => ({
        getCurrentPage: () => currentPage,
        getTotalPages: () => 8,
        scrollToPreviousPage: () => calls.push('previous-page'),
        scrollToNextPage: () => calls.push('next-page'),
        scrollToPage: ({ pageNumber }: { pageNumber: number }) =>
          calls.push(`page:${pageNumber}`),
      }),
      onPageChange: pageChange.subscribe,
      onLayoutReady: layoutReady.subscribe,
    },
    zoom: {
      forDocument: () => ({
        getState: () => ({
          zoomLevel: 'fit-width',
          currentZoomLevel: 1.25,
          isMarqueeZoomActive: false,
        }),
        zoomOut: () => calls.push('zoom-out'),
        zoomIn: () => calls.push('zoom-in'),
        requestZoom: (mode: string) => calls.push(`zoom:${mode}`),
      }),
      onStateChange: zoomChange.subscribe,
    },
    search: {
      forDocument: () => ({
        getState: () => ({
          flags: [],
          results: [],
          total: 3,
          activeResultIndex: 1,
          showAllResults: true,
          query: 'A3S',
          loading: false,
          active: true,
        }),
        searchAllPages: (query: string) => {
          calls.push(`search:${query}`);
          return {
            toPromise: () => Promise.resolve({ results: [], total: 0 }),
          };
        },
        stopSearch: () => calls.push('stop-search'),
        previousResult: () => calls.push('previous-result'),
        nextResult: () => calls.push('next-result'),
      }),
      onStateChange: searchChange.subscribe,
    },
    history: {
      forDocument: () => ({
        canUndo: () => true,
        canRedo: () => false,
        undo: () => calls.push('undo'),
        redo: () => calls.push('redo'),
      }),
      onHistoryChange: historyChange.subscribe,
    },
    export: {
      forDocument: () => ({
        saveAsCopy: () => ({
          toPromise: () => Promise.resolve(new Uint8Array([37, 80, 68, 70])),
        }),
      }),
    },
  };

  const registry = {
    pluginsReady: () => Promise.resolve(),
    getPlugin: (id: keyof typeof capabilities) => ({
      provides: () => capabilities[id],
    }),
  } as unknown as PluginRegistry;

  const { result } = renderHook(() => usePdfViewerController(registry));

  await waitFor(() => expect(result.current.state.ready).toBe(true));
  expect(result.current.state).toMatchObject({
    canUndo: true,
    canRedo: false,
    currentPage: 2,
    totalPages: 8,
    zoomPercent: 125,
    zoomMode: 'fit-width',
    search: {
      activeResultIndex: 1,
      query: 'A3S',
      total: 3,
    },
  });

  act(() => {
    result.current.previousPage();
    result.current.nextPage();
    result.current.goToPage(6);
    result.current.zoomOut();
    result.current.zoomIn();
    result.current.fitPage();
    result.current.fitWidth();
    result.current.search('roadmap');
    result.current.previousSearchResult();
    result.current.nextSearchResult();
    result.current.undo();
    result.current.redo();
  });

  expect(calls).toEqual([
    'previous-page',
    'next-page',
    'page:6',
    'zoom-out',
    'zoom-in',
    'zoom:fit-page',
    'zoom:fit-width',
    'search:roadmap',
    'previous-result',
    'next-result',
    'undo',
    'redo',
  ]);

  const copy = await result.current.saveAsCopy();
  expect(copy.type).toBe('application/pdf');
  expect(copy.size).toBe(4);

  currentPage = 5;
  act(() => pageChange.emit(undefined));
  await waitFor(() => expect(result.current.state.currentPage).toBe(5));
});

function createEvent<T = unknown>() {
  const listeners = new Set<(value: T) => void>();
  return {
    emit: (value: T) => {
      for (const listener of listeners) listener(value);
    },
    subscribe: (listener: (value: T) => void) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
}
