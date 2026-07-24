import {
  DocumentManagerPlugin,
  type DocumentManagerCapability,
  ExportPlugin,
  type ExportCapability,
  HistoryPlugin,
  type HistoryCapability,
  type PluginRegistry,
  SearchPlugin,
  type SearchCapability,
  ScrollPlugin,
  type ScrollCapability,
  ZoomMode,
  ZoomPlugin,
  type ZoomCapability,
} from '@embedpdf/react-pdf-viewer';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';

export interface PdfSearchState {
  active: boolean;
  activeResultIndex: number;
  error: boolean;
  loading: boolean;
  query: string;
  total: number;
}

export interface PdfViewerFeatures {
  export: boolean;
  history: boolean;
  navigation: boolean;
  search: boolean;
  zoom: boolean;
}

export interface PdfViewerControllerState {
  canRedo: boolean;
  canUndo: boolean;
  currentPage: number;
  documentOpen: boolean;
  error: string | null;
  features: PdfViewerFeatures;
  ready: boolean;
  search: PdfSearchState;
  totalPages: number;
  zoomMode: ZoomMode | null;
  zoomPercent: number;
}

export interface PdfViewerController {
  state: PdfViewerControllerState;
  clearSearch: () => void;
  fitPage: () => void;
  fitWidth: () => void;
  goToPage: (page: number) => void;
  nextPage: () => void;
  nextSearchResult: () => void;
  previousPage: () => void;
  previousSearchResult: () => void;
  redo: () => void;
  saveAsCopy: () => Promise<Blob>;
  search: (query: string) => void;
  undo: () => void;
  zoomIn: () => void;
  zoomOut: () => void;
}

interface PdfViewerCapabilities {
  documentManager: DocumentManagerCapability | null;
  export: ExportCapability | null;
  history: HistoryCapability | null;
  search: SearchCapability | null;
  scroll: ScrollCapability | null;
  zoom: ZoomCapability | null;
}

const EMPTY_FEATURES: PdfViewerFeatures = {
  export: false,
  history: false,
  navigation: false,
  search: false,
  zoom: false,
};

const EMPTY_SEARCH: PdfSearchState = {
  active: false,
  activeResultIndex: -1,
  error: false,
  loading: false,
  query: '',
  total: 0,
};

export const initialPdfViewerControllerState: PdfViewerControllerState = {
  canRedo: false,
  canUndo: false,
  currentPage: 0,
  documentOpen: false,
  error: null,
  features: EMPTY_FEATURES,
  ready: false,
  search: EMPTY_SEARCH,
  totalPages: 0,
  zoomMode: null,
  zoomPercent: 100,
};

export function usePdfViewerController(
  registry: PluginRegistry | null,
): PdfViewerController {
  const capabilitiesRef = useRef<PdfViewerCapabilities | null>(null);
  const [state, setState] = useState<PdfViewerControllerState>(
    initialPdfViewerControllerState,
  );

  useEffect(() => {
    let disposed = false;
    const unsubscribe: Array<() => void> = [];
    capabilitiesRef.current = null;
    setState(initialPdfViewerControllerState);
    if (!registry) return;

    void registry
      .pluginsReady()
      .then(() => {
        if (disposed) return;
        const capabilities = readCapabilities(registry);
        capabilitiesRef.current = capabilities;

        const sync = () => {
          if (!disposed) {
            setState(readControllerState(capabilities));
          }
        };

        subscribe(
          unsubscribe,
          capabilities.documentManager?.onActiveDocumentChanged,
          sync,
        );
        subscribe(
          unsubscribe,
          capabilities.documentManager?.onDocumentOpened,
          sync,
        );
        subscribe(
          unsubscribe,
          capabilities.documentManager?.onDocumentClosed,
          sync,
        );
        subscribe(unsubscribe, capabilities.scroll?.onPageChange, sync);
        subscribe(unsubscribe, capabilities.scroll?.onLayoutReady, sync);
        subscribe(unsubscribe, capabilities.zoom?.onStateChange, sync);
        subscribe(unsubscribe, capabilities.search?.onStateChange, sync);
        subscribe(unsubscribe, capabilities.history?.onHistoryChange, sync);
        sync();
      })
      .catch((error: unknown) => {
        if (disposed) return;
        setState({
          ...initialPdfViewerControllerState,
          error: errorMessage(error),
        });
      });

    return () => {
      disposed = true;
      capabilitiesRef.current = null;
      for (const stop of unsubscribe) stop();
    };
  }, [registry]);

  const previousPage = useCallback(() => {
    withActiveDocument(capabilitiesRef.current, ({ scroll }, documentId) => {
      scroll?.forDocument(documentId).scrollToPreviousPage('smooth');
    });
  }, []);

  const nextPage = useCallback(() => {
    withActiveDocument(capabilitiesRef.current, ({ scroll }, documentId) => {
      scroll?.forDocument(documentId).scrollToNextPage('smooth');
    });
  }, []);

  const goToPage = useCallback((page: number) => {
    withActiveDocument(capabilitiesRef.current, ({ scroll }, documentId) => {
      const scope = scroll?.forDocument(documentId);
      if (!scope) return;
      const totalPages = scope.getTotalPages();
      if (totalPages < 1) return;
      const nextPage = Math.min(totalPages, Math.max(1, Math.round(page)));
      if (Number.isFinite(nextPage)) {
        scope.scrollToPage({ pageNumber: nextPage, behavior: 'smooth' });
      }
    });
  }, []);

  const zoomOut = useCallback(() => {
    withActiveDocument(capabilitiesRef.current, ({ zoom }, documentId) => {
      zoom?.forDocument(documentId).zoomOut();
    });
  }, []);

  const zoomIn = useCallback(() => {
    withActiveDocument(capabilitiesRef.current, ({ zoom }, documentId) => {
      zoom?.forDocument(documentId).zoomIn();
    });
  }, []);

  const fitPage = useCallback(() => {
    withActiveDocument(capabilitiesRef.current, ({ zoom }, documentId) => {
      zoom?.forDocument(documentId).requestZoom(ZoomMode.FitPage);
    });
  }, []);

  const fitWidth = useCallback(() => {
    withActiveDocument(capabilitiesRef.current, ({ zoom }, documentId) => {
      zoom?.forDocument(documentId).requestZoom(ZoomMode.FitWidth);
    });
  }, []);

  const search = useCallback((query: string) => {
    const trimmedQuery = query.trim();
    withActiveDocument(
      capabilitiesRef.current,
      ({ search: searchCapability }, documentId) => {
        const scope = searchCapability?.forDocument(documentId);
        if (!scope) return;
        if (!trimmedQuery) {
          scope.searchAllPages('');
          scope.stopSearch();
          setState((previous) => ({
            ...previous,
            search: EMPTY_SEARCH,
          }));
          return;
        }
        setState((previous) => ({
          ...previous,
          search: {
            ...previous.search,
            error: false,
            loading: true,
            query: trimmedQuery,
          },
        }));
        void scope
          .searchAllPages(trimmedQuery)
          .toPromise()
          .catch(() => {
            setState((previous) =>
              previous.search.query !== trimmedQuery
                ? previous
                : {
                    ...previous,
                    search: {
                      ...previous.search,
                      error: true,
                      loading: false,
                    },
                  },
            );
          });
      },
    );
  }, []);

  const clearSearch = useCallback(() => {
    search('');
  }, [search]);

  const previousSearchResult = useCallback(() => {
    withActiveDocument(
      capabilitiesRef.current,
      ({ search: searchCapability }, documentId) => {
        searchCapability?.forDocument(documentId).previousResult();
      },
    );
  }, []);

  const nextSearchResult = useCallback(() => {
    withActiveDocument(
      capabilitiesRef.current,
      ({ search: searchCapability }, documentId) => {
        searchCapability?.forDocument(documentId).nextResult();
      },
    );
  }, []);

  const undo = useCallback(() => {
    withActiveDocument(capabilitiesRef.current, ({ history }, documentId) => {
      history?.forDocument(documentId).undo();
    });
  }, []);

  const redo = useCallback(() => {
    withActiveDocument(capabilitiesRef.current, ({ history }, documentId) => {
      history?.forDocument(documentId).redo();
    });
  }, []);

  const saveAsCopy = useCallback(async () => {
    const capabilities = capabilitiesRef.current;
    const documentId = activeDocumentId(capabilities);
    if (!capabilities?.export || !documentId) {
      throw new Error('PDF export is unavailable.');
    }
    const buffer = await capabilities.export
      .forDocument(documentId)
      .saveAsCopy()
      .toPromise();
    return new Blob([buffer], { type: 'application/pdf' });
  }, []);

  return useMemo(
    () => ({
      state,
      clearSearch,
      fitPage,
      fitWidth,
      goToPage,
      nextPage,
      nextSearchResult,
      previousPage,
      previousSearchResult,
      redo,
      saveAsCopy,
      search,
      undo,
      zoomIn,
      zoomOut,
    }),
    [
      state,
      clearSearch,
      fitPage,
      fitWidth,
      goToPage,
      nextPage,
      nextSearchResult,
      previousPage,
      previousSearchResult,
      redo,
      saveAsCopy,
      search,
      undo,
      zoomIn,
      zoomOut,
    ],
  );
}

function readCapabilities(registry: PluginRegistry): PdfViewerCapabilities {
  return {
    documentManager:
      registry
        .getPlugin<DocumentManagerPlugin>(DocumentManagerPlugin.id)
        ?.provides() ?? null,
    export:
      registry.getPlugin<ExportPlugin>(ExportPlugin.id)?.provides() ?? null,
    history:
      registry.getPlugin<HistoryPlugin>(HistoryPlugin.id)?.provides() ?? null,
    search:
      registry.getPlugin<SearchPlugin>(SearchPlugin.id)?.provides() ?? null,
    scroll:
      registry.getPlugin<ScrollPlugin>(ScrollPlugin.id)?.provides() ?? null,
    zoom: registry.getPlugin<ZoomPlugin>(ZoomPlugin.id)?.provides() ?? null,
  };
}

function readControllerState(
  capabilities: PdfViewerCapabilities,
): PdfViewerControllerState {
  const documentId = activeDocumentId(capabilities);
  const features = {
    export: Boolean(capabilities.export),
    history: Boolean(capabilities.history),
    navigation: Boolean(capabilities.scroll),
    search: Boolean(capabilities.search),
    zoom: Boolean(capabilities.zoom),
  };
  if (!documentId) {
    return {
      ...initialPdfViewerControllerState,
      features,
      ready: true,
    };
  }

  const scroll = safely(() => capabilities.scroll?.forDocument(documentId));
  const zoom = safely(() => capabilities.zoom?.forDocument(documentId));
  const search = safely(() => capabilities.search?.forDocument(documentId));
  const history = safely(() => capabilities.history?.forDocument(documentId));
  const zoomState = safely(() => zoom?.getState());
  const searchState = safely(() => search?.getState());

  return {
    canRedo: safely(() => history?.canRedo()) ?? false,
    canUndo: safely(() => history?.canUndo()) ?? false,
    currentPage: safely(() => scroll?.getCurrentPage()) ?? 1,
    documentOpen: true,
    error: null,
    features,
    ready: true,
    search: searchState
      ? {
          active: searchState.active,
          activeResultIndex: searchState.activeResultIndex,
          error: false,
          loading: searchState.loading,
          query: searchState.query,
          total: searchState.total,
        }
      : EMPTY_SEARCH,
    totalPages: safely(() => scroll?.getTotalPages()) ?? 0,
    zoomMode:
      typeof zoomState?.zoomLevel === 'string' ? zoomState.zoomLevel : null,
    zoomPercent: Math.round((zoomState?.currentZoomLevel ?? 1) * 100),
  };
}

function activeDocumentId(
  capabilities: PdfViewerCapabilities | null,
): string | null {
  return (
    safely(() => capabilities?.documentManager?.getActiveDocumentId()) ?? null
  );
}

function withActiveDocument(
  capabilities: PdfViewerCapabilities | null,
  action: (capabilities: PdfViewerCapabilities, documentId: string) => void,
): void {
  const documentId = activeDocumentId(capabilities);
  if (!capabilities || !documentId) return;
  safely(() => action(capabilities, documentId));
}

function subscribe<T>(
  unsubscribe: Array<() => void>,
  event: ((listener: (value: T) => void) => () => void) | undefined,
  listener: () => void,
): void {
  if (!event) return;
  const stop = safely(() => event(() => listener()));
  if (stop) unsubscribe.push(stop);
}

function safely<T>(read: () => T): T | undefined {
  try {
    return read();
  } catch {
    return undefined;
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error && error.message
    ? error.message
    : 'PDF viewer initialization failed.';
}
