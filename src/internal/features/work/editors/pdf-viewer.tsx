import {
  PDFViewer,
  type PluginRegistry,
  type UISchema,
} from '@embedpdf/react-pdf-viewer';
import { AlertCircle, Loader2 } from 'lucide-react';
import {
  useCallback,
  useEffect,
  useId,
  useMemo,
  useRef,
  useState,
} from 'react';
import type { WorkOfficeCollaborationSession } from '../../../collaboration/office-collaboration';
import type { WorkOfficeCollaborationParticipant } from '../../../collaboration/office-collaboration-presence';
import {
  assertWorkOfficePdfCollaborationSource,
  readWorkOfficePdfCollaborationSource,
} from '../../../collaboration/office-pdf-collaboration';
import type { WorkPdfCollaborationContent } from '../../../collaboration/office-pdf-collaboration-types';
import { Button, StateView } from '../../../design-system/primitives';
import { useDialogFocusScope } from '../../../design-system/primitives/overlay/dialog-focus-scope';
import { usePdfAnnotationController } from './pdf-annotation-controller';
import { createWorkPdfCollaborationProjection } from './pdf-collaboration-projection';
import { createPdfEditorExtensions } from './pdf-editor-extensions';
import type {
  PdfEvidenceOverlay,
  PdfEvidenceRegion,
} from './pdf-evidence-contract';
import { useOfficeCollaborationLocationNavigator } from './office-collaboration-presence-context';
import { useOfficePublishPresenceLocation } from './office-collaboration-presence-ui';
import { PdfCollaborationPresenceLayer } from './pdf-collaboration-presence';
import { PdfEvidenceOverlayLayer } from './pdf-evidence-overlay';
import { PdfPageOrganizerDialog } from './pdf-page-organizer-dialog';
import { PdfThumbnailRail } from './pdf-thumbnail-rail';
import { type PdfSaveState, PdfToolbar } from './pdf-toolbar';
import { usePdfViewerController } from './pdf-viewer-controller';
import { useOfficeEditorKeyboardShortcuts } from './use-office-editor-keyboard-shortcuts';
import { useOfficeEditorRuntime } from './use-office-editor-runtime';
import { useOfficeEditorWheelZoom } from './use-office-editor-wheel-zoom';
import {
  usePdfPageOrganization,
  type PdfPageOrganizationExport,
} from './use-pdf-page-organization';

const PDFIUM_WASM_PATH = '/vendor/embedpdf/pdfium.wasm';
// PDFium can take longer on its first WASM startup, especially after a fresh
// Playground build. Keep the loading state instead of surfacing a false error
// while the worker is still making progress.
const PDF_VIEWER_READY_TIMEOUT_MS = 45_000;
const PDF_MOBILE_PAGE_NAVIGATION_QUERY = '(max-width: 640px)';

export const a3sPdfUiSchema: UISchema = {
  id: 'a3s-office-pdf',
  version: '1',
  toolbars: {},
  menus: {},
  sidebars: {},
  modals: {},
  overlays: {},
  selectionMenus: {},
};

export interface PdfViewerProps {
  collaboration?: WorkOfficeCollaborationSession;
  evidenceOverlay?: PdfEvidenceOverlay;
  fileName?: string;
  loadSource: () => Promise<Blob>;
  onCollaborationChange?: (content: WorkPdfCollaborationContent) => void;
  onEvidenceRegionSelect?: (region: PdfEvidenceRegion) => void;
  onPageChange?: (pageNumber: number) => void;
  onPageExport?: (
    files: readonly PdfPageOrganizationExport[],
  ) => boolean | Promise<boolean>;
  onSave?: (pdf: Blob) => Promise<boolean>;
  saveLabel?: string;
  selectedEvidenceRegionId?: string;
  sourceKey?: string;
  wasmUrl?: string;
  worker?: boolean;
}

export function PdfViewer({
  collaboration,
  evidenceOverlay,
  fileName = 'document.pdf',
  loadSource,
  onCollaborationChange,
  onEvidenceRegionSelect,
  onPageChange,
  onPageExport,
  onSave,
  saveLabel = '保存',
  selectedEvidenceRegionId,
  sourceKey,
  wasmUrl,
  worker = true,
}: PdfViewerProps) {
  const [sourceBlob, setSourceBlob] = useState<Blob | null>(null);
  const [sourceUrl, setSourceUrl] = useState<string | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [saveState, setSaveState] = useState<PdfSaveState>('idle');
  const [retryCount, setRetryCount] = useState(0);
  const [registry, setRegistry] = useState<PluginRegistry | null>(null);
  const [, refreshCollaborationHistory] = useState(0);
  const [hostSourceGeneration, setHostSourceGeneration] = useState(0);
  const [pageOrganizerOpen, setPageOrganizerOpen] = useState(false);
  const [mobilePageNavigationOpen, setMobilePageNavigationOpen] =
    useState(false);
  const pdfRootRef = useRef<HTMLElement>(null);
  const searchInputRef = useRef<HTMLInputElement>(null);
  const mobilePageNavigationId = useId();
  const mobilePageNavigationToggleRef = useRef<HTMLButtonElement>(null);
  const mobilePageNavigationCloseRef = useRef<HTMLButtonElement>(null);
  const collaborationProjectionRef = useRef<
    ReturnType<typeof createWorkPdfCollaborationProjection> | undefined
  >(undefined);
  const editable = collaboration
    ? collaboration.mode === 'edit'
    : Boolean(onSave);
  const pageOrganizationEnabled =
    Boolean(onSave) && !collaboration && !evidenceOverlay;
  const collaborationHistory = collaboration
    ? {
        canRedo: collaborationProjectionRef.current?.binding.canRedo() ?? false,
        canUndo: collaborationProjectionRef.current?.binding.canUndo() ?? false,
        redo: () => collaborationProjectionRef.current?.binding.redo(),
        undo: () => collaborationProjectionRef.current?.binding.undo(),
      }
    : undefined;
  const controller = usePdfViewerController(registry, collaborationHistory);
  const replacePageSource = useCallback((source: Blob) => {
    setRegistry(null);
    setSourceUrl(null);
    setSourceBlob(
      source.type === 'application/pdf'
        ? source
        : new Blob([source], { type: 'application/pdf' }),
    );
    setLoadError(null);
    setSaveState('idle');
    setMobilePageNavigationOpen(false);
    setPageOrganizerOpen(false);
  }, []);
  const pageOrganization = usePdfPageOrganization({
    enabled: pageOrganizationEnabled,
    fileName,
    onExport: onPageExport,
    readCurrentSource: () => controller.saveAsCopy(),
    replaceSource: replacePageSource,
    resetKey: `${sourceKey ?? 'source'}:${retryCount}:${hostSourceGeneration}`,
  });
  const viewerController = useMemo(
    () => ({
      ...controller,
      state: {
        ...controller.state,
        canRedo: controller.state.canRedo || pageOrganization.state.canRedo,
        canUndo: controller.state.canUndo || pageOrganization.state.canUndo,
        features: {
          ...controller.state.features,
          history:
            controller.state.features.history ||
            pageOrganization.state.available,
        },
      },
      redo: () => {
        if (controller.state.canRedo) controller.redo();
        else pageOrganization.redo();
      },
      undo: () => {
        if (controller.state.canUndo) controller.undo();
        else pageOrganization.undo();
      },
    }),
    [controller, pageOrganization],
  );
  const annotation = usePdfAnnotationController(registry);
  const viewerReady =
    viewerController.state.ready && viewerController.state.documentOpen;
  const mobilePageNavigationModal = usePdfMobilePageNavigationModal();
  const mobilePageNavigationModalOpen =
    mobilePageNavigationOpen && mobilePageNavigationModal;
  const collaborationRef = useRef(collaboration);
  if (collaborationRef.current !== collaboration) {
    throw new Error(
      'PdfViewer collaboration sessions cannot be replaced while mounted. Remount the viewer for another shared PDF.',
    );
  }

  const closeMobilePageNavigation = useCallback(() => {
    setMobilePageNavigationOpen(false);
  }, []);
  useDialogFocusScope<HTMLElement>({
    active: mobilePageNavigationModalOpen,
    onEscape: closeMobilePageNavigation,
    initialFocus: () => mobilePageNavigationCloseRef.current,
    getActiveScope: () =>
      document.getElementById(mobilePageNavigationId) as HTMLElement | null,
    getIsolationExceptions: () => [
      document.querySelector<HTMLElement>('.work-pdf-page-navigation-backdrop'),
    ],
    restoreFocusTarget: () =>
      mobilePageNavigationModal
        ? mobilePageNavigationToggleRef.current
        : (document
            .getElementById(mobilePageNavigationId)
            ?.querySelector<HTMLElement>(
              '[data-pdf-page-thumbnail][aria-current="page"]',
            ) ?? null),
  });

  useEffect(() => {
    let disposed = false;
    setRegistry(null);
    setSaveState('idle');
    setSourceBlob(null);
    setSourceUrl(null);
    setLoadError(null);
    setMobilePageNavigationOpen(false);
    setPageOrganizerOpen(false);

    void loadSource()
      .then(async (source) => {
        if (disposed) return;
        if (collaboration) {
          await assertPdfCollaborationBlob(collaboration, source);
        }
        if (disposed) return;
        setSourceBlob(
          source.type === 'application/pdf'
            ? source
            : new Blob([source], { type: 'application/pdf' }),
        );
        setHostSourceGeneration((value) => value + 1);
      })
      .catch((error: unknown) => {
        if (!disposed) setLoadError(pdfErrorMessage(error));
      });

    return () => {
      disposed = true;
    };
  }, [collaboration, loadSource, retryCount, sourceKey]);

  useEffect(() => {
    if (!sourceBlob) {
      setSourceUrl(null);
      return;
    }
    const objectUrl = URL.createObjectURL(sourceBlob);
    setSourceUrl(objectUrl);
    return () => URL.revokeObjectURL(objectUrl);
  }, [sourceBlob]);

  useEffect(() => {
    if (controller.state.error) {
      setLoadError(controller.state.error);
    }
  }, [controller.state.error]);

  useEffect(() => {
    if (!collaboration || !viewerReady) return;
    if (controller.state.totalPages < 1) return;
    const expectedPages = readWorkPdfSourcePageCount(collaboration);
    if (controller.state.totalPages !== expectedPages) {
      setRegistry(null);
      setLoadError(
        `The loaded PDF has ${controller.state.totalPages} pages, but collaboration artifact '${collaboration.artifactId}' requires ${expectedPages}.`,
      );
    }
  }, [collaboration, controller.state.totalPages, viewerReady]);

  useEffect(() => {
    if (!registry || !collaboration || !viewerReady) return;
    if (controller.state.totalPages < 1) return;
    const expectedPages = readWorkPdfSourcePageCount(collaboration);
    if (controller.state.totalPages !== expectedPages) return;
    const projection = createWorkPdfCollaborationProjection(
      registry,
      collaboration,
    );
    collaborationProjectionRef.current = projection;
    const handleProjectionError = (error: unknown) => {
      if (collaborationProjectionRef.current !== projection) return;
      collaborationProjectionRef.current = undefined;
      setRegistry(null);
      setLoadError(pdfErrorMessage(error));
    };
    const unsubscribeError = projection.subscribeError(handleProjectionError);
    void projection.ready.catch(handleProjectionError);
    const unsubscribeHistory = projection.binding.subscribeHistory(() =>
      refreshCollaborationHistory((value) => value + 1),
    );
    refreshCollaborationHistory((value) => value + 1);
    return () => {
      unsubscribeError();
      unsubscribeHistory();
      if (collaborationProjectionRef.current === projection) {
        collaborationProjectionRef.current = undefined;
      }
      projection.destroy();
    };
  }, [collaboration, controller.state.totalPages, registry, viewerReady]);

  useEffect(() => {
    if (!collaboration || !onCollaborationChange) return;
    const handleUpdate = () => {
      const projection = collaborationProjectionRef.current;
      if (!projection) return;
      queueMicrotask(() => {
        if (collaborationProjectionRef.current === projection) {
          onCollaborationChange(projection.binding.content());
        }
      });
    };
    collaboration.document.on('update', handleUpdate);
    return () => collaboration.document.off('update', handleUpdate);
  }, [collaboration, onCollaborationChange]);

  useEffect(() => {
    if (viewerController.state.currentPage > 0) {
      onPageChange?.(viewerController.state.currentPage);
    }
  }, [onPageChange, viewerController.state.currentPage]);

  useEffect(() => {
    if (!sourceUrl || viewerReady || loadError) return;
    const timeout = window.setTimeout(() => {
      setRegistry(null);
      setLoadError('PDF viewer initialization timed out.');
    }, PDF_VIEWER_READY_TIMEOUT_MS);
    return () => window.clearTimeout(timeout);
  }, [loadError, sourceUrl, viewerReady]);

  const savePdf = useCallback(async () => {
    if (!onSave || saveState === 'saving') return;
    setSaveState('saving');
    try {
      const saved = await onSave(await controller.saveAsCopy());
      setSaveState(saved ? 'saved' : 'error');
    } catch {
      setSaveState('error');
    }
  }, [controller, onSave, saveState]);
  const pdfExtensions = useMemo(createPdfEditorExtensions, []);
  const pdfEditor = useOfficeEditorRuntime(
    {
      annotation,
      editable,
      focusSearch: () => {
        searchInputRef.current?.focus();
        searchInputRef.current?.select();
      },
      openPageOrganizer: () => setPageOrganizerOpen(true),
      pages: pageOrganization,
      save: {
        enabled: Boolean(onSave) && saveState !== 'saving',
        execute: savePdf,
      },
      viewer: viewerController,
    },
    pdfExtensions,
  );
  const pdfCommands = pdfEditor.commands;
  const selectedAnnotationLocation = annotation.getSelectionLocation();
  const pdfPresenceLocation =
    viewerReady && viewerController.state.currentPage > 0
      ? {
          kind: 'pdf' as const,
          pageIndex:
            selectedAnnotationLocation?.pageIndex ??
            viewerController.state.currentPage - 1,
          ...(selectedAnnotationLocation
            ? { annotationId: selectedAnnotationLocation.annotationId }
            : {}),
        }
      : null;
  useOfficePublishPresenceLocation(pdfPresenceLocation);
  const navigateToPdfParticipant = useCallback(
    (participant: WorkOfficeCollaborationParticipant): boolean => {
      const location = participant.location;
      if (
        location?.kind !== 'pdf' ||
        !viewerReady ||
        location.pageIndex >= viewerController.state.totalPages
      ) {
        return false;
      }
      viewerController.goToPage(location.pageIndex + 1);
      if (location.annotationId) {
        annotation.locateAnnotation(location.pageIndex, location.annotationId);
      }
      requestAnimationFrame(() => {
        const target =
          pdfRootRef.current?.querySelector<HTMLElement>('.work-pdf-embed');
        target?.focus({ preventScroll: true });
      });
      return true;
    },
    [annotation, viewerController, viewerReady],
  );
  useOfficeCollaborationLocationNavigator(navigateToPdfParticipant);
  useOfficeEditorKeyboardShortcuts(pdfEditor, {
    capture: true,
    enabled: Boolean(sourceUrl),
    scopeRef: pdfRootRef,
  });
  useOfficeEditorWheelZoom({
    enabled: Boolean(sourceUrl),
    scopeRef: pdfRootRef,
    onZoomIn: pdfCommands.zoomIn,
    onZoomOut: pdfCommands.zoomOut,
  });

  if (loadError) {
    return (
      <StateView
        className="work-pdf-state"
        tone="danger"
        role="alert"
        icon={<AlertCircle size={24} />}
        title="无法打开 PDF"
        description="请重试。"
        descriptionTitle={loadError}
        actions={
          <Button onClick={() => setRetryCount((value) => value + 1)}>
            重试
          </Button>
        }
      />
    );
  }

  if (!sourceUrl) {
    return (
      <StateView
        className="work-pdf-state"
        role="status"
        icon={<Loader2 className="spin" size={22} />}
        title="正在加载 PDF…"
      />
    );
  }

  return (
    <section
      ref={pdfRootRef}
      className="work-pdf-viewer"
      aria-label={`PDF 编辑器：${fileName}`}
    >
      <PdfToolbar
        annotationState={annotation.state}
        can={pdfEditor.can()}
        commands={pdfCommands}
        editable={editable}
        pageOrganizationAvailable={pageOrganization.state.available}
        saveAvailable={editable && Boolean(onSave)}
        pageNavigation={
          viewerReady && registry && viewerController.state.totalPages > 0
            ? {
                controlsId: mobilePageNavigationId,
                expanded: mobilePageNavigationOpen,
                onOpen: () => setMobilePageNavigationOpen(true),
                toggleRef: mobilePageNavigationToggleRef,
              }
            : undefined
        }
        searchInputRef={searchInputRef}
        saveLabel={saveLabel}
        saveState={saveState}
        state={viewerController.state}
      />
      {pageOrganizerOpen &&
        viewerReady &&
        registry &&
        viewerController.state.totalPages > 0 && (
          <PdfPageOrganizerDialog
            busy={pageOrganization.state.busy}
            can={pdfEditor.can()}
            commands={pdfCommands}
            currentPage={viewerController.state.currentPage}
            diagnostics={pageOrganization.state.diagnostics}
            error={pageOrganization.state.error}
            registry={registry}
            restoreFocusTarget={() =>
              pdfPageOrganizerRestoreFocusTarget(pdfRootRef.current)
            }
            totalPages={viewerController.state.totalPages}
            onClose={() => setPageOrganizerOpen(false)}
            onDismissError={pageOrganization.dismissError}
          />
        )}
      <div
        className="work-pdf-workspace"
        data-mobile-page-navigation={
          mobilePageNavigationOpen ? 'open' : 'closed'
        }
      >
        {viewerReady && registry && viewerController.state.totalPages > 0 && (
          <PdfThumbnailRail
            currentPage={viewerController.state.currentPage}
            mobileCloseButtonRef={mobilePageNavigationCloseRef}
            mobileNavigationId={mobilePageNavigationId}
            mobileNavigationModal={mobilePageNavigationModalOpen}
            registry={registry}
            totalPages={viewerController.state.totalPages}
            onCloseMobileNavigation={closeMobilePageNavigation}
            onSelectPage={(page) => {
              viewerController.goToPage(page);
              if (mobilePageNavigationModalOpen) {
                closeMobilePageNavigation();
              }
            }}
          />
        )}
        {mobilePageNavigationModalOpen && (
          <button
            type="button"
            className="work-pdf-page-navigation-backdrop"
            aria-label="关闭 PDF 页面导航遮罩"
            tabIndex={-1}
            onClick={closeMobilePageNavigation}
          />
        )}
        <div
          className="work-pdf-embed"
          role="application"
          aria-label="PDF 页面画布"
          aria-busy={!viewerReady}
          data-ready={viewerReady || undefined}
          // biome-ignore lint/a11y/noNoninteractiveTabindex: EmbedPDF is a composite canvas application and this wrapper is its stable navigation focus target.
          tabIndex={0}
        >
          <PDFViewer
            key={sourceUrl}
            className="work-pdf-native-viewer"
            style={{ width: '100%', height: '100%' }}
            config={{
              src: sourceUrl,
              worker,
              // EmbedPDF creates a Blob worker, so a root-relative URL has no
              // usable base inside WorkerGlobalScope. Keep this absolute.
              wasmUrl:
                wasmUrl ?? new URL(PDFIUM_WASM_PATH, window.location.href).href,
              tabBar: 'never',
              theme: {
                preference: 'system',
                light: { accent: { primary: '#2867d8' } },
                dark: { accent: { primary: '#7da7ff' } },
              },
              i18n: { defaultLocale: 'zh-CN' },
              ui: { schema: a3sPdfUiSchema },
              thumbnails: {
                autoScroll: false,
                buffer: 4,
                gap: 10,
                imagePadding: 0,
                labelHeight: 0,
                paddingY: 0,
                width: 132,
              },
              annotations: {
                annotationAuthor: 'A3S Office 用户',
                autoCommit: true,
              },
              export: { defaultFileName: fileName },
              fonts: { ui: null, signature: null },
              disabledCategories: editable
                ? undefined
                : ['annotation', 'redaction', 'form', 'history'],
            }}
            onReady={setRegistry}
          />
          {viewerReady && viewerController.state.currentPage > 0 && (
            <PdfCollaborationPresenceLayer
              pageIndex={viewerController.state.currentPage - 1}
            />
          )}
          {evidenceOverlay && registry && viewerReady && (
            <PdfEvidenceOverlayLayer
              currentPage={viewerController.state.currentPage}
              evidenceOverlay={evidenceOverlay}
              onEvidenceRegionSelect={onEvidenceRegionSelect}
              registry={registry}
              selectedEvidenceRegionId={selectedEvidenceRegionId}
              sourceKey={sourceKey}
            />
          )}
          {!viewerReady && (
            <div className="work-pdf-loading" role="status">
              <Loader2 className="spin" size={18} />
              正在打开…
            </div>
          )}
        </div>
      </div>
    </section>
  );
}

function pdfErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message) return error.message;
  return 'Unable to read this PDF file.';
}

function pdfPageOrganizerRestoreFocusTarget(
  root: HTMLElement | null,
): HTMLElement | null {
  const directTrigger = root?.querySelector<HTMLElement>(
    '[data-pdf-page-organizer-trigger]',
  );
  if (directTrigger && directTrigger.getClientRects().length > 0) {
    return directTrigger;
  }
  return root?.querySelector<HTMLElement>('.work-pdf-overflow-trigger') ?? null;
}

async function assertPdfCollaborationBlob(
  session: WorkOfficeCollaborationSession,
  source: Blob,
): Promise<void> {
  const expected = readWorkOfficePdfCollaborationSource(session);
  if (source.size !== expected.byteLength) {
    throw new Error(
      `The loaded PDF source does not match collaboration artifact '${session.artifactId}'.`,
    );
  }
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) {
    throw new Error('SHA-256 PDF source verification is unavailable.');
  }
  const digest = new Uint8Array(
    await subtle.digest('SHA-256', await source.arrayBuffer()),
  );
  assertWorkOfficePdfCollaborationSource(session, {
    ...expected,
    sha256: Array.from(digest, (value) =>
      value.toString(16).padStart(2, '0'),
    ).join(''),
  });
}

function readWorkPdfSourcePageCount(
  session: WorkOfficeCollaborationSession,
): number {
  return readWorkOfficePdfCollaborationSource(session).pageCount;
}

function usePdfMobilePageNavigationModal(): boolean {
  const [matches, setMatches] = useState(() =>
    pdfMediaQueryMatches(PDF_MOBILE_PAGE_NAVIGATION_QUERY),
  );

  useEffect(() => {
    if (
      typeof window === 'undefined' ||
      typeof window.matchMedia !== 'function'
    ) {
      return;
    }
    const mediaQuery = window.matchMedia(PDF_MOBILE_PAGE_NAVIGATION_QUERY);
    const update = () => setMatches(mediaQuery.matches);
    update();
    if (typeof mediaQuery.addEventListener === 'function') {
      mediaQuery.addEventListener('change', update);
      return () => mediaQuery.removeEventListener('change', update);
    }
    mediaQuery.addListener(update);
    return () => mediaQuery.removeListener(update);
  }, []);

  return matches;
}

function pdfMediaQueryMatches(query: string): boolean {
  if (
    typeof window === 'undefined' ||
    typeof window.matchMedia !== 'function'
  ) {
    return true;
  }
  return window.matchMedia(query).matches;
}
