import {
  PDFViewer,
  type PluginRegistry,
  type UISchema,
} from '@embedpdf/react-pdf-viewer';
import { AlertCircle, GalleryVerticalEnd, Loader2 } from 'lucide-react';
import {
  useCallback,
  useEffect,
  useId,
  useMemo,
  useRef,
  useState,
} from 'react';
import { Button, StateView } from '../../../design-system/primitives';
import { useDialogFocusScope } from '../../../design-system/primitives/overlay/dialog-focus-scope';
import { usePdfAnnotationController } from './pdf-annotation-controller';
import { createPdfEditorExtensions } from './pdf-editor-extensions';
import { PdfThumbnailRail } from './pdf-thumbnail-rail';
import { PdfToolbar, type PdfSaveState } from './pdf-toolbar';
import { usePdfViewerController } from './pdf-viewer-controller';
import { useOfficeEditorKeyboardShortcuts } from './use-office-editor-keyboard-shortcuts';
import { useOfficeEditorWheelZoom } from './use-office-editor-wheel-zoom';
import { useOfficeEditorRuntime } from './use-office-editor-runtime';

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
  fileName?: string;
  loadSource: () => Promise<Blob>;
  onSave?: (pdf: Blob) => Promise<boolean>;
  saveLabel?: string;
  sourceKey?: string;
  wasmUrl?: string;
}

export function PdfViewer({
  fileName = 'document.pdf',
  loadSource,
  onSave,
  saveLabel = '保存',
  sourceKey,
  wasmUrl,
}: PdfViewerProps) {
  const [sourceUrl, setSourceUrl] = useState<string | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [saveState, setSaveState] = useState<PdfSaveState>('idle');
  const [retryCount, setRetryCount] = useState(0);
  const [registry, setRegistry] = useState<PluginRegistry | null>(null);
  const [mobilePageNavigationOpen, setMobilePageNavigationOpen] =
    useState(false);
  const pdfRootRef = useRef<HTMLElement>(null);
  const searchInputRef = useRef<HTMLInputElement>(null);
  const mobilePageNavigationId = useId();
  const mobilePageNavigationToggleRef = useRef<HTMLButtonElement>(null);
  const mobilePageNavigationCloseRef = useRef<HTMLButtonElement>(null);
  const controller = usePdfViewerController(registry);
  const annotation = usePdfAnnotationController(registry);
  const viewerReady = controller.state.ready && controller.state.documentOpen;
  const mobilePageNavigationModal = usePdfMobilePageNavigationModal();
  const mobilePageNavigationModalOpen =
    mobilePageNavigationOpen && mobilePageNavigationModal;

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
    let objectUrl: string | null = null;
    setRegistry(null);
    setSaveState('idle');
    setSourceUrl(null);
    setLoadError(null);
    setMobilePageNavigationOpen(false);

    void loadSource()
      .then((source) => {
        if (disposed) return;
        objectUrl = URL.createObjectURL(
          source.type === 'application/pdf'
            ? source
            : new Blob([source], { type: 'application/pdf' }),
        );
        setSourceUrl(objectUrl);
      })
      .catch((error: unknown) => {
        if (!disposed) setLoadError(pdfErrorMessage(error));
      });

    return () => {
      disposed = true;
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [loadSource, retryCount, sourceKey]);

  useEffect(() => {
    if (controller.state.error) {
      setLoadError(controller.state.error);
    }
  }, [controller.state.error]);

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
      editable: Boolean(onSave),
      focusSearch: () => {
        searchInputRef.current?.focus();
        searchInputRef.current?.select();
      },
      save: {
        enabled: Boolean(onSave) && saveState !== 'saving',
        execute: savePdf,
      },
      viewer: controller,
    },
    pdfExtensions,
  );
  const pdfCommands = pdfEditor.commands;
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
        editable={Boolean(onSave)}
        searchInputRef={searchInputRef}
        saveLabel={saveLabel}
        saveState={saveState}
        state={controller.state}
      />
      <div
        className="work-pdf-workspace"
        data-mobile-page-navigation={
          mobilePageNavigationOpen ? 'open' : 'closed'
        }
      >
        {viewerReady && registry && controller.state.totalPages > 0 && (
          <>
            <button
              ref={mobilePageNavigationToggleRef}
              type="button"
              className="work-pdf-page-navigation-toggle"
              aria-label="打开 PDF 页面导航"
              aria-controls={mobilePageNavigationId}
              aria-expanded={mobilePageNavigationOpen}
              onClick={() => setMobilePageNavigationOpen(true)}
            >
              <GalleryVerticalEnd size={15} />
              <span>第 {Math.max(1, controller.state.currentPage)} 页</span>
            </button>
            <PdfThumbnailRail
              currentPage={controller.state.currentPage}
              mobileCloseButtonRef={mobilePageNavigationCloseRef}
              mobileNavigationId={mobilePageNavigationId}
              mobileNavigationModal={mobilePageNavigationModalOpen}
              registry={registry}
              totalPages={controller.state.totalPages}
              onCloseMobileNavigation={closeMobilePageNavigation}
              onSelectPage={(page) => {
                controller.goToPage(page);
                if (mobilePageNavigationModalOpen) {
                  closeMobilePageNavigation();
                }
              }}
            />
          </>
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
          aria-busy={!viewerReady}
          data-ready={viewerReady || undefined}
        >
          <PDFViewer
            key={sourceUrl}
            className="work-pdf-native-viewer"
            style={{ width: '100%', height: '100%' }}
            config={{
              src: sourceUrl,
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
              disabledCategories: onSave
                ? undefined
                : ['annotation', 'redaction', 'form', 'history'],
            }}
            onReady={setRegistry}
          />
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
