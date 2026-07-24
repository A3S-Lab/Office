import {
  PDFViewer,
  type PluginRegistry,
  type UISchema,
} from '@embedpdf/react-pdf-viewer';
import { AlertCircle, Loader2 } from 'lucide-react';
import { useCallback, useEffect, useRef, useState } from 'react';
import { Button, StateView } from '../../../design-system/primitives';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import { usePdfAnnotationController } from './pdf-annotation-controller';
import { PdfToolbar, type PdfSaveState } from './pdf-toolbar';
import { usePdfViewerController } from './pdf-viewer-controller';

const PDFIUM_WASM_PATH = '/vendor/embedpdf/pdfium.wasm';
const PDF_VIEWER_READY_TIMEOUT_MS = 20_000;

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
  const searchInputRef = useRef<HTMLInputElement>(null);
  const controller = usePdfViewerController(registry);
  const annotation = usePdfAnnotationController(registry);
  const viewerReady = controller.state.ready && controller.state.documentOpen;

  useEffect(() => {
    let disposed = false;
    let objectUrl: string | null = null;
    setRegistry(null);
    setSaveState('idle');
    setSourceUrl(null);
    setLoadError(null);

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

  useEffect(() => {
    if (!sourceUrl) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.defaultPrevented || event.repeat || event.altKey) {
        return;
      }

      const key = event.key.toLowerCase();
      if (!(event.metaKey || event.ctrlKey)) {
        if (
          onSave &&
          !isOfficeShortcutBlocked(event.target) &&
          (key === 'delete' || key === 'backspace') &&
          annotation.state.selectedCount > 0
        ) {
          event.preventDefault();
          event.stopImmediatePropagation();
          annotation.deleteSelection();
        } else if (
          key === 'escape' &&
          !isOfficeShortcutBlocked(event.target) &&
          annotation.state.activeToolId
        ) {
          event.preventDefault();
          annotation.selectTool(null);
        }
        return;
      }

      if (key === 'f') {
        event.preventDefault();
        event.stopImmediatePropagation();
        searchInputRef.current?.focus();
        searchInputRef.current?.select();
        return;
      }
      if (isOfficeShortcutBlocked(event.target)) return;

      let handled = true;
      if (key === 's' && onSave) {
        void savePdf();
      } else if (key === 'z' && event.shiftKey) {
        controller.redo();
      } else if (key === 'z') {
        controller.undo();
      } else if (key === 'y') {
        controller.redo();
      } else if (key === '+' || key === '=') {
        controller.zoomIn();
      } else if (key === '-') {
        controller.zoomOut();
      } else if (key === '0') {
        controller.fitPage();
      } else {
        handled = false;
      }

      if (handled) {
        event.preventDefault();
        event.stopImmediatePropagation();
      }
    };
    window.addEventListener('keydown', onKeyDown, { capture: true });
    return () =>
      window.removeEventListener('keydown', onKeyDown, { capture: true });
  }, [annotation, controller, onSave, savePdf, sourceUrl]);

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
    <section className="work-pdf-viewer" aria-label={`PDF 编辑器：${fileName}`}>
      <PdfToolbar
        annotation={annotation}
        controller={controller}
        editable={Boolean(onSave)}
        searchInputRef={searchInputRef}
        saveLabel={saveLabel}
        saveState={saveState}
        onSave={onSave ? () => void savePdf() : undefined}
      />
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
    </section>
  );
}

function pdfErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message) return error.message;
  return 'Unable to read this PDF file.';
}
