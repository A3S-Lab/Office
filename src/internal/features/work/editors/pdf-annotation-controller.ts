import {
  AnnotationPlugin,
  type AnnotationCapability,
  DocumentManagerPlugin,
  type DocumentManagerCapability,
  type PluginRegistry,
} from '@embedpdf/react-pdf-viewer';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';

export type PdfAnnotationToolId =
  | 'highlight'
  | 'underline'
  | 'strikeout'
  | 'ink'
  | 'freeText';

export interface PdfAnnotationControllerState {
  activeToolId: string | null;
  annotationColor: string;
  available: boolean;
  hasPendingChanges: boolean;
  selectedCount: number;
}

export interface PdfAnnotationController {
  state: PdfAnnotationControllerState;
  deleteSelection: () => void;
  setAnnotationColor: (color: string) => void;
  selectTool: (toolId: PdfAnnotationToolId | null) => void;
}

interface PdfAnnotationCapabilities {
  annotation: AnnotationCapability | null;
  documentManager: DocumentManagerCapability | null;
}

const INITIAL_STATE: PdfAnnotationControllerState = {
  activeToolId: null,
  annotationColor: '#ffd966',
  available: false,
  hasPendingChanges: false,
  selectedCount: 0,
};

export function usePdfAnnotationController(
  registry: PluginRegistry | null,
): PdfAnnotationController {
  const capabilitiesRef = useRef<PdfAnnotationCapabilities | null>(null);
  const [state, setState] =
    useState<PdfAnnotationControllerState>(INITIAL_STATE);

  useEffect(() => {
    let disposed = false;
    const unsubscribe: Array<() => void> = [];
    capabilitiesRef.current = null;
    setState(INITIAL_STATE);
    if (!registry) return;

    void registry
      .pluginsReady()
      .then(() => {
        if (disposed) return;
        const capabilities: PdfAnnotationCapabilities = {
          annotation:
            registry
              .getPlugin<AnnotationPlugin>(AnnotationPlugin.id)
              ?.provides() ?? null,
          documentManager:
            registry
              .getPlugin<DocumentManagerPlugin>(DocumentManagerPlugin.id)
              ?.provides() ?? null,
        };
        capabilitiesRef.current = capabilities;

        const sync = () => {
          if (disposed) return;
          const scope = activeAnnotationScope(capabilities);
          const annotationState = safely(() => scope?.getState());
          const activeToolId = annotationState?.activeToolId ?? null;
          setState({
            activeToolId,
            annotationColor: currentAnnotationColor(
              capabilities,
              scope,
              activeToolId,
            ),
            available: Boolean(scope),
            hasPendingChanges: annotationState?.hasPendingChanges ?? false,
            selectedCount: annotationState?.selectedUids.length ?? 0,
          });
        };

        subscribe(unsubscribe, capabilities.annotation?.onStateChange, sync);
        subscribe(
          unsubscribe,
          capabilities.annotation?.onActiveToolChange,
          sync,
        );
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
        sync();
      })
      .catch(() => {
        if (!disposed) setState(INITIAL_STATE);
      });

    return () => {
      disposed = true;
      capabilitiesRef.current = null;
      for (const stop of unsubscribe) stop();
    };
  }, [registry]);

  const selectTool = useCallback((toolId: PdfAnnotationToolId | null) => {
    const capabilities = capabilitiesRef.current;
    const scope = activeAnnotationScope(capabilities);
    if (!scope) return;
    if (toolId && !capabilities?.annotation?.getTool(toolId)) {
      return;
    }
    safely(() => scope.setActiveTool(toolId));
  }, []);

  const deleteSelection = useCallback(() => {
    const scope = activeAnnotationScope(capabilitiesRef.current);
    if (!scope) return;
    const selected = new Set(scope.getSelectedAnnotationIds());
    if (selected.size === 0) return;
    const deletions = Object.entries(scope.getState().pages).flatMap(
      ([pageIndex, annotationIds]) =>
        annotationIds
          .filter((annotationId) => selected.has(annotationId))
          .map((annotationId) => ({
            pageIndex: Number(pageIndex),
            id: annotationId,
          })),
    );
    if (deletions.length > 0) scope.deleteAnnotations(deletions);
  }, []);

  const setAnnotationColor = useCallback((color: string) => {
    const normalizedColor = normalizeAnnotationColor(color);
    const capabilities = capabilitiesRef.current;
    const scope = activeAnnotationScope(capabilities);
    if (!normalizedColor || !capabilities?.annotation || !scope) return;
    const activeToolId = scope.getState().activeToolId;
    const patch = {
      color: normalizedColor,
      strokeColor: normalizedColor,
    };
    if (activeToolId) {
      capabilities.annotation.setToolDefaults(activeToolId, patch);
    }
    const selected = scope.getSelectedAnnotations();
    if (selected.length) {
      scope.updateAnnotations(
        selected.map(({ object }) => ({
          pageIndex: object.pageIndex,
          id: object.id,
          patch,
        })),
      );
    }
    setState((current) => ({
      ...current,
      annotationColor: normalizedColor,
    }));
  }, []);

  return useMemo(
    () => ({ state, deleteSelection, setAnnotationColor, selectTool }),
    [deleteSelection, setAnnotationColor, selectTool, state],
  );
}

function currentAnnotationColor(
  capabilities: PdfAnnotationCapabilities,
  scope: ReturnType<typeof activeAnnotationScope>,
  activeToolId: string | null,
): string {
  const selectedObject = safely(
    () => scope?.getSelectedAnnotations()[0]?.object,
  );
  const selectedColor = annotationColorFromRecord(selectedObject);
  if (selectedColor) return selectedColor;
  const activeTool = activeToolId
    ? safely(() => capabilities.annotation?.getTool(activeToolId))
    : null;
  return (
    annotationColorFromRecord(activeTool?.defaults) ??
    INITIAL_STATE.annotationColor
  );
}

function annotationColorFromRecord(value: unknown): string | null {
  if (!value || typeof value !== 'object') return null;
  const record = value as Record<string, unknown>;
  return normalizeAnnotationColor(record.color ?? record.strokeColor);
}

function normalizeAnnotationColor(value: unknown): string | null {
  if (typeof value !== 'string' || !/^#[0-9a-f]{6}$/i.test(value.trim())) {
    return null;
  }
  return value.trim().toLowerCase();
}

function activeAnnotationScope(capabilities: PdfAnnotationCapabilities | null) {
  const documentId =
    safely(() => capabilities?.documentManager?.getActiveDocumentId()) ?? null;
  if (!documentId) return null;
  return (
    safely(() => capabilities?.annotation?.forDocument(documentId)) ?? null
  );
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

function safely<T>(action: () => T): T | undefined {
  try {
    return action();
  } catch {
    return undefined;
  }
}
