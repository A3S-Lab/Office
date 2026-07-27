import type { PluginRegistry } from '@embedpdf/react-pdf-viewer';
import { expect, test } from '@rstest/core';
import { act, renderHook, waitFor } from '@testing-library/react';
import { usePdfAnnotationController } from '../src/internal/features/work/editors/pdf-annotation-controller';

test('selects and deletes PDF annotations through the annotation capability', async () => {
  const calls: string[] = [];
  const annotationChange = createEvent();
  const activeToolChange = createEvent();
  const documentChange = createEvent();
  const documentOpened = createEvent();
  const documentClosed = createEvent();
  const annotationState = {
    pages: { 0: ['annotation-1'], 2: ['annotation-2'] },
    byUid: {},
    selectedUids: ['annotation-2'],
    selectedUid: 'annotation-2',
    activeToolId: 'highlight',
    hasPendingChanges: true,
    locked: { type: 'none' },
  };
  const selectedAnnotation = {
    commitState: 'dirty',
    object: {
      id: 'annotation-2',
      pageIndex: 2,
      type: 9,
      rect: { origin: { x: 0, y: 0 }, size: { width: 10, height: 10 } },
      color: '#ffd966',
      strokeColor: '#ffd966',
      opacity: 0.75,
      strokeWidth: 6,
    },
  };
  const capabilities = {
    annotation: {
      forDocument: () => ({
        getState: () => annotationState,
        getSelectedAnnotations: () => [selectedAnnotation],
        getSelectedAnnotationIds: () => annotationState.selectedUids,
        setActiveTool: (toolId: string | null) =>
          calls.push(`tool:${toolId ?? 'pointer'}`),
        updateAnnotations: (
          annotations: Array<{
            pageIndex: number;
            id: string;
            patch: {
              color?: string;
              opacity?: number;
              strokeColor?: string;
              strokeWidth?: number;
            };
          }>,
        ) => {
          const patch = annotations[0]?.patch;
          const kind =
            patch && 'opacity' in patch
              ? 'opacity'
              : patch && 'strokeWidth' in patch
                ? 'stroke-width'
                : 'color';
          calls.push(`${kind}:${JSON.stringify(annotations)}`);
        },
        deleteAnnotations: (
          annotations: Array<{ pageIndex: number; id: string }>,
        ) => calls.push(`delete:${JSON.stringify(annotations)}`),
      }),
      getTool: (toolId: string) =>
        ['highlight', 'underline', 'strikeout', 'ink', 'freeText'].includes(
          toolId,
        )
          ? {
              id: toolId,
              defaults: {
                color: '#ffd966',
                strokeColor: '#ffd966',
                opacity: 1,
                ...(toolId === 'ink' ? { strokeWidth: 6 } : {}),
              },
            }
          : undefined,
      setToolDefaults: (
        toolId: string,
        patch: {
          color?: string;
          opacity?: number;
          strokeColor?: string;
          strokeWidth?: number;
        },
      ) => calls.push(`defaults:${toolId}:${JSON.stringify(patch)}`),
      onStateChange: annotationChange.subscribe,
      onActiveToolChange: activeToolChange.subscribe,
    },
    'document-manager': {
      getActiveDocumentId: () => 'document-1',
      onActiveDocumentChanged: documentChange.subscribe,
      onDocumentOpened: documentOpened.subscribe,
      onDocumentClosed: documentClosed.subscribe,
    },
  };
  const registry = {
    pluginsReady: () => Promise.resolve(),
    getPlugin: (id: keyof typeof capabilities) => ({
      provides: () => capabilities[id],
    }),
  } as unknown as PluginRegistry;

  const { result } = renderHook(() => usePdfAnnotationController(registry));
  await waitFor(() => expect(result.current.state.available).toBe(true));
  expect(result.current.state).toMatchObject({
    activeToolId: 'highlight',
    annotationColor: '#ffd966',
    annotationOpacity: 0.75,
    annotationStrokeWidth: 6,
    hasPendingChanges: true,
    selectedCount: 1,
    supportsOpacity: true,
    supportsStrokeWidth: true,
  });

  act(() => {
    result.current.selectTool('ink');
    result.current.setAnnotationColor('#ff0000');
    result.current.setAnnotationOpacity(0.5);
    result.current.setAnnotationStrokeWidth(4);
    result.current.selectTool(null);
    result.current.deleteSelection();
  });

  expect(calls).toEqual([
    'tool:ink',
    'defaults:highlight:{"color":"#ff0000","strokeColor":"#ff0000"}',
    'color:[{"pageIndex":2,"id":"annotation-2","patch":{"color":"#ff0000","strokeColor":"#ff0000"}}]',
    'defaults:highlight:{"opacity":0.5}',
    'opacity:[{"pageIndex":2,"id":"annotation-2","patch":{"opacity":0.5}}]',
    'stroke-width:[{"pageIndex":2,"id":"annotation-2","patch":{"strokeWidth":4}}]',
    'tool:pointer',
    'delete:[{"pageIndex":2,"id":"annotation-2"}]',
  ]);
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
